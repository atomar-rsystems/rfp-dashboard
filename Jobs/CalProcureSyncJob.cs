using System.Data;
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Hangfire;
using Hangfire.Console;
using Hangfire.Server;
using Microsoft.Playwright;
using Npgsql;

namespace RfpDashboard.Jobs;

public sealed class CalProcureSyncJob
{
    private const string SourceUrl = "https://caleprocure.ca.gov/pages/Events-BS3/event-search.aspx";
    private readonly NpgsqlConnection _connection;
    private readonly ILogger<CalProcureSyncJob> _logger;

    public CalProcureSyncJob(NpgsqlConnection connection, ILogger<CalProcureSyncJob> logger)
    {
        _connection = connection;
        _logger = logger;
    }

    public async Task ExecuteAsync(PerformContext? context)
    {
        LogProgress(context, "Starting California eProcure sync job.");

        var downloadPath = await DownloadAttachmentAsync(context);
        LogProgress(context, "Downloaded source file: {0}", downloadPath);

        var records = ParseDownloadedFile(downloadPath);
        LogProgress(context, "Parsed {0} records from download.", records.Count);

        if (records.Count == 0)
        {
            LogProgress(context, "No California eProcure records were found in downloaded file.");
            return;
        }

        if (_connection.State != ConnectionState.Open)
        {
            await _connection.OpenAsync();
        }

        var inserted = 0;

        foreach (var record in records)
        {
            await using var command = _connection.CreateCommand();
            command.CommandText = """
                INSERT INTO public.t_cal_eprocure
                (department_code, department_name, event_id, event_name, format, event_type, end_date, status, buyer_name, buyer_email, triage_status)
                VALUES
                (@departmentCode, @departmentName, @eventId, @eventName, @format, @eventType, @endDate, @status, @buyerName, @buyerEmail, 'PENDING')
                ON CONFLICT (department_code, event_id) DO NOTHING;
                """;

            command.Parameters.AddWithValue("departmentCode", record.DepartmentCode);
            command.Parameters.AddWithValue("departmentName", record.DepartmentName);
            command.Parameters.AddWithValue("eventId", record.EventId);
            command.Parameters.AddWithValue("eventName", record.EventName);
            command.Parameters.AddWithValue("format", record.Format);
            command.Parameters.AddWithValue("eventType", record.EventType);
            command.Parameters.AddWithValue("endDate", record.EndDate.ToUniversalTime());
            command.Parameters.AddWithValue("status", record.Status);
            command.Parameters.AddWithValue("buyerName", (object?)record.BuyerName ?? DBNull.Value);
            command.Parameters.AddWithValue("buyerEmail", (object?)record.BuyerEmail ?? DBNull.Value);

            inserted += await command.ExecuteNonQueryAsync();
        }

        LogProgress(context, "California eProcure sync completed. Parsed: {0}, inserted: {1}", records.Count, inserted);

        if (inserted > 0)
        {
            BackgroundJob.Enqueue<CalProcureClassifyJob>(job => job.ExecuteAsync(null, CancellationToken.None));
            LogProgress(context, "Enqueued CalProcureClassifyJob to evaluate {0} new events.", inserted);
        }
    }

    private void LogProgress(PerformContext? context, string format, params object?[] args)
    {
        var formatted = args.Length > 0 ? string.Format(format, args) : format;
        _logger.LogInformation("{Message}", formatted);

        if (context is not null)
        {
            context.WriteLine(formatted);
        }
    }

    private async Task<string> DownloadAttachmentAsync(PerformContext? context)
    {
        using var playwright = await Playwright.CreateAsync();
        await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
        {
            Channel = "msedge",
            Headless = true,
            Args = ["--disable-blink-features=AutomationControlled"]
        });

        await using var browserContext = await browser.NewContextAsync(new BrowserNewContextOptions
        {
            AcceptDownloads = true,
            ViewportSize = new ViewportSize { Width = 1920, Height = 1080 },
            UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0"
        });

        browserContext.SetDefaultTimeout(120000);

        var page = await browserContext.NewPageAsync();
        await page.GotoAsync(SourceUrl, new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded });

        await page.WaitForTimeoutAsync(3000);

        try
        {
            LogProgress(context, "Attempting to trigger search on California eProcure...");
            await ClickSearchAsync(page);
            await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
        }
        catch (TimeoutException)
        {
            LogProgress(context, "Search button was not clicked (results may already be loaded). Proceeding to locate export button.");
        }

        var download = await ClickExportAndDownloadAsync(browserContext, page, context);

        var localFile = Path.Combine(Path.GetTempPath(), $"cal-eprocure-{DateTime.UtcNow:yyyyMMddHHmmss}-{download.SuggestedFilename}");
        await download.SaveAsAsync(localFile);
        return localFile;
    }

    private async Task<IDownload> ClickExportAndDownloadAsync(IBrowserContext context, IPage page, PerformContext? contextLog)
    {
        var exportButton = await FindVisibleExportButtonAsync(page, 60000);
        if (exportButton is null)
        {
            var debugPath = Path.Combine(Path.GetTempPath(), $"caleprocure-timeout-{DateTime.UtcNow:yyyyMMddHHmmss}.png");
            await page.ScreenshotAsync(new PageScreenshotOptions { Path = debugPath, FullPage = true });
            throw new TimeoutException($"Download button was not found on the California eProcure page. Screenshot saved to: {debugPath}");
        }

        LogProgress(contextLog, "Clicking export button...");

        var directDownloadTask = page.WaitForDownloadAsync(new PageWaitForDownloadOptions { Timeout = 15000 });
        var popupTask = context.WaitForPageAsync(new BrowserContextWaitForPageOptions { Timeout = 15000 });

        await exportButton.ClickAsync();

        try
        {
            return await directDownloadTask;
        }
        catch (TimeoutException)
        {
            // Expected when the portal opens #attachmentWrapperModal first
        }

        var modal = page.Locator("#attachmentWrapperModal");
        if (await modal.CountAsync() > 0 && await modal.IsVisibleAsync())
        {
            LogProgress(contextLog, "Detected #attachmentWrapperModal dialog; waiting for Download Attachment control...");

            var modalDownloadBtn = modal.Locator(
                "a:has-text('Download Attachment'), button:has-text('Download Attachment'), " +
                "a:has-text('Download'), button:has-text('Download'), a[href*='psfile'], a[id*='ATTACH']"
            ).First;

            await modalDownloadBtn.WaitForAsync(new LocatorWaitForOptions
            {
                State = WaitForSelectorState.Visible,
                Timeout = 30000
            });

            return await page.RunAndWaitForDownloadAsync(async () =>
            {
                await modalDownloadBtn.ClickAsync();
            }, new PageRunAndWaitForDownloadOptions { Timeout = 120000 });
        }

        IPage targetPage = page;
        try
        {
            targetPage = await popupTask;
        }
        catch (TimeoutException)
        {
            // No popup opened; check current page
        }

        var attachmentLink = targetPage.GetByText("Download Attachment", new PageGetByTextOptions { Exact = false }).First;
        return await targetPage.RunAndWaitForDownloadAsync(async () =>
        {
            await attachmentLink.ClickAsync();
        }, new PageRunAndWaitForDownloadOptions { Timeout = 120000 });
    }

    private static async Task<ILocator?> FindVisibleExportButtonAsync(IPage page, int timeoutMs)
    {
        var selectors = new[]
        {
            "[id$='$hexcel$0']",
            "a[id*='hexcel']",
            "button[id*='hexcel']",
            "a[title*='Download' i]",
            "button[title*='Download' i]",
            "img[alt*='Download' i]",
            "[aria-label*='Download' i]"
        };

        var startedAt = DateTime.UtcNow;

        while ((DateTime.UtcNow - startedAt).TotalMilliseconds < timeoutMs)
        {
            foreach (var frame in page.Frames)
            {
                foreach (var selector in selectors)
                {
                    var element = frame.Locator(selector).First;
                    if (await element.CountAsync() == 0)
                    {
                        continue;
                    }

                    if (await element.IsVisibleAsync())
                    {
                        return element;
                    }
                }
            }

            await page.WaitForTimeoutAsync(500);
        }

        return null;
    }

    private static async Task ClickSearchAsync(IPage page)
    {
        var selectorCandidates = new[]
        {
            "#RESP_INQA_WK_INQ_AUC_GO_PB",
            "input[id$='INQ_AUC_GO_PB']",
            "input[value='Search']",
            "button:has-text('Search')"
        };

        var startedAt = DateTime.UtcNow;

        while ((DateTime.UtcNow - startedAt).TotalMilliseconds < 30000)
        {
            foreach (var frame in page.Frames)
            {
                foreach (var selector in selectorCandidates)
                {
                    var locator = frame.Locator(selector).First;
                    if (await locator.CountAsync() > 0 && await locator.IsVisibleAsync())
                    {
                        await locator.ClickAsync();
                        return;
                    }
                }

                var roleButton = frame.GetByRole(AriaRole.Button, new FrameGetByRoleOptions { Name = "Search" }).First;
                if (await roleButton.CountAsync() > 0 && await roleButton.IsVisibleAsync())
                {
                    await roleButton.ClickAsync();
                    return;
                }
            }

            await page.WaitForTimeoutAsync(500);
        }

        throw new TimeoutException("Search button was not found on the California eProcure page.");
    }

    private static List<CalProcureRecord> ParseDownloadedFile(string filePath)
    {
        var text = File.ReadAllText(filePath);
        if (filePath.EndsWith(".html", StringComparison.OrdinalIgnoreCase) || text.Contains("<table", StringComparison.OrdinalIgnoreCase))
        {
            return ParseHtmlTable(text);
        }

        return ParseDelimited(text);
    }

    private static List<CalProcureRecord> ParseHtmlTable(string html)
    {
        var rows = new List<string[]>();

        var trChunks = Regex.Split(html, "<tr[^>]*>", RegexOptions.IgnoreCase);
        foreach (var chunk in trChunks.Skip(1))
        {
            var endIdx = chunk.IndexOf("</tr>", StringComparison.OrdinalIgnoreCase);
            var rowContent = endIdx >= 0 ? chunk[..endIdx] : chunk;

            var row = new List<string>();
            foreach (Match cell in Regex.Matches(rowContent, "<t[hd][^>]*>(?<value>.*?)</t[hd]>", RegexOptions.IgnoreCase | RegexOptions.Singleline))
            {
                var plain = Cleanup(cell.Groups["value"].Value) ?? string.Empty;
                row.Add(plain);
            }

            if (row.Count > 0)
            {
                rows.Add(row.ToArray());
            }
        }

        return ConvertRowsToRecords(rows);
    }

    private static List<CalProcureRecord> ParseDelimited(string text)
    {
        using var reader = new StringReader(text);
        var firstLine = reader.ReadLine();
        if (string.IsNullOrWhiteSpace(firstLine))
        {
            return [];
        }

        var delimiter = DetectDelimiter(firstLine);
        var rows = new List<string[]> { SplitLine(firstLine, delimiter) };

        while (reader.ReadLine() is { } line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            rows.Add(SplitLine(line, delimiter));
        }

        return ConvertRowsToRecords(rows);
    }

    private static List<CalProcureRecord> ConvertRowsToRecords(List<string[]> rows)
    {
        if (rows.Count < 2)
        {
            return [];
        }

        var headers = rows[0].Select(NormalizeHeader).ToArray();
        var records = new List<CalProcureRecord>();
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var i = 1; i < rows.Count; i++)
        {
            var row = rows[i];
            var eventIdText = GetValue(row, headers, "eventid", "eventnumber", "solicitationid", "bidid", "bidnumber");
            var departmentCode = GetValue(row, headers, "department", "departmentcode", "businessunit", "departmentid", "orgcode");
            var departmentName = GetValue(row, headers, "departmentname", "businessunitdescription") ?? "Unknown";
            var eventName = GetValue(row, headers, "eventname", "title", "description") ?? "Unknown";
            var format = GetValue(row, headers, "format", "submissionformat") ?? "Unknown";
            var eventType = GetValue(row, headers, "eventtype", "type") ?? "Unknown";
            var status = GetValue(row, headers, "status", "eventstatus") ?? "Unknown";
            var buyerName = GetValue(row, headers, "buyername", "contactname", "buyer");
            var buyerEmail = GetValue(row, headers, "buyeremail", "contactemail", "email");
            var endDateText = GetValue(row, headers, "enddate", "closingdate", "bidenddate", "eventenddate", "enddatetime");

            if (string.IsNullOrWhiteSpace(eventIdText))
            {
                continue;
            }

            var eventId = eventIdText;
            if (eventIdText.Contains('-', StringComparison.Ordinal) && string.IsNullOrWhiteSpace(departmentCode))
            {
                var parts = eventIdText.Split('-', 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2)
                {
                    departmentCode = parts[0];
                    eventId = parts[1];
                }
            }

            if (string.IsNullOrWhiteSpace(departmentCode) || string.IsNullOrWhiteSpace(eventId))
            {
                continue;
            }

            var key = $"{departmentCode}|{eventId}";
            if (!keys.Add(key))
            {
                continue;
            }

            var endDate = ParseDate(endDateText);
            records.Add(new CalProcureRecord(
                departmentCode,
                departmentName,
                eventId,
                eventName,
                format,
                eventType,
                endDate,
                status,
                buyerName,
                buyerEmail));
        }

        return records;
    }

    private static string? GetValue(string[] row, string[] headers, params string[] names)
    {
        foreach (var name in names)
        {
            var index = Array.FindIndex(headers, h => h == name);
            if (index < 0 || index >= row.Length)
            {
                continue;
            }

            var value = row[index].Trim();
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return null;
    }

    private static string[] SplitLine(string line, char delimiter)
    {
        var values = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }

                continue;
            }

            if (ch == delimiter && !inQuotes)
            {
                values.Add(current.ToString());
                current.Clear();
                continue;
            }

            current.Append(ch);
        }

        values.Add(current.ToString());
        return values.ToArray();
    }

    private static char DetectDelimiter(string line)
    {
        var candidates = new[] { ',', '\t', '|', ';' };
        return candidates
            .Select(c => (Delimiter: c, Count: line.Count(ch => ch == c)))
            .OrderByDescending(x => x.Count)
            .First().Delimiter;
    }

    private static DateTimeOffset ParseDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return DateTimeOffset.UtcNow;
        }

        var cleaned = Regex.Replace(value.Trim(), @"\s+(PST|PDT|UTC|GMT)$", string.Empty, RegexOptions.IgnoreCase);

        if (DateTime.TryParse(cleaned, CultureInfo.InvariantCulture, DateTimeStyles.None, out var localDt))
        {
            var pacificZone = TimeZoneInfo.FindSystemTimeZoneById("America/Los_Angeles");
            var offset = pacificZone.GetUtcOffset(localDt);
            return new DateTimeOffset(localDt, offset).ToUniversalTime();
        }

        return DateTimeOffset.UtcNow;
    }

    private static string NormalizeHeader(string value)
    {
        return Regex.Replace(value.ToLowerInvariant(), "[^a-z0-9]", string.Empty);
    }

    private static string? Cleanup(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var noTags = Regex.Replace(value, "<.*?>", " ", RegexOptions.Singleline);
        var decoded = WebUtility.HtmlDecode(noTags);
        var compact = Regex.Replace(decoded, @"\s+", " ").Trim();
        return string.IsNullOrWhiteSpace(compact) ? null : compact;
    }

    private sealed record CalProcureRecord(
        string DepartmentCode,
        string DepartmentName,
        string EventId,
        string EventName,
        string Format,
        string EventType,
        DateTimeOffset EndDate,
        string Status,
        string? BuyerName,
        string? BuyerEmail);
}
