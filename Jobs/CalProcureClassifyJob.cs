using System.Data;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Hangfire.Console;
using Hangfire.Server;
using Microsoft.Playwright;
using Npgsql;
using UglyToad.PdfPig;

namespace RfpDashboard.Jobs;

public sealed class CalProcureClassifyJob
{
    private readonly NpgsqlConnection _connection;
    private readonly IConfiguration _configuration;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<CalProcureClassifyJob> _logger;

    private static readonly string[] ItUnspscPrefixes = ["43", "8111", "80101507"];
    private static readonly string[] HighSignalItDepartments = ["7502", "8885"]; // CDT, FI$Cal

    private static readonly string[] StrongItKeywords =
    [
        "software", "saas", "cloud", "cybersecurity", "cyber security", "information technology",
        "network", "server", "database", "sql", "crm", "erp", "api", "data analytics",
        "artificial intelligence", "machine learning", "help desk", "helpdesk", "it consulting",
        "application development", "web application", "mobile app", "microsoft", "azure", "aws",
        "salesforce", "servicenow", "workday", "oracle", "cisco", "palo alto", "splunk",
        "snowflake", "tableau", "power bi", "sharepoint", "dynamics 365", "mainframe",
        "telecom", "telecommunications", "voip", "fiber optic", "data center", "penetration testing",
        "vulnerability", "iam", "identity management", "devops", "agile", "scrum", "qa testing"
    ];

    private static readonly string[] StrongNonItKeywords =
    [
        "janitorial", "custodial", "landscaping", "tree trimming", "arborist", "weed abatement",
        "paving", "asphalt", "concrete", "roofing", "plumbing", "hvac", "electrical repair",
        "guardrail", "fencing", "painting", "flooring", "carpet", "elevator maintenance",
        "towing", "trash", "waste disposal", "hazardous waste", "laundry", "linen",
        "pest control", "security guard", "armed guard", "unarmed guard",
        "fleet", "vehicle", "truck", "tractor", "forklift", "tire", "automotive",
        "medical supplies", "dental", "pharmaceutical", "laboratory supplies", "ambulance",
        "catering", "food service", "groceries", "uniforms", "furniture", "office supplies"
    ];

    private static readonly string[] BoilerplateFileKeywords =
    [
        "std204", "std 204", "dvbe", "darfur", "payee", "civil rights", "iran contracting",
        "commercially useful", "cuf", "bidder instructions", "general provisions", "sample contract"
    ];

    public CalProcureClassifyJob(
        NpgsqlConnection connection,
        IConfiguration configuration,
        IHttpClientFactory httpClientFactory,
        ILogger<CalProcureClassifyJob> logger)
    {
        _connection = connection;
        _configuration = configuration;
        _httpClientFactory = httpClientFactory;
        _logger = logger;
    }

    public async Task ExecuteAsync(PerformContext? context, CancellationToken cancellationToken = default)
    {
        if (_connection.State != ConnectionState.Open)
        {
            await _connection.OpenAsync(cancellationToken);
        }

        var pendingEvents = await LoadPendingEventsAsync(cancellationToken);
        LogProgress(context, "Found {0} pending Cal eProcure events to classify.", pendingEvents.Count);

        if (pendingEvents.Count == 0)
        {
            return;
        }

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

        browserContext.SetDefaultTimeout(60000);
        var page = await browserContext.NewPageAsync();

        for (int idx = 0; idx < pendingEvents.Count; idx++)
        {
            var ev = pendingEvents[idx];
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                await ProcessSingleEventAsync(browserContext, page, ev, context, cancellationToken);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed classifying event {Dept}-{EventId}", ev.DepartmentCode, ev.EventId);
                LogProgress(context, "Error processing {0}-{1}: {2}", ev.DepartmentCode, ev.EventId, ex.Message);
                await MarkEventErrorAsync(ev, ex.Message, cancellationToken);
            }

            LogProgress(context, "Finished event {0}-{1} ({2}/{3})", ev.DepartmentCode, ev.EventId, idx + 1, pendingEvents.Count);
            await page.WaitForTimeoutAsync(2000);
        }
    }

    private async Task ProcessSingleEventAsync(
        IBrowserContext browserContext,
        IPage page,
        PendingEvent ev,
        PerformContext? context,
        CancellationToken cancellationToken)
    {
        var eventUrl = $"https://caleprocure.ca.gov/event/{ev.DepartmentCode}/{ev.EventId}";
        LogProgress(context, "Visiting {0} ({1})", eventUrl, ev.EventName);

        await page.GotoAsync(eventUrl, new PageGotoOptions { WaitUntil = WaitUntilState.DOMContentLoaded });
        await page.WaitForTimeoutAsync(3000);

        var (description, unspscCodes) = await ScrapeLandingPageDetailsAsync(page);
        var triage = EvaluateInMemoryTriage(ev, description, unspscCodes);
        LogProgress(context, "Triage decision for {0}-{1}: {2}", ev.DepartmentCode, ev.EventId, triage);

        if (triage == TriageDecision.DefinitelyNonIt)
        {
            await UpdateEventClassificationAsync(
                ev,
                triageStatus: "NON_IT_SKIPPED",
                isIt: false,
                confidence: 0.95m,
                reason: "Filtered out by in-memory UNSPSC/keyword rules on landing page.",
                description: null,
                unspscCodes: string.Join(",", unspscCodes),
                cancellationToken);

            LogProgress(context, "Skipped Non-IT event: {0}-{1}", ev.DepartmentCode, ev.EventId);
            return;
        }

        LogProgress(context, "Potential IT candidate ({0}-{1}). Opening Event Package...", ev.DepartmentCode, ev.EventId);

        var attachmentTexts = await DownloadAndExtractAttachmentsAsync(browserContext, page, ev, context, cancellationToken);
        LogProgress(context, "Downloaded and extracted {0} attachment text blocks for {1}-{2}", attachmentTexts.Count, ev.DepartmentCode, ev.EventId);

        var llmResult = await ConfirmWithLlmAsync(ev, description, unspscCodes, attachmentTexts, context, cancellationToken);
        var finalStatus = llmResult.IsIt ? "IT_CONFIRMED" : "NON_IT_CONFIRMED";

        await UpdateEventClassificationAsync(
            ev,
            triageStatus: finalStatus,
            isIt: llmResult.IsIt,
            confidence: llmResult.Confidence,
            reason: llmResult.Reason,
            description: llmResult.IsIt ? description : null,
            unspscCodes: string.Join(",", unspscCodes),
            cancellationToken);

        LogProgress(
            context,
            "Classified {0}-{1} as {2} (Confidence: {3:P0}) - {4}",
            ev.DepartmentCode,
            ev.EventId,
            finalStatus,
            llmResult.Confidence,
            llmResult.Reason);
    }

    private static async Task<(string Description, List<string> UnspscCodes)> ScrapeLandingPageDetailsAsync(IPage page)
    {
        var allText = new StringBuilder();
        var unspscSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var frame in page.Frames.ToList())
        {
            if (frame.IsDetached)
            {
                continue;
            }

            try
            {
                var bodyText = await frame.InnerTextAsync("body");
                if (string.IsNullOrWhiteSpace(bodyText))
                {
                    continue;
                }

                allText.AppendLine(bodyText);

                foreach (Match match in Regex.Matches(bodyText, @"\b\d{8}\b"))
                {
                    unspscSet.Add(match.Value);
                }
            }
            catch
            {
                // Ignore cross-origin or detached frames
            }
        }

        var combined = Regex.Replace(allText.ToString(), @"\s+", " ").Trim();
        if (combined.Length > 6000)
        {
            combined = combined[..6000];
        }

        return (combined, unspscSet.ToList());
    }

    private static TriageDecision EvaluateInMemoryTriage(
        PendingEvent ev,
        string landingDescription,
        List<string> unspscCodes)
    {
        if (unspscCodes.Count > 0)
        {
            bool hasItUnspsc = unspscCodes.Any(code =>
                ItUnspscPrefixes.Any(prefix => code.StartsWith(prefix, StringComparison.Ordinal)));

            if (hasItUnspsc)
            {
                return TriageDecision.CouldBeIt;
            }
        }

        if (HighSignalItDepartments.Contains(ev.DepartmentCode, StringComparer.OrdinalIgnoreCase))
        {
            return TriageDecision.CouldBeIt;
        }

        var combinedText = $"{ev.EventName} {ev.DepartmentName} {landingDescription}".ToLowerInvariant();

        int itHits = StrongItKeywords.Count(k => combinedText.Contains(k, StringComparison.Ordinal));
        int nonItHits = StrongNonItKeywords.Count(k => combinedText.Contains(k, StringComparison.Ordinal));

        if (itHits > 0)
        {
            return TriageDecision.CouldBeIt;
        }

        if (nonItHits > 0 && itHits == 0)
        {
            return TriageDecision.DefinitelyNonIt;
        }

        if (unspscCodes.Count > 0 && itHits == 0)
        {
            return TriageDecision.DefinitelyNonIt;
        }

        return TriageDecision.CouldBeIt;
    }

    private async Task<List<(string FileName, string ExtractedText)>> DownloadAndExtractAttachmentsAsync(
        IBrowserContext browserContext,
        IPage page,
        PendingEvent ev,
        PerformContext? context,
        CancellationToken cancellationToken)
    {
        var results = new List<(string FileName, string ExtractedText)>();

        var packageBtn = await FindViewEventPackageButtonAsync(page, 15000);
        if (packageBtn is null)
        {
            LogProgress(context, "View Event Package button not found for {0}-{1}.", ev.DepartmentCode, ev.EventId);
            return results;
        }

        await packageBtn.ClickAsync();
        await page.WaitForTimeoutAsync(3000);

        var downloadFolder = Path.Combine(Path.GetTempPath(), "caleprocure_pkg", $"{ev.DepartmentCode}_{ev.EventId}");
        Directory.CreateDirectory(downloadFolder);
        LogProgress(context, "Using download folder: {0}", downloadFolder);

        int buttonCount = await WaitForAttachmentButtonsAsync(page, 12000);
        LogProgress(context, "Found {0} attachment download buttons for {1}-{2}.", buttonCount, ev.DepartmentCode, ev.EventId);

        int maxToProcess = Math.Min(buttonCount, 8);
        for (int i = 0; i < maxToProcess; i++)
        {
            await EnsureAttachmentModalClosedAsync(page);
            await CloseExtraPagesAsync(browserContext, page);

            var rowButton = await FindAttachmentRowButtonAsync(page, i);
            if (rowButton is null)
            {
                LogProgress(context, "Attachment button index {0} was not visible for {1}-{2}.", i, ev.DepartmentCode, ev.EventId);
                continue;
            }

            try
            {
                var downloadedFile = await DownloadAttachmentViaModalHrefAsync(browserContext, page, rowButton, downloadFolder, context);
                if (downloadedFile is null)
                {
                    LogProgress(context, "Attachment download returned empty for {0}-{1} (index {2}).", ev.DepartmentCode, ev.EventId, i);
                    continue;
                }

                var (fileName, savedPath) = downloadedFile.Value;

                if (BoilerplateFileKeywords.Any(b => fileName.Contains(b, StringComparison.OrdinalIgnoreCase)))
                {
                    LogProgress(context, "Skipped boilerplate attachment: {0}", fileName);
                    continue;
                }

                var text = ExtractTextFromFile(savedPath);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    results.Add((fileName, text));
                    await SaveAttachmentMetadataAsync(ev, fileName, savedPath, text, cancellationToken);
                    LogProgress(context, "Extracted {0} chars from {1}", text.Length, fileName);
                }
            }
            catch (Exception ex)
            {
                LogProgress(context, "Attachment download error for {0}-{1} (index {2}): {3}", ev.DepartmentCode, ev.EventId, i, ex.Message);
            }
            finally
            {
                await EnsureAttachmentModalClosedAsync(page);
                await CloseExtraPagesAsync(browserContext, page);
            }
        }

        return results;
    }

    private async Task<(string FileName, string SavedPath)?> DownloadAttachmentViaModalHrefAsync(
        IBrowserContext browserContext,
        IPage page,
        ILocator rowButton,
        string downloadFolder,
        PerformContext? context)
    {
        await rowButton.ScrollIntoViewIfNeededAsync();
        await rowButton.ClickAsync();

        // Wait for #attachmentWrapperModal to render <a id="downloadButton" href="...">
        var downloadLink = await WaitForModalDownloadAnchorAsync(page, 15000);
        if (downloadLink is null)
        {
            return null;
        }

        var href = await downloadLink.GetAttributeAsync("href");
        if (string.IsNullOrWhiteSpace(href))
        {
            return null;
        }

        var absoluteUrl = new Uri(new Uri(page.Url), href).ToString();
        var rawFileName = Path.GetFileName(new Uri(absoluteUrl).LocalPath);
        var decodedFileName = WebUtility.UrlDecode(rawFileName);
        if (string.IsNullOrWhiteSpace(decodedFileName))
        {
            decodedFileName = $"attachment-{Guid.NewGuid():N}.bin";
        }

        // Fetch file bytes directly using the browser context's authenticated cookies (avoids Edge inline PDF viewer lock)
        var apiResponse = await browserContext.APIRequest.GetAsync(absoluteUrl, new APIRequestContextOptions
        {
            Timeout = 45000
        });

        if (!apiResponse.Ok)
        {
            LogProgress(context, "Authenticated download HTTP {0} for {1}", apiResponse.Status, decodedFileName);
            return null;
        }

        var bytes = await apiResponse.BodyAsync();
        var savedPath = Path.Combine(downloadFolder, decodedFileName);
        await File.WriteAllBytesAsync(savedPath, bytes);
        LogProgress(context, "Saved attachment ({0} bytes): {1}", bytes.Length, decodedFileName);

        return (decodedFileName, savedPath);
    }

    private static async Task<ILocator?> WaitForModalDownloadAnchorAsync(IPage page, int timeoutMs)
    {
        var selectors = new[]
        {
            "#attachmentWrapperModal a#downloadButton[href]",
            "#attachmentWrapperModal a[href*='viewredirect']",
            "#attachmentWrapperModal a:has-text('Download Attachment')[href]",
            ".modal.in a#downloadButton[href]",
            ".modal.show a#downloadButton[href]"
        };

        var startedAt = DateTime.UtcNow;

        while ((DateTime.UtcNow - startedAt).TotalMilliseconds < timeoutMs)
        {
            foreach (var selector in selectors)
            {
                var topAnchor = page.Locator(selector).First;
                if (await topAnchor.CountAsync() > 0 && await topAnchor.IsVisibleAsync())
                {
                    var href = await topAnchor.GetAttributeAsync("href");
                    if (!string.IsNullOrWhiteSpace(href) && href != "#")
                    {
                        return topAnchor;
                    }
                }
            }

            foreach (var frame in page.Frames.ToList())
            {
                if (frame.IsDetached)
                {
                    continue;
                }

                foreach (var selector in selectors)
                {
                    try
                    {
                        var frameAnchor = frame.Locator(selector).First;
                        if (await frameAnchor.CountAsync() > 0 && await frameAnchor.IsVisibleAsync())
                        {
                            var href = await frameAnchor.GetAttributeAsync("href");
                            if (!string.IsNullOrWhiteSpace(href) && href != "#")
                            {
                                return frameAnchor;
                            }
                        }
                    }
                    catch
                    {
                        // Ignore transient frame detachments
                    }
                }
            }

            await page.WaitForTimeoutAsync(300);
        }

        return null;
    }

    private static async Task<ILocator?> FindViewEventPackageButtonAsync(IPage page, int timeoutMs)
    {
        var selectors = new[]
        {
            "button:has-text('View Event Package')",
            "a:has-text('View Event Package')",
            "input[value*='View Event Package' i]",
            "[id*='VIEW_PKG' i]",
            "[id*='EVENT_PKG' i]",
            "[aria-label*='View Event Package' i]"
        };

        var startedAt = DateTime.UtcNow;

        while ((DateTime.UtcNow - startedAt).TotalMilliseconds < timeoutMs)
        {
            foreach (var frame in page.Frames.ToList())
            {
                if (frame.IsDetached)
                {
                    continue;
                }

                foreach (var selector in selectors)
                {
                    try
                    {
                        var candidate = frame.Locator(selector).First;
                        if (await candidate.CountAsync() > 0 && await candidate.IsVisibleAsync())
                        {
                            return candidate;
                        }
                    }
                    catch
                    {
                        // Ignore transient frame detachments
                    }
                }
            }

            await page.WaitForTimeoutAsync(500);
        }

        return null;
    }

    private static async Task<int> WaitForAttachmentButtonsAsync(IPage page, int timeoutMs)
    {
        var startedAt = DateTime.UtcNow;

        while ((DateTime.UtcNow - startedAt).TotalMilliseconds < timeoutMs)
        {
            int count = await CountAttachmentRowButtonsAsync(page);
            if (count > 0)
            {
                return count;
            }

            await page.WaitForTimeoutAsync(500);
        }

        return 0;
    }

    private static async Task<int> CountAttachmentRowButtonsAsync(IPage page)
    {
        foreach (var frame in page.Frames.ToList())
        {
            if (frame.IsDetached)
            {
                continue;
            }

            try
            {
                var buttons = frame.Locator("button[id^='PV_ATTACH_WRK_SCM_DOWNLOAD$'], a[id^='PV_ATTACH_WRK_SCM_DOWNLOAD$']");
                int count = await buttons.CountAsync();
                if (count > 0)
                {
                    return count;
                }
            }
            catch
            {
                // Ignore detached frame
            }
        }

        return 0;
    }

    private static async Task<ILocator?> FindAttachmentRowButtonAsync(IPage page, int index)
    {
        foreach (var frame in page.Frames.ToList())
        {
            if (frame.IsDetached)
            {
                continue;
            }

            try
            {
                var buttons = frame.Locator("button[id^='PV_ATTACH_WRK_SCM_DOWNLOAD$'], a[id^='PV_ATTACH_WRK_SCM_DOWNLOAD$']");
                int count = await buttons.CountAsync();
                if (index < count)
                {
                    var target = buttons.Nth(index);
                    if (await target.IsVisibleAsync())
                    {
                        return target;
                    }
                }
            }
            catch
            {
                // Ignore detached frame
            }
        }

        return null;
    }

    private static async Task EnsureAttachmentModalClosedAsync(IPage page)
    {
        try
        {
            var modal = page.Locator("#attachmentWrapperModal, .modal.in, .modal.show").First;
            if (await modal.CountAsync() > 0 && await modal.IsVisibleAsync())
            {
                var closeBtn = modal.Locator("button[data-dismiss='modal'], button:has-text('Close'), button.close").First;
                if (await closeBtn.CountAsync() > 0 && await closeBtn.IsVisibleAsync())
                {
                    await closeBtn.ClickAsync(new LocatorClickOptions { Timeout = 3000 });
                }
                else
                {
                    await page.Keyboard.PressAsync("Escape");
                }

                await modal.WaitForAsync(new LocatorWaitForOptions
                {
                    State = WaitForSelectorState.Hidden,
                    Timeout = 5000
                });
            }
        }
        catch
        {
            // Best-effort modal cleanup
        }
    }

    private static async Task CloseExtraPagesAsync(IBrowserContext browserContext, IPage primaryPage)
    {
        foreach (var p in browserContext.Pages.ToList())
        {
            if (p != primaryPage && !p.IsClosed)
            {
                try
                {
                    await p.CloseAsync();
                }
                catch
                {
                    // Ignore already closed pages
                }
            }
        }
    }

    private static string ExtractTextFromFile(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var sb = new StringBuilder();

        try
        {
            if (ext == ".pdf")
            {
                using var pdf = PdfDocument.Open(filePath);
                foreach (var page in pdf.GetPages().Take(12))
                {
                    sb.AppendLine(page.Text);
                }
            }
            else if (ext == ".docx")
            {
                using var doc = WordprocessingDocument.Open(filePath, false);
                var body = doc.MainDocumentPart?.Document.Body;
                if (body is not null)
                {
                    sb.AppendLine(body.InnerText);
                }
            }
            else if (ext is ".txt" or ".csv" or ".html" or ".htm")
            {
                sb.AppendLine(File.ReadAllText(filePath));
            }
        }
        catch
        {
            // Ignore unreadable or password-protected files
        }

        var cleaned = Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
        return cleaned.Length > 8000 ? cleaned[..8000] : cleaned;
    }

    private async Task<LlmClassificationResult> ConfirmWithLlmAsync(
        PendingEvent ev,
        string landingDescription,
        List<string> unspscCodes,
        List<(string FileName, string ExtractedText)> attachments,
        PerformContext? context,
        CancellationToken cancellationToken)
    {
        var apiKey = _configuration["Llm:ApiKey"] ?? _configuration["OpenAI:ApiKey"];
        var model = _configuration["Llm:Model"] ?? _configuration["OpenAI:Model"] ?? "gpt-4o-mini";
        var configuredEndpoint = _configuration["Llm:Endpoint"];

        if (string.IsNullOrWhiteSpace(apiKey))
        {
            return new LlmClassificationResult(true, 0.60m, "Llm:ApiKey not configured; flagged as IT based on initial triage.");
        }

        var endpoint = ResolveLlmEndpoint(model, configuredEndpoint);

        var promptBuilder = new StringBuilder();
        promptBuilder.AppendLine($"Event ID: {ev.DepartmentCode}-{ev.EventId}");
        promptBuilder.AppendLine($"Department: {ev.DepartmentName}");
        promptBuilder.AppendLine($"Title: {ev.EventName}");
        promptBuilder.AppendLine($"UNSPSC Codes: {string.Join(", ", unspscCodes)}");
        promptBuilder.AppendLine($"Landing Description: {landingDescription}");

        foreach (var (fileName, text) in attachments.Take(3))
        {
            promptBuilder.AppendLine($"\n--- Attachment: {fileName} ---");
            promptBuilder.AppendLine(text.Length > 4000 ? text[..4000] : text);
        }

        bool supportsTemperature = !model.StartsWith("o1", StringComparison.OrdinalIgnoreCase)
                                && !model.StartsWith("o3", StringComparison.OrdinalIgnoreCase)
                                && !model.StartsWith("gpt-5", StringComparison.OrdinalIgnoreCase);

        object payload = supportsTemperature
            ? new
            {
                model,
                temperature = 0.1,
                messages = BuildMessages(promptBuilder.ToString())
            }
            : new
            {
                model,
                messages = BuildMessages(promptBuilder.ToString())
            };

        var client = _httpClientFactory.CreateClient();
        using var request = new HttpRequestMessage(HttpMethod.Post, endpoint);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        using var response = await client.SendAsync(request, cancellationToken);
        var responseBody = await response.Content.ReadAsStringAsync(cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            LogProgress(context, "LLM API error ({0}) from {1} [model={2}]: {3}", (int)response.StatusCode, endpoint, model, responseBody);
            throw new HttpRequestException($"LLM API returned {(int)response.StatusCode} ({response.ReasonPhrase}): {responseBody}");
        }

        using var doc = JsonDocument.Parse(responseBody);
        var content = doc.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString();

        if (string.IsNullOrWhiteSpace(content))
        {
            return new LlmClassificationResult(false, 0.50m, "Empty response from LLM.");
        }

        var cleanedJson = StripMarkdownCodeFences(content);
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        var parsed = JsonSerializer.Deserialize<LlmClassificationResult>(cleanedJson, options);
        return parsed ?? new LlmClassificationResult(false, 0.50m, "Could not parse LLM response.");
    }

    private static object[] BuildMessages(string userPrompt) =>
    [
        new
        {
            role = "system",
            content = """
                You classify State of California procurement solicitations as IT (Information Technology) or Non-IT.
                IT includes software development, SaaS/cloud licenses, cybersecurity, IT staffing/consulting, network/telecom infrastructure, data analytics, ERP/CRM, and computer hardware.
                Non-IT includes physical construction, HVAC, vehicles, medical/lab equipment, janitorial, landscaping, general management consulting without IT deliverables, and physical supplies.
                Respond strictly as raw JSON (no markdown fences) with keys: "is_it" (boolean), "confidence" (number 0.0-1.0), and "reason" (concise 1-sentence explanation).
                """
        },
        new
        {
            role = "user",
            content = userPrompt
        }
    ];

    private static string ResolveLlmEndpoint(string model, string? configuredEndpoint)
    {
        if (!string.IsNullOrWhiteSpace(configuredEndpoint))
        {
            return configuredEndpoint.Trim();
        }

        if (model.StartsWith("gemini", StringComparison.OrdinalIgnoreCase))
        {
            return "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions";
        }

        if (model.StartsWith("llama", StringComparison.OrdinalIgnoreCase) ||
            model.StartsWith("mixtral", StringComparison.OrdinalIgnoreCase))
        {
            return "https://api.groq.com/openai/v1/chat/completions";
        }

        return "https://api.openai.com/v1/chat/completions";
    }

    private static string StripMarkdownCodeFences(string raw)
    {
        var trimmed = raw.Trim();
        if (trimmed.StartsWith("```", StringComparison.Ordinal))
        {
            var firstNewLine = trimmed.IndexOf('\n');
            if (firstNewLine >= 0)
            {
                trimmed = trimmed[(firstNewLine + 1)..];
            }

            if (trimmed.EndsWith("```", StringComparison.Ordinal))
            {
                trimmed = trimmed[..^3];
            }
        }

        return trimmed.Trim();
    }

    private async Task<List<PendingEvent>> LoadPendingEventsAsync(CancellationToken cancellationToken)
    {
        var list = new List<PendingEvent>();
        await using var cmd = _connection.CreateCommand();
        cmd.CommandText = """
            SELECT department_code, department_name, event_id, event_name
            FROM public.t_cal_eprocure
            WHERE triage_status = 'PENDING'
            ORDER BY end_date ASC
            LIMIT 50;
            """;

        await using var reader = await cmd.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            list.Add(new PendingEvent(
                reader.GetString(0),
                reader.IsDBNull(1) ? string.Empty : reader.GetString(1),
                reader.GetString(2),
                reader.IsDBNull(3) ? string.Empty : reader.GetString(3)));
        }

        return list;
    }

    private async Task UpdateEventClassificationAsync(
        PendingEvent ev,
        string triageStatus,
        bool isIt,
        decimal confidence,
        string reason,
        string? description,
        string unspscCodes,
        CancellationToken cancellationToken)
    {
        await using var cmd = _connection.CreateCommand();
        cmd.CommandText = """
            UPDATE public.t_cal_eprocure
            SET triage_status = @triageStatus,
                is_it = @isIt,
                llm_confidence = @confidence,
                llm_reason = @reason,
                full_description = @description,
                unspsc_codes = @unspscCodes,
                visited_at = NOW()
            WHERE department_code = @departmentCode
              AND event_id = @eventId;
            """;

        cmd.Parameters.AddWithValue("triageStatus", triageStatus);
        cmd.Parameters.AddWithValue("isIt", isIt);
        cmd.Parameters.AddWithValue("confidence", confidence);
        cmd.Parameters.AddWithValue("reason", reason);
        cmd.Parameters.AddWithValue("description", (object?)description ?? DBNull.Value);
        cmd.Parameters.AddWithValue("unspscCodes", string.IsNullOrWhiteSpace(unspscCodes) ? DBNull.Value : unspscCodes);
        cmd.Parameters.AddWithValue("departmentCode", ev.DepartmentCode);
        cmd.Parameters.AddWithValue("eventId", ev.EventId);

        await cmd.ExecuteNonQueryAsync(cancellationToken);
    }

    private async Task SaveAttachmentMetadataAsync(
        PendingEvent ev,
        string fileName,
        string localPath,
        string extractedText,
        CancellationToken cancellationToken)
    {
        await using var cmd = _connection.CreateCommand();
        cmd.CommandText = """
            INSERT INTO public.t_cal_eprocure_attachments
            (department_code, event_id, file_name, local_path, extracted_text)
            VALUES
            (@departmentCode, @eventId, @fileName, @localPath, @extractedText);
            """;

        cmd.Parameters.AddWithValue("departmentCode", ev.DepartmentCode);
        cmd.Parameters.AddWithValue("eventId", ev.EventId);
        cmd.Parameters.AddWithValue("fileName", fileName);
        cmd.Parameters.AddWithValue("localPath", localPath);
        cmd.Parameters.AddWithValue("extractedText", extractedText);

        await cmd.ExecuteNonQueryAsync(cancellationToken);
    }

    private async Task MarkEventErrorAsync(PendingEvent ev, string error, CancellationToken cancellationToken)
    {
        await using var cmd = _connection.CreateCommand();
        cmd.CommandText = """
            UPDATE public.t_cal_eprocure
            SET triage_status = 'ERROR',
                llm_reason = @error,
                visited_at = NOW()
            WHERE department_code = @departmentCode
              AND event_id = @eventId;
            """;

        cmd.Parameters.AddWithValue("error", error.Length > 500 ? error[..500] : error);
        cmd.Parameters.AddWithValue("departmentCode", ev.DepartmentCode);
        cmd.Parameters.AddWithValue("eventId", ev.EventId);

        await cmd.ExecuteNonQueryAsync(cancellationToken);
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

    private enum TriageDecision
    {
        DefinitelyNonIt,
        CouldBeIt
    }

    private sealed record PendingEvent(
        string DepartmentCode,
        string DepartmentName,
        string EventId,
        string EventName);

    private sealed record LlmClassificationResult(
        [property: JsonPropertyName("is_it")] bool IsIt,
        [property: JsonPropertyName("confidence")] decimal Confidence,
        [property: JsonPropertyName("reason")] string Reason);
}
