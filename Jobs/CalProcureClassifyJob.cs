using DocumentFormat.OpenXml.Packaging;
using Hangfire.Console;
using Hangfire.Server;
using Microsoft.Playwright;
using Npgsql;
using System.Data;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
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
        LogProgress(context, "Starting Cal eProcure classification job run.");

        if (_connection.State != ConnectionState.Open)
        {
            LogProgress(context, "Opening database connection...");
            await _connection.OpenAsync(cancellationToken);
        }

        var pendingEvents = await LoadPendingEventsAsync(cancellationToken);
        LogProgress(context, "Found {0} pending Cal eProcure events to classify.", pendingEvents.Count);

        if (pendingEvents.Count == 0)
        {
            LogProgress(context, "No pending events found. Job run completed.");
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
        LogProgress(context, "Browser initialized. Starting event processing loop.");

        var processedCount = 0;
        foreach (var ev in pendingEvents)
        {
            cancellationToken.ThrowIfCancellationRequested();
            LogProgress(context, "Processing event {0}-{1} ({2}/{3})", ev.DepartmentCode, ev.EventId, processedCount + 1, pendingEvents.Count);

            try
            {
                await ProcessSingleEventAsync(page, ev, context, cancellationToken);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed classifying event {Dept}-{EventId}", ev.DepartmentCode, ev.EventId);
                LogProgress(context, "Error processing {0}-{1}: {2}", ev.DepartmentCode, ev.EventId, ex.Message);
                await MarkEventErrorAsync(ev, ex.Message, cancellationToken);
            }

            processedCount++;
            LogProgress(context, "Finished event {0}-{1} ({2}/{3})", ev.DepartmentCode, ev.EventId, processedCount, pendingEvents.Count);

            // Polite delay to avoid rate limiting
            await page.WaitForTimeoutAsync(2000);
        }

        LogProgress(context, "Cal eProcure classification job run completed. Processed {0} events.", processedCount);
    }

    private async Task ProcessSingleEventAsync(
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
        LogProgress(context, "Scraped landing page for {0}-{1}. UNSPSC count: {2}, description length: {3}", ev.DepartmentCode, ev.EventId, unspscCodes.Count, description.Length);

        // Step 1: Fast in-memory triage
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

        // Step 2: Could be IT -> Open "View Event Package", download scope files, and confirm with LLM
        LogProgress(context, "Potential IT candidate ({0}-{1}). Opening Event Package...", ev.DepartmentCode, ev.EventId);

        var attachmentTexts = await DownloadAndExtractAttachmentsAsync(page, ev, context, cancellationToken);
        LogProgress(context, "Downloaded and extracted {0} attachment text blocks for {1}-{2}", attachmentTexts.Count, ev.DepartmentCode, ev.EventId);

        var llmResult = await ConfirmWithLlmAsync(ev, description, unspscCodes, attachmentTexts, cancellationToken);

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

        foreach (var frame in page.Frames)
        {
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
        IPage page,
        PendingEvent ev,
        PerformContext? context,
        CancellationToken cancellationToken)
    {
        var results = new List<(string FileName, string ExtractedText)>();

        ILocator? packageBtn = null;
        foreach (var frame in page.Frames)
        {
            var candidate = frame.Locator("button:has-text('View Event Package'), a:has-text('View Event Package'), input[value*='View Event Package']").First;
            if (await candidate.CountAsync() > 0 && await candidate.IsVisibleAsync())
            {
                packageBtn = candidate;
                break;
            }
        }

        if (packageBtn is null)
        {
            LogProgress(context, "View Event Package button not found for {0}-{1}.", ev.DepartmentCode, ev.EventId);
            return results;
        }

        LogProgress(context, "Opening Event Package for {0}-{1}.", ev.DepartmentCode, ev.EventId);
        await packageBtn.ClickAsync();
        await page.WaitForTimeoutAsync(4000);

        var downloadFolder = Path.Combine(Path.GetTempPath(), "caleprocure_pkg", $"{ev.DepartmentCode}_{ev.EventId}");
        Directory.CreateDirectory(downloadFolder);
        LogProgress(context, "Using download folder: {0}", downloadFolder);

        foreach (var frame in page.Frames)
        {
            var downloadLinks = frame.Locator("a:has-text('Download'), button:has-text('Download'), a[id*='ATTACH'], button[id*='ATTACH']");
            var count = await downloadLinks.CountAsync();

            for (int i = 0; i < Math.Min(count, 8); i++)
            {
                var link = downloadLinks.Nth(i);
                if (!await link.IsVisibleAsync())
                {
                    continue;
                }

                try
                {
                    var download = await page.RunAndWaitForDownloadAsync(async () =>
                    {
                        await link.ClickAsync();

                        var modalBtn = page.Locator("#attachmentWrapperModal a, #attachmentWrapperModal button").First;
                        if (await modalBtn.CountAsync() > 0 && await modalBtn.IsVisibleAsync())
                        {
                            await modalBtn.ClickAsync();
                        }
                    }, new PageRunAndWaitForDownloadOptions { Timeout = 25000 });

                    var suggestedName = download.SuggestedFilename;

                    if (BoilerplateFileKeywords.Any(b => suggestedName.Contains(b, StringComparison.OrdinalIgnoreCase)))
                    {
                        LogProgress(context, "Skipping boilerplate attachment for {0}-{1}: {2}", ev.DepartmentCode, ev.EventId, suggestedName);
                        continue;
                    }

                    var savedPath = Path.Combine(downloadFolder, suggestedName);
                    await download.SaveAsAsync(savedPath);

                    var text = ExtractTextFromFile(savedPath);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        results.Add((suggestedName, text));
                        await SaveAttachmentMetadataAsync(ev, suggestedName, savedPath, text, cancellationToken);
                        LogProgress(context, "Captured attachment text for {0}-{1}: {2}", ev.DepartmentCode, ev.EventId, suggestedName);
                    }
                }
                catch (TimeoutException)
                {
                    LogProgress(context, "Attachment download timed out for {0}-{1}. Continuing.", ev.DepartmentCode, ev.EventId);
                }
            }
        }

        LogProgress(context, "Attachment extraction complete for {0}-{1}. Captured {2} usable files.", ev.DepartmentCode, ev.EventId, results.Count);
        return results;
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
        CancellationToken cancellationToken)
    {
        // Supports both "Llm:*" and legacy "OpenAI:*" config keys
        var apiKey = _configuration["Llm:ApiKey"] ?? _configuration["OpenAI:ApiKey"];
        var model = _configuration["Llm:Model"] ?? _configuration["OpenAI:Model"] ?? "gpt-4o-mini";
        var configuredEndpoint = _configuration["Llm:Endpoint"];

        if (string.IsNullOrWhiteSpace(apiKey) && !model.StartsWith("ollama/", StringComparison.OrdinalIgnoreCase))
        {
            _logger.LogWarning("LLM API key not configured for event {Dept}-{EventId}; defaulting to IT classification fallback.", ev.DepartmentCode, ev.EventId);
            return new LlmClassificationResult(true, 0.60m, "Llm:ApiKey not configured; flagged as IT based on initial triage.");
        }

        var endpoint = ResolveEndpoint(model, configuredEndpoint);
        var normalizedModel = model.StartsWith("ollama/", StringComparison.OrdinalIgnoreCase)
            ? model["ollama/".Length..]
            : model;

        _logger.LogInformation(
            "Submitting LLM classification request for event {Dept}-{EventId} using model {Model} at {Endpoint}. Attachments included: {AttachmentCount}",
            ev.DepartmentCode,
            ev.EventId,
            normalizedModel,
            endpoint,
            attachments.Count);

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

        var payload = new
        {
            model = normalizedModel,
            messages = new[]
            {
                new
                {
                    role = "system",
                    content = """
                        You classify State of California procurement solicitations as IT (Information Technology) or Non-IT.
                        IT includes software development, SaaS/cloud licenses, cybersecurity, IT staffing/consulting, network/telecom infrastructure, data analytics, ERP/CRM, and computer hardware.
                        Non-IT includes physical construction, HVAC, vehicles, medical/lab equipment, janitorial, landscaping, general management consulting without IT deliverables, and physical supplies.
                        Respond strictly with raw JSON (no markdown code blocks) containing keys: "is_it" (boolean), "confidence" (number 0.0-1.0), and "reason" (concise 1-sentence explanation).
                        """
                },
                new
                {
                    role = "user",
                    content = promptBuilder.ToString()
                }
            },
            temperature = 0.1
        };

        var client = _httpClientFactory.CreateClient();
        using var request = new HttpRequestMessage(HttpMethod.Post, endpoint);

        if (!string.IsNullOrWhiteSpace(apiKey))
        {
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        }

        request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        using var response = await client.SendAsync(request, cancellationToken);
        var responseBody = await response.Content.ReadAsStringAsync(cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            throw new HttpRequestException($"LLM API ({endpoint}) returned {(int)response.StatusCode}: {responseBody}");
        }

        using var doc = JsonDocument.Parse(responseBody);
        var content = doc.RootElement
            .GetProperty("choices")
            .GetProperty("message")
            .GetProperty("content")
            .GetString();

        if (string.IsNullOrWhiteSpace(content))
        {
            _logger.LogWarning("LLM returned empty content for event {Dept}-{EventId}.", ev.DepartmentCode, ev.EventId);
            return new LlmClassificationResult(false, 0.50m, "Empty response from LLM.");
        }

        var cleanJson = StripMarkdownCodeFences(content);
        var parsed = JsonSerializer.Deserialize<LlmClassificationResult>(cleanJson);

        if (parsed is null)
        {
            _logger.LogWarning("Unable to parse LLM response JSON for event {Dept}-{EventId}.", ev.DepartmentCode, ev.EventId);
            return new LlmClassificationResult(false, 0.50m, "Could not parse LLM response.");
        }

        _logger.LogInformation(
            "LLM classification complete for event {Dept}-{EventId}. IsIt={IsIt}, Confidence={Confidence}",
            ev.DepartmentCode,
            ev.EventId,
            parsed.IsIt,
            parsed.Confidence);

        return parsed;
    }

    private static string ResolveEndpoint(string model, string? configuredEndpoint)
    {
        if (!string.IsNullOrWhiteSpace(configuredEndpoint))
        {
            return configuredEndpoint.Trim();
        }

        if (model.StartsWith("gemini", StringComparison.OrdinalIgnoreCase))
        {
            return "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions";
        }

        if (model.StartsWith("ollama/", StringComparison.OrdinalIgnoreCase))
        {
            return "http://localhost:11434/v1/chat/completions";
        }

        if (model.StartsWith("llama", StringComparison.OrdinalIgnoreCase) ||
            model.StartsWith("mixtral", StringComparison.OrdinalIgnoreCase))
        {
            return "https://api.groq.com/openai/v1/chat/completions";
        }

        // Default to OpenAI Chat Completions endpoint
        return "https://api.openai.com/v1/chat/completions";
    }

    private static string StripMarkdownCodeFences(string raw)
    {
        var trimmed = raw.Trim();
        if (trimmed.StartsWith("```", StringComparison.Ordinal))
        {
            trimmed = Regex.Replace(trimmed, @"^```(?:json)?\s*", string.Empty, RegexOptions.IgnoreCase);
            trimmed = Regex.Replace(trimmed, @"\s*```$", string.Empty);
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
