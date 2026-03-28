using System.Net.Http;
using System.Text;
using System.Text.Json;
using OcrPdf.Models;

namespace OcrPdf.Services;

public sealed class GeminiNormalizationService
{
    private const string GeminiApiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite-preview:generateContent";
    private const string GeminiFallbackApiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent";
    private static readonly TimeSpan GeminiRequestTimeout = TimeSpan.FromMinutes(2);
    private readonly HttpClient _httpClient;

    public GeminiNormalizationService()
    {
        _httpClient = new HttpClient
        {
            Timeout = GeminiRequestTimeout
        };
    }

    public async Task<IReadOnlyList<OllamaExcelRow>> NormalizeOcrBlocksAsync(
        IReadOnlyList<OcrPageBlock> ocrBlocks,
        string geminiApiKey,
        Action<string> log,
        CancellationToken cancellationToken = default)
    {
        if (ocrBlocks.Count == 0)
        {
            log("Khong co OCR block de gui Gemini.");
            return [];
        }

        var rows = new List<OllamaExcelRow>();
        log($"Bat dau normalize {ocrBlocks.Count} block OCR (1 request Gemini cho batch nay).");

        cancellationToken.ThrowIfCancellationRequested();
        var chunkList = ocrBlocks.OrderBy(c => c.PageNumber).ToList();
        var firstPage = chunkList.First().PageNumber;
        var lastPage = chunkList.Last().PageNumber;
        log($"Normalize batch: {chunkList.Count} block (page {firstPage}-{lastPage}) qua Gemini.");

        try
        {
            var prompt = BuildNormalizationPrompt(chunkList);
            var generatedText = await GenerateNormalizationJsonAsync(prompt, chunkList.Count, geminiApiKey, log, cancellationToken);
            var normalizedRows = ParseNormalizedRows(generatedText, chunkList);
            rows.AddRange(normalizedRows);
            log($"Normalize batch thanh cong: {normalizedRows.Count} dong.");
        }
        catch (Exception ex)
        {
            log($"Normalize batch that bai: {ex.Message}");
            log($"Chi tiet loi batch: {ex}");
            rows.AddRange(chunkList.Select(c => CreateFailedRow(c.PageNumber, $"Normalize failed: {ex.Message}")));
        }

        return rows.OrderBy(r => r.PageNumber).ToList();
    }

    public async Task<string> GenerateNormalizationJsonAsync(
        string prompt,
        int expectedCount,
        string geminiApiKey,
        Action<string>? log = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(geminiApiKey))
        {
            var error = "Thieu Gemini API Key.";
            log?.Invoke(error);
            throw new InvalidOperationException(error);
        }

        var request = new
        {
            system_instruction = new
            {
                parts = new[]
                {
                    new
                    {
                        text = $"You are a technical data extraction expert. Extract 'drawingName' and 'drawingNo' from OCR blocks provided.\n\nGUIDELINES:\n1. MAPPING: Use the 'Page' number from each block for the 'pageNumber' field.\n2. MULTILINGUAL: Correct OCR errors (e.g., 'MÂT BÂNG' -> 'MẶT BẰNG').\n3. FORMATTING: Clean 'drawingNo' (e.g., 'V T . 1 2 3' -> 'VT.123').\n4. OUTPUT: Return ONLY a JSON ARRAY of exactly {expectedCount} objects. Each must have 'drawingName', 'drawingNo', and 'pageNumber'. No conversational text."
                    }
                }
            },
            contents = new[]
            {
                new
                {
                    role = "user",
                    parts = new[]
                    {
                        new { text = prompt }
                    }
                }
            },
            generationConfig = new
            {
                response_mime_type = "application/json",
                temperature = 0.0,
                topP = 0.95
            }
        };

        var currentApiUrl = GeminiApiUrl;
        var body = JsonSerializer.Serialize(request);
        const int maxRetries = 3;
        const int delayMs = 5000;

        for (var attempt = 0; attempt <= maxRetries; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var url = $"{currentApiUrl}?key={Uri.EscapeDataString(geminiApiKey)}";
            using var content = new StringContent(body, Encoding.UTF8, "application/json");
            HttpResponseMessage response;
            try
            {
                response = await _httpClient.PostAsync(url, content, cancellationToken);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                var error = $"Gemini request timeout sau {GeminiRequestTimeout.TotalSeconds:0} giay.";
                log?.Invoke(error);
                throw new InvalidOperationException(error);
            }

            using (response)
            {
                var responseBody = await response.Content.ReadAsStringAsync(cancellationToken);

                if (response.IsSuccessStatusCode)
                {
                    var normalizedJson = ExtractGeminiText(responseBody);
                    if (string.IsNullOrWhiteSpace(normalizedJson))
                    {
                        var error = "Gemini tra ve response rong sau khi extract text.";
                        log?.Invoke(error);
                        throw new InvalidOperationException(error);
                    }

                    return normalizedJson;
                }

                var statusCode = (int)response.StatusCode;
                if (statusCode == 503 && currentApiUrl != GeminiFallbackApiUrl)
                {
                    currentApiUrl = GeminiFallbackApiUrl;
                    log?.Invoke("Gemini HTTP 503. Chuyen sang model gemini-3.1-flash-lite-preview.");
                }

                if ((statusCode == 500 || statusCode == 503) && attempt < maxRetries)
                {
                    log?.Invoke($"Gemini HTTP {statusCode}. Retry {attempt + 1}/{maxRetries} sau {delayMs / 1000}s...");
                    await Task.Delay(delayMs, cancellationToken);
                    continue;
                }

                var error2 = $"Gemini HTTP {statusCode}: {responseBody}";
                log?.Invoke(error2);
                throw new InvalidOperationException(error2);
            }
        }

        throw new InvalidOperationException("Gemini request that bai sau tat ca retry.");
    }

    private static string BuildNormalizationPrompt(IReadOnlyList<OcrPageBlock> blocks)
    {
        var sb = new StringBuilder();
        sb.AppendLine("Extract data from these OCR blocks:");

        for (var i = 0; i < blocks.Count; i++)
        {
            sb.AppendLine($"[{i + 1}] [Page: {blocks[i].PageNumber}] OCR: {blocks[i].OcrText}");
        }

        return sb.ToString();
    }

    private static IReadOnlyList<OllamaExcelRow> ParseNormalizedRows(string rawResponse, IReadOnlyList<OcrPageBlock> sourceChunk)
    {
        var cleaned = ExtractJsonArrayText(rawResponse);
        using var jsonDocument = JsonDocument.Parse(cleaned);
        if (jsonDocument.RootElement.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("Normalize response khong phai JSON array.");
        }

        var result = new List<OllamaExcelRow>();
        foreach (var item in jsonDocument.RootElement.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var drawingName = item.TryGetProperty("drawingName", out var drawingNameElement)
                ? drawingNameElement.GetString() ?? string.Empty
                : string.Empty;

            var drawingNoRaw = item.TryGetProperty("drawingNo", out var drawingNoElement)
                ? drawingNoElement.GetString() ?? string.Empty
                : string.Empty;

            var pageNumber = TryReadPageNumber(item);
            if (pageNumber is null)
            {
                continue;
            }

            var drawingNo = new string(drawingNoRaw.Where(c => !char.IsWhiteSpace(c)).ToArray());
            result.Add(new OllamaExcelRow(drawingName.Trim(), drawingNo, pageNumber.Value, "OK", string.Empty));
        }

        var sourcePages = sourceChunk.Select(c => c.PageNumber).ToHashSet();
        var mappedPages = result.Select(r => r.PageNumber).ToHashSet();

        foreach (var page in sourcePages.Where(p => !mappedPages.Contains(p)))
        {
            result.Add(CreateFailedRow(page, "Khong co ket qua normalize cho page nay."));
        }

        return result.OrderBy(r => r.PageNumber).ToList();
    }

    private static int? TryReadPageNumber(JsonElement item)
    {
        if (!item.TryGetProperty("pageNumber", out var pageElement))
        {
            return null;
        }

        if (pageElement.ValueKind == JsonValueKind.Number && pageElement.TryGetInt32(out var pageNumber))
        {
            return pageNumber;
        }

        if (pageElement.ValueKind == JsonValueKind.String
            && int.TryParse(pageElement.GetString(), out var parsedPageNumber))
        {
            return parsedPageNumber;
        }

        return null;
    }

    private static string ExtractJsonArrayText(string raw)
    {
        var cleaned = raw.Trim();
        if (cleaned.StartsWith("```", StringComparison.Ordinal))
        {
            cleaned = cleaned.Replace("```json", string.Empty, StringComparison.OrdinalIgnoreCase)
                             .Replace("```", string.Empty, StringComparison.Ordinal)
                             .Trim();
        }

        var firstBracket = cleaned.IndexOf('[');
        var lastBracket = cleaned.LastIndexOf(']');
        if (firstBracket >= 0 && lastBracket > firstBracket)
        {
            return cleaned.Substring(firstBracket, lastBracket - firstBracket + 1);
        }

        return cleaned;
    }

    private static OllamaExcelRow CreateFailedRow(int pageNumber, string errorMessage)
    {
        var normalizedMessage = string.IsNullOrWhiteSpace(errorMessage)
            ? "Unknown error"
            : errorMessage.Trim();

        return new OllamaExcelRow(string.Empty, string.Empty, pageNumber, "FAILED", normalizedMessage);
    }

    private static string ExtractGeminiText(string responseBody)
    {
        try
        {
            using var jsonDocument = JsonDocument.Parse(responseBody);
            var root = jsonDocument.RootElement;
            if (root.TryGetProperty("candidates", out var candidates)
                && candidates.ValueKind == JsonValueKind.Array
                && candidates.GetArrayLength() > 0)
            {
                var first = candidates[0];
                if (first.TryGetProperty("content", out var content)
                    && content.TryGetProperty("parts", out var parts)
                    && parts.ValueKind == JsonValueKind.Array
                    && parts.GetArrayLength() > 0)
                {
                    var text = parts[0].GetProperty("text").GetString();
                    return text ?? string.Empty;
                }
            }
        }
        catch
        {
            // Fall through with raw response.
        }

        return responseBody;
    }
}