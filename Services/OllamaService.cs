using System.Net.Http;
using System.Text;
using System.Text.Json;
using DrawingOcrExtractor.Models;

namespace DrawingOcrExtractor.Services;

public sealed class OllamaService
{
    private const string OllamaApiUrl = "http://localhost:11434/api/generate";
    private const string OllamaOcrModel = "glm-ocr:latest";
    private const int OllamaTimeoutSeconds = 60;

    public async Task<OcrExtractionResult> ProcessPagesAsync(
        IReadOnlyList<Base64PageEntry> pages,
        Action<string> log,
        Action<int, int, string>? onProgress = null,
        CancellationToken cancellationToken = default)
    {
        if (pages.Count == 0)
        {
            return new OcrExtractionResult([], []);
        }

        using var httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(OllamaTimeoutSeconds)
        };

        var rows = new List<OllamaExcelRow>();
        var ocrBlocks = new List<OcrPageBlock>();
        var orderedPages = pages.OrderBy(p => p.PageNumber).ToList();

        for (var i = 0; i < orderedPages.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var page = orderedPages[i];
            var progressMessage = $"Đã xử lý Ollama trang {page.PageNumber}";
            log($"Gửi trang {page.PageNumber} lên Ollama");

            try
            {
                var prompt = $@"Extract all visible text from this image. The text may be in Vietnamese or English. \n\n
                    - Preserve line breaks and spacing as much as possible. \n
                    - Keep the original reading order (top to bottom, left to right). \n
                    - Do not summarize or interpret. \n
                    Return only the extracted text. No explanation.";
                var request = new OllamaRequest(OllamaOcrModel, prompt, [page.ImageBase64], false);
                var body = JsonSerializer.Serialize(request);

                try
                {
                    using var content = new StringContent(body, Encoding.UTF8, "application/json");
                    using var response = await httpClient.PostAsync(OllamaApiUrl, content, cancellationToken);
                    var responseBody = await response.Content.ReadAsStringAsync(cancellationToken);

                    if (!response.IsSuccessStatusCode)
                    {
                        var errorMessage = $"HTTP {(int)response.StatusCode}: {responseBody}";
                        log($"Trang {page.PageNumber} lỗi: {errorMessage}");
                        rows.Add(CreateFailedRow(page.PageNumber, errorMessage));
                        progressMessage = $"Trang {page.PageNumber} lỗi: {errorMessage}";
                    }
                    else
                    {
                        var ocrText = ExtractOllamaResponseText(responseBody);

                        if (string.IsNullOrWhiteSpace(ocrText))
                        {
                            var errorMessage = "OCR response rỗng.";
                            log($"Trang {page.PageNumber} lỗi: {errorMessage}");
                            rows.Add(CreateFailedRow(page.PageNumber, errorMessage));
                            progressMessage = $"Trang {page.PageNumber} lỗi: {errorMessage}";
                        }
                        else
                        {
                            ocrBlocks.Add(new OcrPageBlock(page.PageNumber, ocrText));
                        }
                    }
                }
                catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
                {
                    var errorMessage = $"Timeout khi gọi Ollama (vượt {OllamaTimeoutSeconds} giây).";
                    log($"Trang {page.PageNumber} lỗi: {errorMessage}");
                    rows.Add(CreateFailedRow(page.PageNumber, errorMessage));
                    progressMessage = $"Trang {page.PageNumber} lỗi: {errorMessage}";
                }
                catch (Exception ex)
                {
                    var errorMessage = ex.Message;
                    log($"Trang {page.PageNumber} lỗi: {errorMessage}");
                    rows.Add(CreateFailedRow(page.PageNumber, errorMessage));
                    progressMessage = $"Trang {page.PageNumber} lỗi: {errorMessage}";
                }
            }
            finally
            {
                onProgress?.Invoke(i + 1, orderedPages.Count, progressMessage);
                await Task.Delay(200, cancellationToken);
            }
        }

        if (ocrBlocks.Count == 0)
        {
            log("Không có OCR block hợp lệ để normalize.");
            return new OcrExtractionResult([], rows.OrderBy(r => r.PageNumber).ToList());
        }

        return new OcrExtractionResult(
            ocrBlocks.OrderBy(x => x.PageNumber).ToList(),
            rows.OrderBy(r => r.PageNumber).ToList());
    }

    private static string ExtractOllamaResponseText(string responseBody)
    {
        try
        {
            using var jsonDocument = JsonDocument.Parse(responseBody);
            if (jsonDocument.RootElement.TryGetProperty("response", out var responseElement))
            {
                return responseElement.GetString() ?? string.Empty;
            }
        }
        catch
        {
            // Fall back to raw body when response is not JSON.
        }

        return responseBody;
    }

    private static OllamaExcelRow CreateFailedRow(int pageNumber, string errorMessage)
    {
        var normalizedMessage = string.IsNullOrWhiteSpace(errorMessage)
            ? "Unknown error"
            : errorMessage.Trim();

        return new OllamaExcelRow(string.Empty, string.Empty, pageNumber, "FAILED", normalizedMessage);
    }
}

