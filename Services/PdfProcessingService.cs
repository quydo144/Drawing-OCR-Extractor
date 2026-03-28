using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text.Json;
using DrawingOcrExtractor.Models;
using PdfiumViewer;

namespace DrawingOcrExtractor.Services;

public sealed class PdfProcessingService
{
    private const double OutputScale = 0.5;
    private const int MaxParallelism = 10;
    private const int OutputDpi = 170;

    public ConversionResult ConvertPdfPagesToImagesAndBase64(
        string pdfPath,
        Action<ConversionPageProgress>? onProgress = null,
        CancellationToken cancellationToken = default)
    {
        var logs = new List<string>();
        var pdfDirectory = Path.GetDirectoryName(pdfPath) ?? Environment.CurrentDirectory;
        var pdfName = Path.GetFileNameWithoutExtension(pdfPath);
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
        var outputDir = Path.Combine(pdfDirectory, $"{pdfName}_pages_{timestamp}");
        Directory.CreateDirectory(outputDir);

        var base64File = Path.Combine(outputDir, "pages_base64.json");

        using var pdfDocument = PdfDocument.Load(pdfPath);
        var pageCount = pdfDocument.PageCount;
        logs.Add($"Sá»‘ trang PDF: {pageCount}");
        var configuredParallelism = Math.Max(1, Math.Min(MaxParallelism, Environment.ProcessorCount));
        var effectiveParallelism = Math.Min(configuredParallelism, Math.Max(1, pageCount));
        logs.Add($"Xu ly song song PDF -> image -> base64 (toi da {effectiveParallelism} luong, DPI={OutputDpi}).");

        var parallelResults = new Base64PageEntry[pageCount];
        var parallelLogs = new string[pageCount];
        var completedPages = 0;

        Parallel.For(0, pageCount, new ParallelOptions
        {
            MaxDegreeOfParallelism = effectiveParallelism,
            CancellationToken = cancellationToken
        }, () => PdfDocument.Load(pdfPath), (pageIndex, _, localPdfDocument) =>
            {
                cancellationToken.ThrowIfCancellationRequested();

                using var renderedBitmap = RenderPdfPage(localPdfDocument, pageIndex);
                using var processedBitmap = ProcessBitmap(renderedBitmap);

                var pngBytes = EncodePng(processedBitmap);
                var imageFileName = $"page_{pageIndex + 1:D4}.png";
                var base64 = Convert.ToBase64String(pngBytes);
                var pageKey = $"page_{pageIndex + 1:D4}";
                parallelResults[pageIndex] = new Base64PageEntry(pageIndex + 1, pageKey, imageFileName, base64);

                var done = Interlocked.Increment(ref completedPages);
                onProgress?.Invoke(new ConversionPageProgress(done, pageCount, parallelLogs[pageIndex]));

                return localPdfDocument;
            },
            localPdfDocument => localPdfDocument.Dispose());

        for (var i = 0; i < pageCount; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (parallelResults[i] is null)
            {
                throw new InvalidOperationException($"Khong tao duoc du lieu Base64 cho trang {i + 1}.");
            }

            logs.Add(parallelLogs[i]);
        }

        WriteBase64Entries(base64File, parallelResults, cancellationToken);

        return new ConversionResult(outputDir, base64File, logs);
    }

    private static void WriteBase64Entries(string base64File, IReadOnlyList<Base64PageEntry> entries, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        using var stream = File.Create(base64File);
        using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true });

        writer.WriteStartArray();
        foreach (var entry in entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            JsonSerializer.Serialize(writer, entry);
        }

        writer.WriteEndArray();
        writer.Flush();
    }

    private static Bitmap RenderPdfPage(PdfDocument pdfDocument, int pageIndex)
    {
        var pageSize = pdfDocument.PageSizes[pageIndex];
        var width = Math.Max(1, (int)Math.Round(pageSize.Width / 72f * OutputDpi));
        var height = Math.Max(1, (int)Math.Round(pageSize.Height / 72f * OutputDpi));
        return (Bitmap)pdfDocument.Render(pageIndex, width, height, OutputDpi, OutputDpi, PdfRenderFlags.Annotations);
    }

    private static byte[] EncodePng(Bitmap bitmap)
    {
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return stream.ToArray();
    }

    private static Bitmap ProcessBitmap(Bitmap sourceBitmap)
    {
        var cropWidth = Math.Max(1, sourceBitmap.Width / 4);
        var cropX = sourceBitmap.Width - cropWidth;
        var cropHeight = Math.Max(1, sourceBitmap.Height / 2);
        var cropY = sourceBitmap.Height - cropHeight;

        var outputWidth = Math.Max(1, (int)Math.Round(cropWidth * OutputScale));
        var outputHeight = Math.Max(1, (int)Math.Round(cropHeight * OutputScale));

        var outputBitmap = new Bitmap(outputWidth, outputHeight, PixelFormat.Format24bppRgb);

        using var graphics = Graphics.FromImage(outputBitmap);
        graphics.CompositingQuality = CompositingQuality.HighSpeed;
        graphics.InterpolationMode = InterpolationMode.Bilinear;
        graphics.SmoothingMode = SmoothingMode.None;
        graphics.PixelOffsetMode = PixelOffsetMode.HighSpeed;
        graphics.Clear(Color.White);

        var destinationRect = new Rectangle(0, 0, outputWidth, outputHeight);
        var sourceRect = new Rectangle(cropX, cropY, cropWidth, cropHeight);
        graphics.DrawImage(sourceBitmap, destinationRect, sourceRect, GraphicsUnit.Pixel);

        return outputBitmap;
    }
}

