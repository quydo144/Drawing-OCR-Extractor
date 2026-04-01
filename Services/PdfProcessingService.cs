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
    private const double A0OutputScale = 0.75;
    private const int MaxParallelism = 10;
    private const int OutputDpi = 170;
    private const int A0OutputDpi = 230;

    public ConversionResult ConvertPdfPagesToImagesAndBase64(
        string pdfPath,
        Action<ConversionPageProgress>? onProgress = null,
        CancellationToken cancellationToken = default)
    {
        var outputDir = AppContext.BaseDirectory;
        var base64File = Path.Combine(AppContext.BaseDirectory, "pages_base64.json");

        using var pdfDocument = PdfDocument.Load(pdfPath);
        var pageCount = pdfDocument.PageCount;
        var configuredParallelism = Math.Max(1, Math.Min(MaxParallelism, Environment.ProcessorCount));
        var effectiveParallelism = Math.Min(configuredParallelism, Math.Max(1, pageCount));

        var parallelResults = new Base64PageEntry[pageCount];
        var completedPages = 0;

        Parallel.For(0, pageCount, new ParallelOptions
        {
            MaxDegreeOfParallelism = effectiveParallelism,
            CancellationToken = cancellationToken
        }, () => PdfDocument.Load(pdfPath), (pageIndex, _, localPdfDocument) =>
            {
                cancellationToken.ThrowIfCancellationRequested();

                var pageSize = localPdfDocument.PageSizes[pageIndex];
                using var renderedBitmap = RenderPdfPage(localPdfDocument, pageIndex, pageSize);
                using var processedBitmap = ProcessBitmap(renderedBitmap, pageSize);

                var pngBytes = EncodePng(processedBitmap);
                var imageFileName = $"page_{pageIndex + 1:D4}.png";
                // var imageFilePath = Path.Combine(outputDir, imageFileName);
                // File.WriteAllBytes(imageFilePath, pngBytes);
                var base64 = Convert.ToBase64String(pngBytes);
                var pageKey = $"page_{pageIndex + 1:D4}";
                parallelResults[pageIndex] = new Base64PageEntry(pageIndex + 1, pageKey, imageFileName, base64);

                var done = Interlocked.Increment(ref completedPages);
                onProgress?.Invoke(new ConversionPageProgress(done, pageCount, string.Empty));

                return localPdfDocument;
            },
            localPdfDocument => localPdfDocument.Dispose());

        for (var i = 0; i < pageCount; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (parallelResults[i] is null)
            {
                throw new InvalidOperationException($"Không tạo được dữ liệu Base64 cho trang {i + 1}.");
            }
        }

        WriteBase64Entries(base64File, parallelResults, cancellationToken);

        return new ConversionResult(outputDir, base64File);
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

    private static Bitmap RenderPdfPage(PdfDocument pdfDocument, int pageIndex, SizeF pageSize)
    {
        var paperSize = DetectPaperSize(pageSize);
        var dpi = paperSize == "A0" ? A0OutputDpi : OutputDpi;
        var width = Math.Max(1, (int)Math.Round(pageSize.Width / 72f * dpi));
        var height = Math.Max(1, (int)Math.Round(pageSize.Height / 72f * dpi));
        return (Bitmap)pdfDocument.Render(pageIndex, width, height, dpi, dpi, PdfRenderFlags.Annotations);
    }

    private static byte[] EncodePng(Bitmap bitmap)
    {
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return stream.ToArray();
    }

    private static Bitmap ProcessBitmap(Bitmap sourceBitmap, SizeF pageSize)
    {
        var paperSize = DetectPaperSize(pageSize);
        int cropWidth, cropHeight, cropX, cropY;

        if (paperSize == "A0")
        {
            // A0: 1/8 ở bên phải và 1/4 từ dưới lên
            cropWidth = Math.Max(1, sourceBitmap.Width / 8);
            cropX = sourceBitmap.Width - cropWidth;
            cropHeight = Math.Max(1, sourceBitmap.Height / 4);
            cropY = sourceBitmap.Height - cropHeight;
        }
        else
        {
            // Default: 1/4 ở bên phải và 1/2 từ dưới lên
            cropWidth = Math.Max(1, sourceBitmap.Width / 4);
            cropX = sourceBitmap.Width - cropWidth;
            cropHeight = Math.Max(1, sourceBitmap.Height / 2);
            cropY = sourceBitmap.Height - cropHeight;
        }

        var scale = paperSize == "A0" ? A0OutputScale : OutputScale;
        var outputWidth = Math.Max(1, (int)Math.Round(cropWidth * scale));
        var outputHeight = Math.Max(1, (int)Math.Round(cropHeight * scale));

        var outputBitmap = new Bitmap(outputWidth, outputHeight, PixelFormat.Format24bppRgb);

        using var graphics = Graphics.FromImage(outputBitmap);
        graphics.CompositingQuality = CompositingQuality.HighQuality;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.SmoothingMode = SmoothingMode.HighQuality;
        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        graphics.Clear(Color.White);

        var destinationRect = new Rectangle(0, 0, outputWidth, outputHeight);
        var sourceRect = new Rectangle(cropX, cropY, cropWidth, cropHeight);
        graphics.DrawImage(sourceBitmap, destinationRect, sourceRect, GraphicsUnit.Pixel);

        return outputBitmap;
    }

    private static string DetectPaperSize(SizeF pageSize)
    {
        const float PointsPerInch = 72f;
        const float InchPerMm = 1f / 25.4f;

        var widthMm = pageSize.Width / PointsPerInch / InchPerMm;
        var heightMm = pageSize.Height / PointsPerInch / InchPerMm;

        var (minDim, maxDim) = widthMm > heightMm
            ? (heightMm, widthMm)
            : (widthMm, heightMm);

        return (minDim, maxDim) switch
        {
            ( >= 206 and <= 214, >= 290 and <= 304) => "A4",
            ( >= 279 and <= 287, >= 407 and <= 421) => "A3",
            ( >= 420 and <= 430, >= 594 and <= 612) => "A2",
            ( >= 594 and <= 604, >= 840 and <= 860) => "A1",
            ( >= 840 and <= 850, > 860) => "A0",
            _ => "Unknown"
        };
    }
}

