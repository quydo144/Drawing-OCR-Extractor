namespace DrawingOcrExtractor.Models;

public sealed record OcrExtractionResult(
    IReadOnlyList<OcrPageBlock> OcrBlocks,
    IReadOnlyList<OllamaExcelRow> FailedRows);
