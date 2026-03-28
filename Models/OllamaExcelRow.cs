namespace DrawingOcrExtractor.Models;

public sealed record OllamaExcelRow(
    string DrawingName,
    string DrawingNo,
    int PageNumber,
    string Status = "OK",
    string ErrorMessage = "");

