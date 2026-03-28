namespace OcrPdf.Models;

public sealed record ConversionPageProgress(int CompletedPages, int TotalPages, string Message);