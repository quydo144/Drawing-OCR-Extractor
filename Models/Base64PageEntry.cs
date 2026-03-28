namespace OcrPdf.Models;

public sealed record Base64PageEntry(int PageNumber, string PageKey, string ImageFileName, string ImageBase64);
