namespace DrawingOcrExtractor.Models;

public sealed record ConversionResult(string OutputImageDirectory, string Base64OutputFile, IReadOnlyList<string> Logs);

