namespace DrawingOcrExtractor.Models;

public sealed record OllamaRequest(string Model, string Prompt, string[] Images, bool Stream);

