using ClosedXML.Excel;
using System.IO;
using DrawingOcrExtractor.Models;

namespace DrawingOcrExtractor.Services;

public sealed class ExcelExportService
{
    public string Export(string pdfPath, IReadOnlyList<OllamaExcelRow> rows, string exportTimestamp, string? outputDirectory = null)
    {
        var pdfFileName = Path.GetFileNameWithoutExtension(pdfPath);
        var safePdfFileName = SanitizeFileName(pdfFileName);
        var resultDirectory = string.IsNullOrWhiteSpace(outputDirectory)
            ? Path.Combine(AppContext.BaseDirectory, "Result")
            : outputDirectory;
        Directory.CreateDirectory(resultDirectory);
        var excelPath = Path.Combine(resultDirectory, $"List_{safePdfFileName}_{exportTimestamp}.xlsx");

        using var workbook = File.Exists(excelPath)
            ? new XLWorkbook(excelPath)
            : new XLWorkbook();

        var worksheet = workbook.Worksheets.FirstOrDefault(w => w.Name == "Results")
            ?? workbook.Worksheets.Add("Results");

        EnsureHeader(worksheet);
        ClearDataRows(worksheet);

        for (var i = 0; i < rows.Count; i++)
        {
            var rowIndex = i + 2;
            worksheet.Cell(rowIndex, 1).Value = rows[i].DrawingName;
            worksheet.Cell(rowIndex, 2).Value = rows[i].DrawingNo;
            worksheet.Cell(rowIndex, 3).Value = rows[i].PageNumber;
            worksheet.Cell(rowIndex, 4).Value = rows[i].Status;
            worksheet.Cell(rowIndex, 5).Value = rows[i].ErrorMessage;
        }

        var headerRange = worksheet.Range(1, 1, 1, 5);
        headerRange.Style.Font.Bold = true;
        worksheet.Columns(1, 5).AdjustToContents();

        workbook.SaveAs(excelPath);
        return excelPath;
    }

    private static string SanitizeFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return "unknown";
        }

        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = new string(fileName.Select(c => invalidChars.Contains(c) ? '_' : c).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(sanitized) ? "unknown" : sanitized;
    }

    private static void EnsureHeader(IXLWorksheet worksheet)
    {
        if (!string.IsNullOrWhiteSpace(worksheet.Cell(1, 1).GetString()))
        {
            return;
        }

        worksheet.Cell(1, 1).Value = "Drawing Name";
        worksheet.Cell(1, 2).Value = "Drawing No";
        worksheet.Cell(1, 3).Value = "Page Number";
        worksheet.Cell(1, 4).Value = "Status";
        worksheet.Cell(1, 5).Value = "Error Message";
    }

    private static void ClearDataRows(IXLWorksheet worksheet)
    {
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        if (lastRow < 2)
        {
            return;
        }

        worksheet.Rows(2, lastRow).Delete();
    }
}

