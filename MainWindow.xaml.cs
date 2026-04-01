using System.IO;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Threading;
using Forms = System.Windows.Forms;
using Microsoft.Win32;
using DrawingOcrExtractor.Models;
using DrawingOcrExtractor.Services;

namespace DrawingOcrExtractor;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : System.Windows.Window
{
    private const int OcrGeminiBatchSize = 10;
    private static readonly string SettingsDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Drawing OCR Extractor");
    private static readonly string SettingsFilePath = Path.Combine(SettingsDirectory, "ui-settings.json");
    private static readonly string DefaultExcelOutputDirectory = Path.Combine(AppContext.BaseDirectory, "Result");
    private readonly PdfProcessingService _pdfProcessingService;
    private readonly OllamaService _ollamaService;
    private readonly GeminiNormalizationService _geminiNormalizationService;
    private readonly ExcelExportService _excelExportService;
    private readonly Stopwatch _runStopwatch;
    private readonly DispatcherTimer _runTimer;
    private bool _isRestoringGeminiApiKey;
    private string? _pdfPath;
    private string? _excelOutputDirectory;
    private CancellationTokenSource? _runCts;

    public MainWindow()
    {
        _pdfProcessingService = new PdfProcessingService();
        _ollamaService = new OllamaService();
        _geminiNormalizationService = new GeminiNormalizationService();
        _excelExportService = new ExcelExportService();
        _runStopwatch = new Stopwatch();
        _runTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1)
        };
        _runTimer.Tick += RunTimer_Tick;
        InitializeComponent();
        LoadUiSettingsFromLocalStore();
        RefreshExcelOutputFolderText();
    }

    private void GeminiApiKeyPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (_isRestoringGeminiApiKey)
        {
            return;
        }

        SaveGeminiApiKeyToLocalStore(GeminiApiKeyPasswordBox.Password.Trim());
    }

    private void SelectPdfButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Chọn file PDF",
            Filter = "PDF files (*.pdf)|*.pdf",
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        _pdfPath = dialog.FileName;
        PdfPathTextBox.Text = _pdfPath;
        AppendLine($"Đã chọn: {_pdfPath}");
    }

    private void SelectExcelOutputFolderButton_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new Forms.FolderBrowserDialog
        {
            Description = "Chọn thư mục lưu file Excel",
            UseDescriptionForTitle = true,
            InitialDirectory = string.IsNullOrWhiteSpace(_excelOutputDirectory)
                ? DefaultExcelOutputDirectory
                : _excelOutputDirectory
        };

        if (dialog.ShowDialog() != Forms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedPath))
        {
            return;
        }

        _excelOutputDirectory = dialog.SelectedPath;
        RefreshExcelOutputFolderText();
        SaveGeminiApiKeyToLocalStore(GeminiApiKeyPasswordBox.Password.Trim());
        AppendLine($"Đã chọn thư mục lưu Excel: {_excelOutputDirectory}");
    }

    private async void ConvertButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_pdfPath) || !File.Exists(_pdfPath))
        {
            MessageBox.Show("Vui lòng chọn file PDF hợp lệ trước.", "Thiếu file", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var geminiApiKey = GeminiApiKeyPasswordBox.Password.Trim();
        if (string.IsNullOrWhiteSpace(geminiApiKey))
        {
            MessageBox.Show("Vui lòng nhập Gemini API Key trước khi bấm Start.", "Thiếu Gemini API Key", MessageBoxButton.OK, MessageBoxImage.Warning);
            GeminiApiKeyPasswordBox.Focus();
            return;
        }

        ConvertButton.IsEnabled = false;
        CancelButton.IsEnabled = true;
        SelectPdfButton.IsEnabled = false;
        OutputTextBox.Clear();
        ResetProgressUi();
        StartRunTimer();
        _runCts = new CancellationTokenSource();

        try
        {
            AppendLine("Đang xử lý file PDF, vui lòng chờ...");
            var cancellationToken = _runCts.Token;

            var result = await Task.Run(
                () => _pdfProcessingService.ConvertPdfPagesToImagesAndBase64(_pdfPath, UpdateConversionProgress, cancellationToken),
                cancellationToken);

            AppendLine($"Đã hoàn tất chuyển đổi PDF. Output: {result.Base64OutputFile}");
            SetPdfStageComplete();

            AppendLine("Đã hoàn tất bước chuyển đổi PDF sang Base64.");
            AppendLine(string.Empty);
            AppendLine($"Đang xử lý theo batch {OcrGeminiBatchSize}: OCR Ollama -> Gemini -> batch tiếp theo...");
            StartOllamaStage();
            var pages = await LoadBase64PagesAsync(result.Base64OutputFile, cancellationToken);
            var totalPages = pages.Count;
            var processedPages = 0;

            var rows = new List<OllamaExcelRow>();
            string? excelPath = null;
            var exportTimestamp = DateTime.Now.ToString("yyyyMMddHHmmss");

            foreach (var chunk in pages.Chunk(OcrGeminiBatchSize))
            {
                cancellationToken.ThrowIfCancellationRequested();

                var batchPages = chunk.OrderBy(p => p.PageNumber).ToList();
                var firstPage = batchPages.First().PageNumber;
                var lastPage = batchPages.Last().PageNumber;
                AppendLine($"OCR batch page {firstPage}-{lastPage} ({batchPages.Count} trang)...");

                var ocrResult = await _ollamaService.ProcessPagesAsync(
                    batchPages,
                    AppendLine,
                    (completedInBatch, _, message) =>
                    {
                        UpdateOllamaProgress(processedPages + completedInBatch, totalPages, message);
                    },
                    cancellationToken);

                processedPages += batchPages.Count;

                if (ocrResult.FailedRows.Count > 0)
                {
                    rows.AddRange(ocrResult.FailedRows);
                }

                if (ocrResult.OcrBlocks.Count > 0)
                {
                    AppendLine($"Gửi request Gemini cho OCR batch page {firstPage}-{lastPage}...");
                    var normalizedRows = await _geminiNormalizationService.NormalizeOcrBlocksAsync(
                        ocrResult.OcrBlocks,
                        geminiApiKey,
                        AppendLine,
                        cancellationToken);
                    rows.AddRange(normalizedRows);

                    if (normalizedRows.Count > 0)
                    {
                        excelPath = _excelExportService.Export(_pdfPath, rows, exportTimestamp, _excelOutputDirectory);
                        AppendLine($"Đã cập nhật Excel sau batch: +{normalizedRows.Count} dòng, tổng {rows.Count} dòng.");
                        AppendLine($"File Excel hiện tại: {excelPath}");
                    }
                }
                else
                {
                    AppendLine($"Batch page {firstPage}-{lastPage}: không có OCR block hợp lệ để gửi Gemini.");
                }
            }

            if (rows.Count == 0)
            {
                AppendLine("Không có dữ liệu hợp lệ để xuất Excel.");
            }
            else
            {
                excelPath ??= _excelExportService.Export(_pdfPath, rows, exportTimestamp, _excelOutputDirectory);
                AppendLine($"Đã ghi {rows.Count} dòng vào Excel.");
                AppendLine($"Đã xuất Excel: {excelPath}");
            }

            AppendLine("Đã hoàn tất OCR và normalize dữ liệu.");

            try
            {
                if (File.Exists(result.Base64OutputFile))
                {
                    File.Delete(result.Base64OutputFile);
                    AppendLine("Đã xóa file tạm pages_base64.json.");
                }
            }
            catch (Exception ex)
            {
                AppendLine($"Không thể xóa pages_base64.json: {ex.Message}");
            }

            AppendLine($"Tổng thời gian chạy: {FormatElapsed(_runStopwatch.Elapsed)}");
            SetProgressUiComplete();
        }
        catch (OperationCanceledException)
        {
            AppendLine("Đã dừng theo yêu cầu người dùng.");
            ConversionStatusTextBlock.Text = "Đã hủy";
        }
        catch (Exception ex)
        {
            AppendLine($"Lỗi: {ex.Message}");

            if (ex.Message.Contains("GEMINI_QUOTA_EXCEEDED", StringComparison.Ordinal))
            {
                GeminiApiKeyPasswordBox.Password = string.Empty;
                SaveGeminiApiKeyToLocalStore(string.Empty);
                MessageBox.Show(
                    "Gemini API đã hết quota (HTTP 429 Too Many Requests). Đã xóa API key cũ, vui lòng thay Gemini API key khác.",
                    "Hết quota Gemini",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            else
            {
                MessageBox.Show(ex.Message, "Lỗi chuyển đổi", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            ConversionStatusTextBlock.Text = "Đã lỗi";
        }
        finally
        {
            StopRunTimer();
            _runCts?.Dispose();
            _runCts = null;
            ConvertButton.IsEnabled = true;
            CancelButton.IsEnabled = false;
            SelectPdfButton.IsEnabled = true;
        }
    }

    private static async Task<List<Base64PageEntry>> LoadBase64PagesAsync(string base64JsonPath, CancellationToken cancellationToken)
    {
        var json = await File.ReadAllTextAsync(base64JsonPath, Encoding.UTF8, cancellationToken);
        var pages = JsonSerializer.Deserialize<List<Base64PageEntry>>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

        if (pages is null || pages.Count == 0)
        {
            throw new InvalidOperationException("File Base64 rỗng hoặc sai định dạng.");
        }

        return pages.OrderBy(p => p.PageNumber).ToList();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        if (_runCts is null || _runCts.IsCancellationRequested)
        {
            return;
        }

        var confirm = MessageBox.Show(
            "Bạn có chắc muốn hủy tác vụ đang chạy không?",
            "Xác nhận hủy",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        _runCts.Cancel();
        CancelButton.IsEnabled = false;
        ConversionStatusTextBlock.Text = "Đang hủy...";
        AppendLine("Đang gửi yêu cầu hủy...");
    }

    private void UpdateConversionProgress(ConversionPageProgress progress)
    {
        if (Dispatcher.CheckAccess())
        {
            ApplyConversionProgress(progress);
            return;
        }

        Dispatcher.BeginInvoke(() => ApplyConversionProgress(progress));
    }

    private void ApplyConversionProgress(ConversionPageProgress progress)
    {
        var percent = progress.TotalPages == 0
            ? 0
            : Math.Clamp(progress.CompletedPages * 100.0 / progress.TotalPages, 0, 100);

        ConversionProgressBar.Value = percent;
        ConversionStatusTextBlock.Text = $"Đang xử lý PDF: {progress.CompletedPages}/{progress.TotalPages}";
        if (!string.IsNullOrWhiteSpace(progress.Message))
        {
            AppendLine(progress.Message);
        }
    }

    private void UpdateOllamaProgress(int completedPages, int totalPages, string message)
    {
        if (Dispatcher.CheckAccess())
        {
            ApplyOllamaProgress(completedPages, totalPages, message);
            return;
        }

        Dispatcher.BeginInvoke(() => ApplyOllamaProgress(completedPages, totalPages, message));
    }

    private void ApplyOllamaProgress(int completedPages, int totalPages, string message)
    {
        var percent = totalPages == 0
            ? 0
            : Math.Clamp(completedPages * 100.0 / totalPages, 0, 100);

        ConversionProgressBar.Value = percent;
        ConversionStatusTextBlock.Text = $"Đang gửi Ollama: {completedPages}/{totalPages}";
        AppendLine(message);
    }

    private void ResetProgressUi()
    {
        ConversionProgressBar.Value = 0;
        ConversionStatusTextBlock.Text = "Đang chờ xử lý";
        RunTimeTextBlock.Text = "Thời gian chạy: 00:00";
    }

    private void SetProgressUiComplete()
    {
        ConversionProgressBar.Value = 100;
        ConversionStatusTextBlock.Text = "Hoàn tất";
    }

    private void SetPdfStageComplete()
    {
        ConversionProgressBar.Value = 100;
        ConversionStatusTextBlock.Text = "Hoàn tất xử lý PDF";
    }

    private void StartOllamaStage()
    {
        ConversionProgressBar.Value = 0;
        ConversionStatusTextBlock.Text = "Đang gửi Ollama: 0/0";
    }

    private void StartRunTimer()
    {
        _runStopwatch.Restart();
        RunTimeTextBlock.Text = "Thời gian chạy: 00:00";
        _runTimer.Start();
    }

    private void StopRunTimer()
    {
        if (_runTimer.IsEnabled)
        {
            _runTimer.Stop();
        }

        RunTimeTextBlock.Text = $"Thời gian chạy: {FormatElapsed(_runStopwatch.Elapsed)}";
        _runStopwatch.Stop();
    }

    private void RunTimer_Tick(object? sender, EventArgs e)
    {
        RunTimeTextBlock.Text = $"Thời gian chạy: {FormatElapsed(_runStopwatch.Elapsed)}";
    }

    private static string FormatElapsed(TimeSpan elapsed)
    {
        var totalHours = (int)elapsed.TotalHours;
        return totalHours > 0
            ? $"{totalHours:00}:{elapsed.Minutes:00}:{elapsed.Seconds:00}"
            : $"{elapsed.Minutes:00}:{elapsed.Seconds:00}";
    }

    private void AppendLine(string text)
    {
        if (Dispatcher.CheckAccess())
        {
            OutputTextBox.AppendText(text + Environment.NewLine);
            OutputTextBox.ScrollToEnd();
            return;
        }

        Dispatcher.BeginInvoke(() =>
        {
            OutputTextBox.AppendText(text + Environment.NewLine);
            OutputTextBox.ScrollToEnd();
        });
    }

    private void LoadUiSettingsFromLocalStore()
    {
        try
        {
            if (!File.Exists(SettingsFilePath))
            {
                return;
            }

            var json = File.ReadAllText(SettingsFilePath, Encoding.UTF8);
            var settings = JsonSerializer.Deserialize<UiSettings>(json);
            if (settings is null)
            {
                return;
            }

            _excelOutputDirectory = string.IsNullOrWhiteSpace(settings.ExcelOutputDirectory)
                ? null
                : settings.ExcelOutputDirectory;

            if (!string.IsNullOrWhiteSpace(settings.GeminiApiKeyProtected))
            {
                var apiKey = UnprotectText(settings.GeminiApiKeyProtected);
                _isRestoringGeminiApiKey = true;
                GeminiApiKeyPasswordBox.Password = apiKey;
            }
        }
        catch
        {
            // Ignore invalid saved settings and allow user to re-enter key.
        }
        finally
        {
            _isRestoringGeminiApiKey = false;
            RefreshExcelOutputFolderText();
        }
    }

    private void SaveGeminiApiKeyToLocalStore(string apiKey)
    {
        try
        {
            Directory.CreateDirectory(SettingsDirectory);
            var settings = new UiSettings
            {
                GeminiApiKeyProtected = string.IsNullOrWhiteSpace(apiKey)
                    ? string.Empty
                    : ProtectText(apiKey),
                ExcelOutputDirectory = _excelOutputDirectory ?? string.Empty
            };

            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsFilePath, json, Encoding.UTF8);
        }
        catch
        {
            // Ignore persistence errors and continue normal app flow.
        }
    }

    private static string ProtectText(string plainText)
    {
        var plainBytes = Encoding.UTF8.GetBytes(plainText);
        var protectedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
        return Convert.ToBase64String(protectedBytes);
    }

    private static string UnprotectText(string protectedText)
    {
        var protectedBytes = Convert.FromBase64String(protectedText);
        var plainBytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(plainBytes);
    }

    private sealed class UiSettings
    {
        public string GeminiApiKeyProtected { get; init; } = string.Empty;
        public string ExcelOutputDirectory { get; init; } = string.Empty;
    }

    private void RefreshExcelOutputFolderText()
    {
        if (ExcelOutputFolderTextBox is null)
        {
            return;
        }

        ExcelOutputFolderTextBox.Text = string.IsNullOrWhiteSpace(_excelOutputDirectory)
            ? DefaultExcelOutputDirectory
            : _excelOutputDirectory;
    }
}
