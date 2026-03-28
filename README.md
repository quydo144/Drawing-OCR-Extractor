# Drawing OCR Extractor

Ứng dụng desktop (WPF, .NET 8) dùng để trích xuất dữ liệu bản vẽ từ PDF theo quy trình:

1. PDF -> ảnh PNG (theo từng trang)
2. PNG -> JSON Base64
3. OCR bằng Ollama model `glm-ocr:latest`
4. Chuẩn hóa dữ liệu bằng Gemini
5. Xuất kết quả ra Excel

## Tính năng chính

- Chọn file PDF trực tiếp từ giao diện.
- Hiển thị tiến trình và log theo từng giai đoạn.
- Hỗ trợ hủy tác vụ đang chạy.
- Xử lý OCR theo batch để giảm rủi ro timeout.
- Xuất Excel có trạng thái theo từng trang.
- Lưu Gemini API key dưới dạng mã hóa cục bộ theo user Windows hiện tại.

## Công nghệ sử dụng

- .NET 8 + WPF
- PdfiumViewer (render PDF)
- ClosedXML (xuất Excel)
- Ollama local API
- Gemini API

## Yêu cầu môi trường

- Windows
- .NET SDK 8.0+
- Đã cài Ollama và chạy được trong PATH
- Model Ollama: `glm-ocr:latest`
- Gemini API key hợp lệ

## Cài đặt nhanh

### 1) Cài Ollama và model

Chạy:

```bat
setup_ollama_glm_ocr.bat
```

Script sẽ:

- Kiểm tra Ollama đã có hay chưa.
- Tự cài Ollama bằng winget nếu thiếu.
- Tự pull `glm-ocr:latest` nếu chưa có model.

### 2) Restore và build

```powershell
dotnet restore "Drawing OCR Extractor.csproj"
dotnet build "Drawing OCR Extractor.csproj"
```

## Chạy ứng dụng

```powershell
dotnet run --project "Drawing OCR Extractor.csproj"
```

## Hướng dẫn sử dụng

1. Mở ứng dụng.
2. Bấm Chọn file PDF.
3. Nhập Gemini API key.
4. Bấm Bắt đầu xử lý.
5. Theo dõi log và thanh tiến trình trên giao diện.
6. Sau khi hoàn tất, mở thư mục output để lấy JSON/Excel.

## File đầu ra

Với mỗi PDF đầu vào, ứng dụng tạo một thư mục dạng:

- `<pdf_name>_pages_yyyyMMdd_HHmmss`

Trong thư mục này có:

- `pages_base64.json`
- `ollama_results.xlsx`

Cột trong file Excel:

- Drawing Name
- Drawing No
- Page Number
- Status
- Error Message

## Chi tiết xử lý

- Giai đoạn PDF:
  - Render trang với DPI 170.
  - Cắt vùng 1/4 phía dưới bên phải.
  - Resize còn 50%.
  - Chuyển ảnh PNG sang Base64.
- Giai đoạn OCR:
  - Gửi từng trang lên Ollama local.
  - Retry tối đa 3 lần nếu lỗi.
- Giai đoạn chuẩn hóa:
  - Gửi block OCR theo batch sang Gemini.
  - Retry khi gặp lỗi server tạm thời.
- Giai đoạn xuất file:
  - Gộp dữ liệu và ghi ra Excel.

## Lưu ý quan trọng

- Ollama phải truy cập được tại `http://localhost:11434`.
- Không có Gemini API key thì ứng dụng không bắt đầu xử lý.
- Chất lượng OCR phụ thuộc vùng cắt ở góc phải phía dưới.

## Lỗi thường gặp

### Không kết nối được Ollama

Kiểm tra:

```powershell
ollama --version
ollama list
```

Đảm bảo model `glm-ocr:latest` đã tồn tại.

### Gemini HTTP 401/403

- Kiểm tra lại API key.
- Kiểm tra quota/billing trong project Google AI.

### OCR hoặc Gemini bị timeout

- Thử với PDF ít trang hơn.
- Đóng bớt ứng dụng nặng tài nguyên.
- Kiểm tra kết nối Internet khi gọi Gemini.

## Cấu trúc dự án chính

- `MainWindow.xaml`, `MainWindow.xaml.cs`: giao diện và điều phối luồng chính
- `Services/PdfProcessingService.cs`: PDF -> PNG -> Base64
- `Services/OllamaService.cs`: OCR qua Ollama
- `Services/GeminiNormalizationService.cs`: chuẩn hóa qua Gemini
- `Services/ExcelExportService.cs`: xuất Excel
- `Models/`: các model dữ liệu

## Release bằng GitHub Actions

File workflow:

- `.github/workflows/release.yml`

Cách kích hoạt:

- Chạy thủ công trong tab Actions.
- Push tag bắt đầu bằng `v` (ví dụ: `v1.0.0`).

Kết quả:

- Artifact: `Drawing-OCR-Extractor-win-x64`
- File zip: `Drawing-OCR-Extractor-win-x64.zip`
- Khi chạy bằng tag, workflow tự tạo GitHub Release và đính kèm file zip.

Lệnh tạo release:

```powershell
git tag v1.0.0
git push origin v1.0.0
```
