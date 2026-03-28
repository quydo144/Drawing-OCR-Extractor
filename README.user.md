# Drawing OCR Extractor - Hướng dẫn cho người dùng

Chào bạn, đây là ứng dụng desktop giúp trích xuất thông tin bản vẽ từ file PDF và xuất kết quả ra Excel.
Bạn chỉ cần làm theo các bước bên dưới là có thể dùng ngay.

## Yêu cầu trước khi chạy

- Windows 10/11 64-bit
- Có Gemini API key hợp lệ

## Cách lấy Gemini API key

1. Truy cập trang: https://aistudio.google.com/
2. Chọn **Get API Key**
3. Chọn **Create API Key**
4. Sao chép API key để dùng trong ứng dụng

## Lần đầu tiên chạy ứng dụng

Ở lần chạy đầu tiên, bạn cần chạy file `setup_ollama_glm_ocr.bat` để cài Ollama và model `glm-ocr:latest`:

```bat
setup_ollama_glm_ocr.bat
```

## Cách sử dụng nhanh

1. Mở `Drawing OCR Extractor.exe`.
2. Bấm **Chọn file PDF**.
3. Nhập Gemini API key.
4. Bấm **Bắt đầu xử lý**.
5. Chờ xử lý hoàn tất, sau đó mở thư mục output để xem kết quả.

Mỗi lần chạy, ứng dụng sẽ tạo một thư mục theo dạng:

- `<ten_pdf>_pages_yyyyMMdd_HHmmss`

Trong thư mục này sẽ có:

- `pages_base64.json`
- `ollama_results.xlsx`

Nếu xử lý chưa thành công, hãy kiểm tra lại API key và chắc chắn Ollama đang chạy trên máy của bạn.