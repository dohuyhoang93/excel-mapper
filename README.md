# Excel Data Mapper

Một ứng dụng mạnh mẽ để ánh xạ và chuyển dữ liệu giữa các file Excel trong khi vẫn giữ nguyên định dạng và style.

## ✨ Tính năng chính

- **Ánh xạ cột linh hoạt**: Tự động gợi ý và cho phép ánh xạ thủ công giữa cột nguồn và đích
- **Giữ nguyên định dạng**: Bảo toàn hoàn toàn format, style, màu sắc, viền của file Excel gốc
- **Xử lý merge cells**: Hỗ trợ đọc và xử lý các ô đã được merge
- **Sắp xếp dữ liệu**: Cho phép sắp xếp dữ liệu theo cột được chỉ định trước khi chuyển
- **Lưu/Tải cấu hình**: Lưu cấu hình ánh xạ vào JSON để tái sử dụng
- **Giao diện thân thiện**: Sử dụng ttkbootstrap với 2 theme (sáng/tối)
- **Xử lý lỗi toàn diện**: Báo lỗi rõ ràng và có backup tự động
- **Validation mạnh mẽ**: Kiểm tra tính hợp lệ của ánh xạ trước khi thực hiện

## 🏗️ Cấu trúc dự án

```
excel_mapper/
├── app.py                   # GUI chính
├── config.py                # Cấu hình chung  
├── logic/
│   ├── parser.py            # Phân tích header, xử lý merge
│   ├── mapper.py            # Gợi ý ánh xạ tiêu đề
│   └── transfer.py          # Ghi dữ liệu theo ánh xạ
├── gui/
│   └── widgets.py           # Các thành phần GUI tái sử dụng
├── configs/                 # Thư mục lưu cấu hình JSON
├── requirements.txt         # Dependencies
├── setup.py                 # Build script cho PyInstaller
├── build.bat                # Windows build script
├── icon.ico                 # Icon ứng dụng
└── README.md               # Tài liệu này
```

## 🚀 Cài đặt và chạy

### Yêu cầu hệ thống
- Windows 10 trở lên
- Python 3.9+
- Không cần cài đặt Microsoft Office

### Cách 1: Chạy từ source code

1. **Clone repository:**
```bash
git clone <repository-url>
cd excel_mapper
```

2. **Cài đặt dependencies:**
```bash
pip install -r requirements.txt
```

3. **Chạy ứng dụng:**
```bash
python app.py
```

### Cách 2: Build executable

1. **Tự động build (Windows):**
```bash
build.bat
```

2. **Hoặc build thủ công:**
```bash
python setup.py build
```

3. **Executable sẽ được tạo tại:** `dist/ExcelDataMapper.exe`

## 📖 Hướng dẫn sử dụng

### Bước 1: Chọn file Excel
- **Source File**: File chứa dữ liệu nguồn cần chuyển
- **Destination File**: File template đích (sẽ giữ nguyên format)

### Bước 2: Cấu hình header
- **Source Header Row**: Dòng chứa tiêu đề trong file nguồn (mặc định: 1)
- **Destination Header Row**: Dòng chứa tiêu đề trong file đích (mặc định: 9)
- Nhấn **"Load Columns"** để tải danh sách cột

### Bước 3: Ánh xạ cột
- Hệ thống sẽ tự động gợi ý ánh xạ dựa trên tên cột
- Điều chỉnh ánh xạ thủ công qua dropdown menu
- Kiểm tra **Confidence** score để đánh giá độ tin cậy

### Bước 4: Cấu hình sắp xếp (tùy chọn)
- Chọn cột để sắp xếp dữ liệu trước khi chuyển

### Bước 5: Lưu/Tải cấu hình
- **Save Configuration**: Lưu cấu hình hiện tại
- **Load Configuration**: Tải cấu hình đã lưu

### Bước 6: Thực hiện chuyển dữ liệu
- Nhấn **"Execute Transfer"** để bắt đầu
- Theo dõi tiến trình qua progress bar
- Ứng dụng sẽ tự động mở thư mục chứa file đích khi hoàn thành

## ⚙️ Cấu hình nâng cao

### File cấu hình JSON
```json
{
  "source_file": "C:/path/to/source.xlsx",
  "dest_file": "C:/path/to/destination.xlsx", 
  "source_header_row": 1,
  "destination_header_row": 9,
  "sort_column": "Content",
  "mapping": {
    "Nội dung": "Contents",
    "Mục đích": "Purpose",
    "Số tiền": "Amount"
  },
  "created_date": "2025-01-15T10:30:00"
}
```

### Tùy chỉnh gợi ý ánh xạ
Chỉnh sửa `config.py` để thêm từ khóa gợi ý:

```python
COMMON_MAPPINGS = {
    'content': ['Contents', 'Content', 'Description'],
    'purpose': ['Purpose', 'Reason', 'Use'],
    'amount': ['Amount', 'Value', 'Total'],
    # Thêm mapping tùy chỉnh...
}
```

## 🔧 Xử lý sự cố

### Lỗi thường gặp

1. **"Could not load columns"**
   - Kiểm tra đường dẫn file
   - Đảm bảo dòng header đúng
   - File không bị khóa bởi Excel

2. **"Mapping validation failed"**
   - Kiểm tra cột trùng lặp
   - Đảm bảo cột đích tồn tại

3. **"Transfer failed"**
   - File đích có thể đang mở
   - Kiểm tra quyền ghi file
   - Xem log chi tiết trong `app.log`

### Log file
Tất cả lỗi được ghi vào file `app.log` với format:
```
2025-01-15 10:30:45,123 - ERROR - Transfer failed: File is locked
```

## 🎨 Tùy chỉnh giao diện

### Chuyển đổi theme
- **Menu > Settings > Switch Theme**
- Flatly (sáng) ↔ Superhero (tối)

### Tùy chỉnh theme
Chỉnh sửa trong `app.py`:
```python
self.root = ttk_boot.Window(themename="cosmo")  # Thay đổi theme
```

## 🧪 Testing và Debug

### Chạy với debug mode
```bash
python app.py --debug
```

### Test với file mẫu
1. Tạo file nguồn đơn giản với cột: Name, Amount, Date
2. Tạo file đích với format phức tạp
3. Test ánh xạ và chuyển dữ liệu

## 📊 Performance

### Khuyến nghị
- **File size**: < 50MB cho performance tốt nhất
- **Rows**: < 100,000 dòng
- **Columns**: < 50 cột

### Tối ưu hóa
- Đóng Excel trước khi chạy
- Sử dụng SSD cho tốc độ I/O
- Tắt antivirus scanning cho thư mục làm việc

## 🤝 Đóng góp

### Báo lỗi
1. Mở issue trên GitHub
2. Đính kèm file `app.log`
3. Mô tả chi tiết bước tái hiện

### Phát triển
1. Fork repository
2. Tạo feature branch
3. Commit với message rõ ràng
4. Tạo Pull Request

## 📝 License

MIT License - Xem file LICENSE để biết chi tiết.

## 🆘 Hỗ trợ

- **Email**: support@example.com
- **Issues**: GitHub Issues
- **Documentation**: Wiki page

---

**Phát triển bởi**: Excel Data Mapper Team  
**Phiên bản**: 1.0.0  
**Ngày cập nhật**: 2025-01-15