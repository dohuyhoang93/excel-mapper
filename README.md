# Excel Data Mapper

A powerful application to map and transfer data between Excel files while preserving formatting and styles.

---

## ✨ Key Features

- **Flexible Column Mapping**: Automatically suggests and allows manual mapping between source and destination columns.
- **Preserve Formatting**: Fully maintains the format, style, color, and borders of the destination Excel file.
- **Merged Cell Handling**: Intelligently supports reading from and writing to merged cells.
- **Data Sorting**: Allows sorting data by a specified column before the transfer.
- **Save/Load Configuration**: Saves mapping configurations to a JSON file for reuse.
- **User-Friendly Interface**: Uses ttkbootstrap with switchable light/dark themes.
- **Comprehensive Error Handling**: Provides clear error messages and automatic backups for the destination file.
- **Robust Validation**: Checks mapping validity (e.g., avoids duplicate destination columns) before execution.

### New Features & Improvements
- **Advanced File Handle Management**: Integrated mechanism to automatically detect and notify if an Excel file is locked by another program (e.g., Microsoft Excel), prompting the user to close it to prevent errors.
- **Automatic Memory Release**: Proactively releases resources after each file read/write operation to increase stability and reduce the risk of the application hanging.
- **Improved Header Reading Logic**: Ensures complex, multi-row headers are read correctly and nameless columns are discarded.
- **Improved Data Writing Logic**: Fixed a bug that could overwrite the destination file's header when it contained vertically merged cells.

## 🏗️ Project Structure (Actual)

The project structure has been streamlined, with most of the logic consolidated into `app.py` to create a self-contained, easily packageable application.

```
excel_mapper/
├── app.py                   # Main file containing the GUI and all application logic
├── logic/
│   └── parser.py            # Helper module for parsing Excel files
├── configs/                 # Default directory for saving .json configuration files
├── requirements.txt         # List of required Python libraries
├── setup.py                 # Script to build the application into an .exe file
├── build.bat                # Convenience script for Windows to run the build
├── icon.ico                 # Application icon
└── README.md                # This documentation file
```

## 🚀 Installation and Execution

### System Requirements
- Windows 10 or later
- Python 3.9+
- Microsoft Office is not required

### Method 1: Run from Source Code

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd excel_mapper
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Run the application:**
    ```bash
    python app.py
    ```

### Method 2: Build the Executable (.exe)

1.  **Automatic build (Windows):**
    Run the `build.bat` file.
    ```bash
    build.bat
    ```

2.  **Or, build manually:**
    ```bash
    python setup.py build
    ```

3.  **The executable will be created at:** `dist/ExcelDataMapper.exe`

## 📖 User Guide

### Step 1: Select Files
- **Source File**: Choose the Excel file containing the data you want to transfer.
- **Destination File**: Choose the Excel template file you want to populate. The formatting of this file will be preserved.

### Step 2: Configure Headers (Important!)
This step tells the application where the headers are located in each file.

- **Source Header Rows**: The row(s) containing the headers in the source file.
- **Destination Header Rows**: The row(s) containing the headers in the destination file.

Press **"Load Columns"** after configuring to have the application read and display the columns.

**ASCII Art Example:**

Assume your **Source** file has a simple header on the first row:
```
Source File (source.xlsx)
+---+--------------+----------+------------+
|   |      A       |    B     |     C      |
+---+--------------+----------+------------+
| 1 |  Employee ID  |  Amount  |    Date    |  <-- Header is on row 1
+---+--------------+----------+------------+
| 2 |    NV001     |   5000   | 2025-07-30 |
+---+--------------+----------+------------+
```
=> Configuration: `Source Header Rows: From [1] To [1]`

Assume your **Destination** file has a complex structure, with the header spanning from row 9 to 10:
```
Destination File (template.xlsx)
... (rows above are ignored)
+---+---------------------+----------------------+
|   |          C          |          D           |
+---+---------------------+----------------------+
| 8 |                     |                      |
+---+---------------------+----------------------+
| 9 |     INFORMATION     |       DETAILS        |  <-- Header starts on row 9
+---+---------------------+----------------------+
| 10|       Emp. ID       |        Amount        |  <-- Header ends on row 10
+---+---------------------+----------------------+
| 11| (data will go here) | (data will go here)  |
+---+---------------------+----------------------+
```
=> Configuration: `Destination Header Rows: From [9] To [10]`

### Step 3: Map Columns
- After clicking "Load Columns", the app will display columns from the source file on the left and from the destination file on the right.
- The system will automatically suggest mappings (e.g., "Amount" -> "Amount").
- You can change these suggestions using the dropdown menu for each source column.

### Step 4: Configure Sorting (Optional)
- In the "Sort Configuration" section, you can select a column from the **source** file to sort the data before it's written to the destination file.

### Step 5: Save/Load Configuration
- **Save Configuration**: Saves all current settings (file paths, headers, mappings) to a `.json` file.
- **Load Configuration**: Loads a previously saved configuration file to avoid re-configuring everything.

### Step 6: Execute
- Press **"Execute Transfer"** to start the data transfer process.
- A progress bar will show the status.
- On success, a notification will appear asking if you want to open the destination folder.

## ⚙️ Advanced Configuration

### JSON Configuration File
You can view and edit the saved configuration file.
```json
{
  "source_file": "C:/path/to/source.xlsx",
  "dest_file": "C:/path/to/destination.xlsx",
  "source_header_start_row": 1,
  "source_header_end_row": 1,
  "dest_header_start_row": 9,
  "dest_header_end_row": 10,
  "sort_column": "Amount",
  "mapping": {
    "Employee ID": "Emp. ID",
    "Amount": "Amount",
    "Date": ""
  },
  "created_date": "2025-07-30T10:30:00"
}
```

## 🔧 Troubleshooting

### Common Errors

1.  **"Could not load columns"**
    -   **Primary Cause**: The header row configuration (Step 2) is incorrect. Double-check your Excel files.
    -   Verify the file paths.
    -   Ensure the file is not locked (open in Microsoft Excel). The application will try to warn you about this.

2.  **"Duplicate destination columns detected"**
    -   You have mapped multiple source columns to the same destination column. Each destination column can only receive data from a single source column.

3.  **"Transfer failed"**
    -   The destination file might be open or locked.
    -   Check write permissions in the destination folder.
    -   Review the `app.log` file for detailed technical reasons.

### Log File
All operations and errors are logged to the `app.log` file in the same directory as the application.
```
2025-07-30 11:00:15,123 - ERROR - File locked by processes: EXCEL.EXE
```

## 🤝 Contributing

### Reporting Bugs
1.  Open an "Issue" on the project's GitHub page.
2.  Attach the `app.log` file if possible.
3.  Describe the steps to reproduce the error in detail.

### Development
1.  Fork the repository.
2.  Create a new feature branch.
3.  Commit your changes with clear messages.
4.  Create a Pull Request.

## 📝 License

APACHE 2.0 License.

---

**Developed by**: Do Huy Hoang
**Version**: 1.1.0
**Last Updated**: 2025-07-30

---
<br>

# (Tiếng Việt) Excel Data Mapper

Một ứng dụng mạnh mẽ để ánh xạ và chuyển dữ liệu giữa các file Excel trong khi vẫn giữ nguyên định dạng và style.

## ✨ Tính năng chính

- **Ánh xạ cột linh hoạt**: Tự động gợi ý và cho phép ánh xạ thủ công giữa cột nguồn và đích.
- **Giữ nguyên định dạng**: Bảo toàn hoàn toàn format, style, màu sắc, viền của file Excel đích.
- **Xử lý merge cells**: Hỗ trợ đọc và ghi dữ liệu vào các ô đã được merge một cách thông minh.
- **Sắp xếp dữ liệu**: Cho phép sắp xếp dữ liệu theo cột được chỉ định trước khi chuyển.
- **Lưu/Tải cấu hình**: Lưu cấu hình ánh xạ vào file JSON để tái sử dụng.
- **Giao diện thân thiện**: Sử dụng ttkbootstrap với 2 theme (sáng/tối) có thể chuyển đổi.
- **Xử lý lỗi toàn diện**: Báo lỗi rõ ràng và có backup tự động cho file đích.
- **Validation mạnh mẽ**: Kiểm tra tính hợp lệ của ánh xạ (tránh trùng lặp cột đích) trước khi thực hiện.

### Tính năng mới & Cải tiến
- **Quản lý File Handle nâng cao**: Tích hợp cơ chế tự động phát hiện và thông báo nếu file Excel đang bị khóa bởi một chương trình khác (ví dụ: Microsoft Excel), yêu cầu người dùng đóng lại để tránh lỗi.
- **Tự động giải phóng bộ nhớ**: Chủ động giải phóng tài nguyên sau mỗi thao tác đọc/ghi file để tăng tính ổn định và giảm thiểu rủi ro treo ứng dụng.
- **Cải thiện logic đọc header**: Đảm bảo đọc được các header phức tạp trên nhiều dòng và loại bỏ các cột không có tên.
- **Cải thiện logic ghi dữ liệu**: Sửa lỗi ghi đè lên header của file đích khi header có các ô được merge theo chiều dọc.

## 🏗️ Cấu trúc dự án (Thực tế)

Cấu trúc dự án đã được tinh gọn, với phần lớn logic được tập trung trong `app.py` để tạo thành một ứng dụng độc lập, dễ đóng gói.

```
excel_mapper/
├── app.py                   # File chính chứa GUI và toàn bộ logic ứng dụng
├── logic/
│   └── parser.py            # Module hỗ trợ phân tích file Excel
├── configs/                 # Thư mục mặc định lưu các file cấu hình .json
├── requirements.txt         # Danh sách các thư viện Python cần thiết
├── setup.py                 # Script để build ứng dụng thành file .exe
├── build.bat                # Script tiện ích cho Windows để chạy build
├── icon.ico                 # Icon của ứng dụng
└── README.md                # Tài liệu hướng dẫn này
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

### Cách 2: Build file thực thi (.exe)

1. **Tự động build (Windows):**
Chạy file `build.bat`.
```bash
build.bat
```

2. **Hoặc build thủ công:**
```bash
python setup.py build
```

3. **File thực thi sẽ được tạo tại:** `dist/ExcelDataMapper.exe`

## 📖 Hướng dẫn sử dụng

### Bước 1: Chọn File
- **Source File**: Chọn file Excel chứa dữ liệu bạn muốn chuyển đi.
- **Destination File**: Chọn file Excel mẫu (template) mà bạn muốn điền dữ liệu vào. Định dạng của file này sẽ được giữ nguyên.

### Bước 2: Cấu hình Header (Quan trọng!)
Đây là bước để chỉ cho ứng dụng biết đâu là dòng tiêu đề trong mỗi file.

- **Source Header Rows**: Các dòng chứa tiêu đề trong file nguồn.
- **Destination Header Rows**: Các dòng chứa tiêu đề trong file đích.

Nhấn **"Load Columns"** sau khi cấu hình xong để ứng dụng đọc và hiển thị các cột.

**Ví dụ minh họa:**

Giả sử file **Source** của bạn có tiêu đề đơn giản ở dòng đầu tiên:

```
Source File (source.xlsx)
+---+--------------+----------+------------+
|   |      A       |    B     |     C      |
+---+--------------+----------+------------+
| 1 |  Mã nhân viên |  Số tiền |  Ngày chi  |  <-- Header ở dòng 1
+---+--------------+----------+------------+
| 2 |    NV001     |   5000   | 2025-07-30 |
+---+--------------+----------+------------+
```
=> Cấu hình: `Source Header Rows: From [1] To [1]`

Giả sử file **Destination** của bạn có cấu trúc phức tạp, tiêu đề nằm từ dòng 9 đến dòng 10:
```
Destination File (template.xlsx)
... (các dòng trên bị bỏ qua)
+---+---------------------+----------------------+
|   |          C          |          D           |
+---+---------------------+----------------------+
| 8 |                     |                      |
+---+---------------------+----------------------+
| 9 |     THÔNG TIN       |     CHI TIẾT         |  <-- Header bắt đầu từ dòng 9
+---+---------------------+----------------------+
| 10|      Mã NV          |      Amount          |  <-- Header kết thúc ở dòng 10
+---+---------------------+----------------------+
| 11| (dữ liệu sẽ vào đây) | (dữ liệu sẽ vào đây) |
+---+---------------------+----------------------+
```
=> Cấu hình: `Destination Header Rows: From [9] To [10]`

### Bước 3: Ánh xạ cột
- Sau khi nhấn "Load Columns", ứng dụng sẽ hiển thị các cột từ file nguồn bên trái và các cột từ file đích bên phải.
- Hệ thống sẽ tự động gợi ý ánh xạ (ví dụ: "Số tiền" -> "Amount").
- Bạn có thể thay đổi các gợi ý này bằng cách chọn từ danh sách dropdown cho mỗi cột nguồn.

### Bước 4: Cấu hình sắp xếp (Tùy chọn)
- Trong phần "Sort Configuration", bạn có thể chọn một cột từ file **nguồn** để sắp xếp dữ liệu trước khi ghi vào file đích.

### Bước 5: Lưu/Tải cấu hình
- **Save Configuration**: Lưu lại toàn bộ cài đặt hiện tại (đường dẫn file, header, ánh xạ) ra một file `.json`.
- **Load Configuration**: Tải lại một file cấu hình đã lưu để không phải chọn lại từ đầu.

### Bước 6: Thực hiện
- Nhấn **"Execute Transfer"** để bắt đầu quá trình chuyển dữ liệu.
- Thanh tiến trình sẽ cập nhật trạng thái.
- Nếu thành công, một thông báo sẽ hiện ra và hỏi bạn có muốn mở thư mục chứa file đích không.

## ⚙️ Cấu hình nâng cao

### File cấu hình JSON
Bạn có thể xem và chỉnh sửa file cấu hình đã lưu.
```json
{
  "source_file": "C:/path/to/source.xlsx",
  "dest_file": "C:/path/to/destination.xlsx",
  "source_header_start_row": 1,
  "source_header_end_row": 1,
  "dest_header_start_row": 9,
  "dest_header_end_row": 10,
  "sort_column": "Số tiền",
  "mapping": {
    "Mã nhân viên": "Mã NV",
    "Số tiền": "Amount",
    "Ngày chi": ""
  },
  "created_date": "2025-07-30T10:30:00"
}
```

## 🔧 Xử lý sự cố

### Lỗi thường gặp

1.  **"Could not load columns"**
    -   **Nguyên nhân chính**: Cấu hình dòng header (Bước 2) không chính xác. Hãy kiểm tra lại file Excel của bạn.
    -   Kiểm tra lại đường dẫn file.
    -   Đảm bảo file không bị khóa (đang mở trong Microsoft Excel). Ứng dụng sẽ cố gắng cảnh báo bạn về điều này.

2.  **"Duplicate destination columns detected"**
    -   Bạn đã ánh xạ nhiều cột nguồn vào cùng một cột đích. Mỗi cột đích chỉ được nhận dữ liệu từ một cột nguồn duy nhất.

3.  **"Transfer failed"**
    -   File đích có thể đang mở hoặc bị khóa.
    -   Kiểm tra quyền ghi file trong thư mục đích.
    -   Xem log chi tiết trong `app.log` để biết nguyên nhân kỹ thuật.

### Log file
Tất cả các hoạt động và lỗi đều được ghi vào file `app.log` trong cùng thư mục với ứng dụng.
```
2025-07-30 11:00:15,123 - ERROR - File locked by processes: EXCEL.EXE
```

## 🤝 Đóng góp

### Báo lỗi
1. Mở một "Issue" trên trang GitHub của dự án.
2. Đính kèm file `app.log` nếu có thể.
3. Mô tả chi tiết các bước để tái hiện lỗi.

### Phát triển
1. Fork repository.
2. Tạo một feature branch mới.
3. Commit các thay đổi với message rõ ràng.
4. Tạo một Pull Request.

## 📝 License

APACHE 2.0 License.

---

**Phát triển bởi**: Do Huy Hoang
**Phiên bản**: 1.1.0
**Ngày cập nhật**: 2025-07-30