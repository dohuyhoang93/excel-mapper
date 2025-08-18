# Hướng dẫn nâng cấp tính năng cho Excel Data Mapper
## Trước thay đổi:
1. Tự sắp xếp theo Sort by Column (optional)
2. Tự ghi tất cả vào 1 excel sheet

## Sau thay đổi:
### Giao diện:
- Sort By Column không còn là tính năng tự động sắp xếp nữa. Mà đổi thành : "Group by column" , là trường để cung cấp cho python giá trị để lọc, gom các dòng thành từng group
- Bổ sung thêm 1 Combobox để chọn sheet trong destination file làm master sheet.

### Logic hoạt động:
- Tự tìm các dòng có cùng nội dung theo Group by Column. Gom thành các nhóm. Đếm số dòng cần thiết để ghi mỗi nhóm vào 1 excel sheet.
- Lần lượt ghi mỗi nhóm vào từng excel sheet riêng. Trước khi ghi, căn cứ vào số dòng cần thiết và End Write Row để chèn thêm dòng cần thiết vào khoảng giữa Start Write Row - End Write Row

### Ví dụ luồng hoạt động:
1. Các bước chọn source, destination, chỉ định Header , setting write zone không có gì thay đổi.
2. Group by column: Chọn cột để lọc, gom nhóm. Ví dụ chọn cột có tên: **Supplier**. App sẽ tự động lọc các dòng có cùng giá trị trong cột "supplier" ra -> tạo thành 1 nhóm. (nhóm chấp nhận cả trường hợp chỉ có 1 phần tử).
3. Đếm xem nhóm này có bao nhiêu phần tử.
4. Người dùng chỉ định master sheet

Quá trình transfer bắt đầu

5. Tạo 1 bản sao từ master sheet với tên là giá trị trong supplier (hay có thể gọi là tên group).
6. Chuyển đến sheet vừa tạo. Kiểm tra xem write zone có bao nhiêu dòng khả dụng. Nếu nhỏ hơn số phần tử trong group sẽ transfer. Tiến hành tính toán và chèn thêm dòng cần thiết.
7. Tiến hành transfer
8. Tìm dòng có cell chứa giá trị được chọn ở trường "Group by Column". Trong ví dụ này là dòng có cell mang giá trị: Supplier. Chỗ này phải xử lý được trường hợp merge cell, hãy đọc ở cell có giá trị đầu tiên: trên cùng bên trái trong merge cell.
9. Ghi giá trị tên nhóm vào cell bên cạnh cell "Supplier". Cũng phải xử lý đúng nếu là merge cell. Thực chất đây chính là bước điền tên nhà cung cấp.
10. Lặp lại cho các group khác cho đến hết

### Đảm bảo:
- Ghi đầy đủ hết tất cả các dòng trong source file sang destination file.
- Mỗi sheet đặt tên theo nội dung được Group By Column (tự loại bỏ ký tự đặc biệt nếu có, tối đa 33 ký tự theo quy ước của excel).
- Tính năng Preview Transfer cũng cần thay đổi theo để đáp ứng.