# Mass Slide Generator - Hệ Thống Soạn Lễ Tự Động

**Mass Slide Generator** là công cụ hỗ trợ soạn bài giảng PowerPoint cho Thánh Lễ Công Giáo. Phần mềm tự động hóa quy trình tìm kiếm lời bài hát, lấy bài đọc Lời Chúa hàng ngày và định dạng slide trình chiếu một cách chuyên nghiệp.

![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)
![Selenium](https://img.shields.io/badge/Selenium-Automation-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

## 🌟 Tính Năng Nổi Bật

### 1. Tự Động Hóa Dữ Liệu
* **Bài đọc Lời Chúa:** Tự động truy cập *ktcgkpv.org* để lấy các bài đọc trong ngày (Ca Nhập Lễ, Đáp Ca, Tung Hô Tin Mừng, Hiệp Lễ...).
* **Tìm kiếm Bài hát:** Tích hợp công cụ tìm kiếm Google để quét dữ liệu từ *thanhcavietnam.net*.

### 2. Xử Lý Văn Bản Thông Minh (Smart Regex Parsing)
* **Lọc Rác HTML:** Tự động loại bỏ các thành phần thừa như nút "PDF", "MP3", "Encore", các dòng quảng cáo "View more", "Copyright".
* **Phân Tách Khổ:** Nhận diện thông minh các đoạn **Điệp Khúc (ĐK)**, **Phiên Khúc (1, 2, 3...)**, **Kết/Coda**.
* **Hỗ Trợ Đa Dạng:** Xử lý tốt cả các bài hát có đánh số thứ tự (1., 2.) và các bài hát **không đánh số** (phân tách dựa trên cấu trúc đoạn).

### 3. Logic Tạo Slide Chuẩn Phụng Vụ
* **Xếp Slide Tự Động:** Tự động chèn **Điệp Khúc** lặp lại sau mỗi **Phiên Khúc** (Người dùng không cần copy paste thủ công).
* **Giao Diện:** Slide nền xanh đậm, chữ trắng, tiêu đề vàng, có kẻ ngang phân cách (Template chuẩn).
* **Ngắt Đoạn:** Tự động chèn slide đen ngăn cách giữa các phần lễ.

## 🛠 Yêu Cầu Hệ Thống

* **Hệ điều hành:** Windows 10/11.
* **Trình duyệt:** Google Chrome (Bắt buộc để Selenium hoạt động).
* **Kết nối Internet:** Cần thiết để tải dữ liệu.

## 📦 Cài Đặt (Dành cho Developer)

Nếu bạn muốn chạy từ mã nguồn Python:

1.  **Clone dự án:**
    ```bash
    git clone [https://github.com/username-cua-ban/mass-slide-generator.git](https://github.com/username-cua-ban/mass-slide-generator.git)
    cd mass-slide-generator
    ```

2.  **Cài đặt thư viện:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Chạy ứng dụng:**
    ```bash
    python main.py
    ```

## 🚀 Hướng Dẫn Sử Dụng
2.  **Bước 1 - Chạy File phần mềm:** mở thư mục dist và chạy main.exe 
Nếu bị Trường hợp 1: Bị Windows Defender chặn (Màn hình xanh "Windows protected your PC")

Bấm vào chữ "More info" (Thông tin thêm).

Chọn nút "Run anyway" (Vẫn chạy).

Trường hợp 2: Bị Antivirus xóa file

Vào lịch sử bảo vệ của Antivirus, chọn file đó và bấm Restore (Khôi phục) hoặc Allow on device (Cho phép trên thiết bị).

Thêm thư mục chứa file vào danh sách loại trừ (Exclusion list).

Trường hợp 3: file chạy thành công !
2.  **Bước 2 - Chọn Cấu Trúc:** Tại màn hình chính, tích chọn các phần lễ muốn soạn (VD: Nhập Lễ, Dâng Lễ, Hiệp Lễ...).
3.  **Bước 3 - Tìm & Soạn Thảo:**
    * Nhập tên bài hát và nhấn **Tìm kiếm**.
    * Chọn kết quả từ danh sách và nhấn **Lấy nội dung**.
    * Phần mềm sẽ tự tách đoạn. Tích chọn các phiên khúc muốn sử dụng.
    * Nhấn **Xác nhận & Tiếp** để sang phần tiếp theo.
4.  **Bước 4 - Xuất File:** Sau khi hoàn tất các phần, nhấn **Xuất File PPTX** và chọn nơi lưu.

## 🔨 Đóng Gói File EXE

Để tạo file `.exe` chạy độc lập (không cần cài Python) có kèm icon:

```bash
pyinstaller --noconsole --onefile --icon=icon.ico main.py