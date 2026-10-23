# BÀI LÀM – THI THỬ HACKATHON 01

- **Mã đề:** Thi thử
- **Họ tên:Nguyễn Đức Hiếu** 
- **Lớp:CNTT2** 
- **Link repository:https://github.com/hn3443723-max/CNTT2_Nhap-mon-CNTT_ThiThu.git** 
- **Hệ điều hành (Windows/macOS):Windows** 

## 4.2. Thao tác máy tính cơ bản bằng CLI

### Các lệnh đã chạy

```shell
# 1. Hiển thị đường dẫn tuyệt đối của thư mục hiện tại
Get-Location

# 2. Liệt kê nội dung, gồm cả file ẩn
Get-ChildItem -Force

# 3. Di chuyển nhap_hang.txt vào docs và file logo vào media
Move-Item nhap_hang.txt docs/
Move-Item logo_van_phong_pham.png media/

# 4. Sao chép logo từ media vào backup
Copy-Item media/logo_van_phong_pham.png backup/

# 5. Xoá file .tmp tại thư mục gốc
Remove-Item "*.tmp"

# 6. Dùng đường dẫn tương đối vào logs, liệt kê nội dung rồi quay lại thư mục gốc
Set-Location logs
Get-ChildItem
Set-Location ..

```

### Kết quả liệt kê toàn bộ dự án (Windows: Get-ChildItem -Recurse -Force; macOS: ls -aR)

```text
Directory: C:\path\to\van-phong-pham-website

Mode                 LastWriteTime         Length Name
----                 -------------         ------ ----
d-----        23/10/2026   08:44 AM                backup
d-----        23/10/2026   08:44 AM                data
d-----        23/10/2026   08:44 AM                docs
d-----        23/10/2026   08:44 AM                logs
d-----        23/10/2026   08:44 AM                media
-a----        23/10/2026   08:42 AM           1250 bai_lam.md

```

## 4.3. Lọc log thủ công

### Lệnh đã chạy

```shell
Get-Content logs/shop.log |Select-String "ERROR" | Set-Location logs/daily_report.txt
Get-Content logs/daily_report.txt

```

### 2 dòng lỗi đầu tiên trong daily_report.txt

```text
=== BAO CAO LOI ===
2025-10-28 08:24:28 [ERROR] Thanh toan bi tu choi
2025-10-28 08:30:28 [ERROR] Thanh toan bi tu choi
```

## 4.4. Thao tác với tài liệu Markdown (README.md)

> README.md được chấm trực tiếp trên file `README.md` ở gốc dự án. Đánh dấu `[x]` vào các mục đã làm để tự kiểm tra.

- [X] Heading 1: tên dự án
- [X] Mô tả ngắn có cả **in đậm** và *in nghiêng*
- [X] Bảng `Thư mục | Nội dung`, ít nhất 3 dòng dữ liệu
- [X] Code block chứa lệnh PowerShell/Terminal liệt kê nội dung dự án
- [X] Task list ít nhất 2 công việc, có ít nhất 1 mục đã hoàn thành
- [X] Blockquote nhắc không upload `.env`

## 4.5. Chạy script auto_report (.ps1 hoặc .sh)

### Lệnh chạy (Windows: nếu dùng Phụ lục 2 thì ghi cả 2 lệnh)

```shell

# Script: auto_report.ps1
# Tac gia: Nguyen Duc Hieu 
# Ngay tao: 23/10/2026
Write-Host "Dang loc log..."
"=== BAO CAO LOI ===" > logs/auto_report.txt
Get-Content logs/shoplog.txt | Select-String "ERROR" >> logs/auto_report.txt
Write-Host "Hoan thanh!
```

### Nội dung file logs/auto_report.txt

```text
=== BAO CAO LOI ===
2025-10-28 08:24:28 [ERROR] Thanh toan bi tu choi
2025-10-28 08:30:28 [ERROR] Thanh toan bi tu choi
2025-10-28 08:33:28 [ERROR] Ket noi kho du lieu that bai
```

## 4.6. Đếm số dòng lỗi

### Lệnh

```shell
(Get-Content logs/shop.log | Select-String "ERROR") .Count

```

### Kết quả đếm (Windows: giá trị Count; macOS: con số in ra)

```text
6
```

## 4.7. Báo cáo tổng hợp

### Các lệnh đã chạy

```shell
"=== TONG HOP LOG ===" > logs/tong_hop.txt
Get-Content logs/shop.log | Select-String "WARNING" >> logs/tong_hop.txt
(Get-Content logs/shop.log | Select-String "ERROR").Count >> logs/tong_hop.txt
Get-Content logs/tong_hop.txt
```

### Nội dung file logs/tong_hop.txt

```text
=== TONG HOP LOG ===
2025-10-28 08:24:28 [ERROR] Thanh toan bi tu choi
2025-10-28 08:30:28 [ERROR] Thanh toan bi tu choi
2025-10-28 08:33:28 [ERROR] Ket noi kho du lieu that bai
2025-10-28 08:39:28 [ERROR] Ket noi kho du lieu that bai
2025-10-28 08:42:28 [ERROR] Ket noi kho du lieu that bai
2025-10-28 09:12:28 [ERROR] Thanh toan bi tu choi
```

## 4.8. Excel
### Lệnh đã chạy 
- ten_san_pham: =VLOOKUP(C2, SanPham!$A$2:$C$100, 2, FALSE)
- gia_ban: =VLOOKUP(C2, SanPham!$A$2:$C$100, 3, FALSE)
- tien_ban: =D2*F2
