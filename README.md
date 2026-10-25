# Quản lý bán hàng - WinForms (.NET Framework 4.8) + SQL Server

## Cách chạy
1. Mở SSMS, chạy file `Database/DATABASE_QuanLyBanHang.sql` (tạo DB `HangHoa` + dữ liệu mẫu).
2. Mở `QuanLyBanHang.csproj` bằng Visual Studio 2019/2022 (cài workload ".NET desktop development").
3. Sửa `Server=` trong `App.config` cho đúng máy (mặc định `.\SQLEXPRESS`).
4. Build (VS tự tải ClosedXML qua NuGet) rồi nhấn F5.

## Tài khoản mẫu
| Vai trò | Tài khoản | Mật khẩu |
|---|---|---|
| Admin | admin | 123 |
| Nhân viên | nv01 / nv02 / nv03 | 123 |

## Phân quyền
- Admin: toàn bộ chức năng (Hàng, Khách, Nhân viên, Hóa đơn, Thống kê & Báo cáo).
- Nhân viên: Hàng (chỉ xem/tìm), Khách hàng, Lập hóa đơn (không được xóa hóa đơn). Không thấy Nhân viên và Thống kê.

## Cấu trúc
| File | Vai trò |
|---|---|
| Program.cs | Vòng lặp Đăng nhập -> Giao diện chính -> Đăng xuất |
| DBHelper.cs | Kết nối SQL, Query/Exec/Scalar, hàm cho Transaction, lớp Session |
| UI.cs | Hàm dựng control, thông báo, sinh mã tự động, dịch lỗi SQL |
| frmDangNhap / frmMain | Đăng nhập, giao diện chính (MDI + menu theo quyền) |
| frmHang / frmKhach / frmNhanVien | CRUD + tìm kiếm |
| frmHoaDon | Lập hóa đơn, trừ kho (Transaction), in Excel |
| frmThongKe | Doanh thu, lãi/lỗ theo tháng/quý/năm/mặt hàng, biểu đồ, xuất Excel |
