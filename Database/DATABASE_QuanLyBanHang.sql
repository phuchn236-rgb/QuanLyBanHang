-- =====================================================================
-- CSDL HangHoa - Quản lý bán hàng (bản đầy đủ: có phân quyền + dữ liệu mẫu)
-- Chạy toàn bộ file trong SQL Server Management Studio (F5).
-- LƯU Ý: file sẽ XÓA và tạo lại database HangHoa nếu đã tồn tại.
-- =====================================================================
USE master;
GO
IF DB_ID('HangHoa') IS NOT NULL
BEGIN
    ALTER DATABASE HangHoa SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE HangHoa;
END
GO
CREATE DATABASE HangHoa;
GO
USE HangHoa;
GO

-- ---------- Tạo bảng (bảng cha trước, bảng con sau) ----------
CREATE TABLE tblLoaiHang (
    MaLoaiHang  nvarchar(50) PRIMARY KEY,
    TenLoaiHang nvarchar(50)
);

CREATE TABLE tblHang (
    MaHang     nvarchar(50) PRIMARY KEY,
    TenHang    nvarchar(50),
    MaLoaiHang nvarchar(50) FOREIGN KEY REFERENCES tblLoaiHang(MaLoaiHang),
    SoLuong    int,
    DonGiaNhap float,
    DonGiaBan  float,
    Anh        nvarchar(255),
    GhiChu     nvarchar(255)
);

CREATE TABLE tblNhanVien (
    MaNhanVien  nvarchar(50) PRIMARY KEY,
    TenNhanVien nvarchar(50),
    GioiTinh    nvarchar(10),
    DiaChi      nvarchar(100),
    DienThoai   nvarchar(20),
    NgaySinh    date
);

CREATE TABLE tblKhach (
    MaKhach   nvarchar(50) PRIMARY KEY,
    TenKhach  nvarchar(50),
    DiaChi    nvarchar(100),
    DienThoai nvarchar(20)
);

CREATE TABLE tblHDBan (
    MaHDBan    nvarchar(50) PRIMARY KEY,
    MaNhanVien nvarchar(50) FOREIGN KEY REFERENCES tblNhanVien(MaNhanVien),
    NgayBan    date,
    MaKhach    nvarchar(50) FOREIGN KEY REFERENCES tblKhach(MaKhach),
    TongTien   float
);

CREATE TABLE tblChiTietHDBan (
    MaHDBan   nvarchar(50) FOREIGN KEY REFERENCES tblHDBan(MaHDBan),
    MaHang    nvarchar(50) FOREIGN KEY REFERENCES tblHang(MaHang),
    SoLuong   int,
    GiamGia   float,          -- tính theo phần trăm (0..100)
    ThanhTien float,
    PRIMARY KEY (MaHDBan, MaHang)
);

-- Bảng tài khoản đăng nhập + phân quyền
CREATE TABLE tblUser (
    userName   nvarchar(50) PRIMARY KEY,
    passWord   nvarchar(50),
    Quyen      nvarchar(20) NOT NULL DEFAULT N'NhanVien',   -- 'Admin' hoặc 'NhanVien'
    MaNhanVien nvarchar(50) NULL FOREIGN KEY REFERENCES tblNhanVien(MaNhanVien)
);
GO

-- ---------- Dữ liệu mẫu ----------
INSERT INTO tblLoaiHang VALUES
 (N'LH01', N'Đồ uống'), (N'LH02', N'Đồ ăn'), (N'LH03', N'Đồ gia dụng'),
 (N'LH04', N'Đồ hàng ngày'), (N'LH05', N'Thuốc');

INSERT INTO tblHang (MaHang, TenHang, MaLoaiHang, SoLuong, DonGiaNhap, DonGiaBan, Anh, GhiChu) VALUES
 (N'H01', N'Gấu bông',          N'LH04',  50,  100000,  150000, NULL, NULL),
 (N'H02', N'Giỏ mây',           N'LH03',  40,   50000,   80000, NULL, NULL),
 (N'H03', N'Nước cam ép',       N'LH01', 100,    8000,   15000, NULL, NULL),
 (N'H04', N'Cà phê hộp',        N'LH01',  80,   12000,   20000, NULL, NULL),
 (N'H05', N'Bánh quy bơ',       N'LH02', 120,   15000,   25000, NULL, NULL),
 (N'H06', N'Mì ly',             N'LH02', 200,    5000,    8000, NULL, NULL),
 (N'H07', N'Nồi cơm điện',      N'LH03',  15,  450000,  600000, NULL, NULL),
 (N'H08', N'Khăn mặt',          N'LH04',  90,   12000,   20000, NULL, NULL),
 (N'H09', N'Paracetamol 500mg', N'LH05', 150,    1500,    3000, NULL, N'Công dụng: giảm đau, hạ sốt. Thành phần: Paracetamol 500mg. HSD: 12/2027'),
 (N'H10', N'Vitamin C 500mg',   N'LH05', 100,    2000,    4000, NULL, N'Công dụng: bổ sung vitamin C. Thành phần: Acid ascorbic 500mg. HSD: 06/2027');

INSERT INTO tblNhanVien VALUES
 (N'NV01', N'Nguyễn Văn A', N'Nam', N'Hà Nội',    N'0901234567', '2000-01-15'),
 (N'NV02', N'Lê Thị C',     N'Nữ',  N'Hải Dương', N'0902345678', '1999-05-20'),
 (N'NV03', N'Phạm Văn D',   N'Nam', N'Nam Định',  N'0903456789', '1998-11-02');

INSERT INTO tblKhach VALUES
 (N'KH01', N'Trần Thị B',    N'Hải Phòng', N'0912345678'),
 (N'KH02', N'Nguyễn Văn E',  N'Hà Nội',    N'0913456789'),
 (N'KH03', N'Hoàng Thị F',   N'Bắc Ninh',  N'0914567890'),
 (N'KH04', N'Vũ Văn G',      N'Hưng Yên',  N'0915678901'),
 (N'KH05', N'Đặng Thị H',    N'Thái Bình', N'0916789012');

-- Tài khoản: admin/123 (Admin), nv01/123, nv02/123, nv03/123 (Nhân viên)
INSERT INTO tblUser (userName, passWord, Quyen, MaNhanVien) VALUES
 (N'admin', N'123', N'Admin',    NULL),
 (N'nv01',  N'123', N'NhanVien', N'NV01'),
 (N'nv02',  N'123', N'NhanVien', N'NV02'),
 (N'nv03',  N'123', N'NhanVien', N'NV03');

-- 10 hóa đơn mẫu (rải trong 2025-2026 để báo cáo có số liệu)
INSERT INTO tblHDBan (MaHDBan, MaNhanVien, NgayBan, MaKhach, TongTien) VALUES
 (N'HD001', N'NV01', '2025-11-10', N'KH01', 0),
 (N'HD002', N'NV02', '2025-12-20', N'KH02', 0),
 (N'HD003', N'NV01', '2026-01-15', N'KH03', 0),
 (N'HD004', N'NV03', '2026-02-18', N'KH04', 0),
 (N'HD005', N'NV02', '2026-04-05', N'KH05', 0),
 (N'HD006', N'NV01', '2026-05-22', N'KH01', 0),
 (N'HD007', N'NV03', '2026-07-09', N'KH02', 0),
 (N'HD008', N'NV02', '2026-08-14', N'KH03', 0),
 (N'HD009', N'NV01', '2026-09-30', N'KH04', 0),
 (N'HD010', N'NV03', '2026-10-12', N'KH05', 0);

INSERT INTO tblChiTietHDBan (MaHDBan, MaHang, SoLuong, GiamGia, ThanhTien) VALUES
 (N'HD001', N'H01',  2,  0, 0), (N'HD001', N'H03', 10,  0, 0),
 (N'HD002', N'H07',  1,  5, 0), (N'HD002', N'H05',  5,  0, 0),
 (N'HD003', N'H02',  3,  0, 0), (N'HD003', N'H06', 20, 10, 0),
 (N'HD004', N'H04', 10,  0, 0), (N'HD004', N'H08',  4,  0, 0),
 (N'HD005', N'H09', 30,  0, 0), (N'HD005', N'H10', 10,  0, 0),
 (N'HD006', N'H01',  1, 10, 0), (N'HD006', N'H07',  1, 30, 0),   -- H07 giảm 30%: bán LỖ (để demo)
 (N'HD007', N'H05', 12,  5, 0), (N'HD007', N'H03',  8,  0, 0),
 (N'HD008', N'H06', 30,  0, 0), (N'HD008', N'H04',  6,  0, 0),
 (N'HD009', N'H02',  2,  0, 0), (N'HD009', N'H08',  6,  5, 0),
 (N'HD010', N'H10', 15,  0, 0), (N'HD010', N'H09', 20,  0, 0);

-- Tính ThanhTien, TongTien và trừ tồn kho cho khớp với hóa đơn mẫu
UPDATE ct SET ThanhTien = ct.SoLuong * h.DonGiaBan * (1 - ct.GiamGia / 100.0)
FROM tblChiTietHDBan ct JOIN tblHang h ON h.MaHang = ct.MaHang;

UPDATE b SET TongTien = (SELECT SUM(ThanhTien) FROM tblChiTietHDBan WHERE MaHDBan = b.MaHDBan)
FROM tblHDBan b;

UPDATE h SET SoLuong = h.SoLuong - x.sl
FROM tblHang h JOIN (SELECT MaHang, SUM(SoLuong) AS sl FROM tblChiTietHDBan GROUP BY MaHang) x ON x.MaHang = h.MaHang;
GO

-- Kiểm tra
SELECT * FROM tblHang;
SELECT * FROM tblHDBan;
GO
