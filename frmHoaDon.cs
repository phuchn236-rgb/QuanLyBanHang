using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace QuanLyBanHang
{
    public class frmHoaDon : Form
    {
        TextBox txtMaHD, txtTim;
        DateTimePicker dtpNgay;
        ComboBox cboNV, cboKhach, cboHang;
        NumericUpDown nudSL, nudGG;
        Label lblInfo, lblTong;
        DataGridView dgvCT, dgvDS;
        Button btnXoa;

        DataTable dtCT;                                              // các dòng chi tiết đang lập (trong bộ nhớ)
        Dictionary<string, int> oldQty = new Dictionary<string, int>(); // số lượng cũ của HĐ đang sửa (để tính tồn kho)

        public frmHoaDon()
        {
            Text = "Lập hóa đơn bán hàng";
            Font = new Font("Segoe UI", 10F);
            ClientSize = new Size(1150, 740);
            StartPosition = FormStartPosition.CenterParent;

            // ----- Thông tin chung -----
            UI.Lbl(this, "Mã HĐ", 15, 12, 80);
            txtMaHD = UI.Txt(this, 100, 12, 110);
            txtMaHD.ReadOnly = true;
            UI.Lbl(this, "Ngày bán", 235, 12, 80);
            dtpNgay = new DateTimePicker { Left = 320, Top = 12, Width = 130, Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy" };
            Controls.Add(dtpNgay);
            UI.Lbl(this, "Nhân viên", 475, 12, 80);
            cboNV = UI.Cbo(this, 560, 12, 200);
            UI.Lbl(this, "Khách hàng", 785, 12, 95);
            cboKhach = UI.Cbo(this, 885, 12, 235);

            // ----- Thêm hàng vào hóa đơn -----
            UI.Lbl(this, "Hàng", 15, 55, 80);
            cboHang = UI.Cbo(this, 100, 55, 300);
            cboHang.SelectedIndexChanged += (s, e) => HienThongTinHang();
            UI.Lbl(this, "SL", 420, 55, 30);
            nudSL = new NumericUpDown { Left = 455, Top = 55, Width = 70, Minimum = 1, Maximum = 100000, Value = 1 };
            Controls.Add(nudSL);
            UI.Lbl(this, "Giảm giá %", 545, 55, 90);
            nudGG = new NumericUpDown { Left = 640, Top = 55, Width = 70, Minimum = 0, Maximum = 100 };
            Controls.Add(nudGG);
            lblInfo = new Label { Left = 730, Top = 59, Width = 390, ForeColor = Color.DarkBlue };
            Controls.Add(lblInfo);

            UI.Btn(this, "Thêm dòng", 100, 92, 120, (s, e) => ThemDong());
            UI.Btn(this, "Xóa dòng", 230, 92, 120, (s, e) => XoaDong());

            // ----- Lưới chi tiết -----
            dtCT = new DataTable();
            dtCT.Columns.Add("MaHang", typeof(string));
            dtCT.Columns.Add("TenHang", typeof(string));
            dtCT.Columns.Add("SoLuong", typeof(int));
            dtCT.Columns.Add("DonGia", typeof(double));
            dtCT.Columns.Add("GiamGia", typeof(double));
            dtCT.Columns.Add("ThanhTien", typeof(double));
            dgvCT = UI.Grid(this, 15, 132, 1120, 200);
            dgvCT.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            dgvCT.DataSource = dtCT;
            UI.Col(dgvCT, "MaHang", "Mã hàng");
            UI.Col(dgvCT, "TenHang", "Tên hàng");
            UI.Col(dgvCT, "SoLuong", "Số lượng", "N0");
            UI.Col(dgvCT, "DonGia", "Đơn giá", "N0");
            UI.Col(dgvCT, "GiamGia", "Giảm giá (%)", "N0");
            UI.Col(dgvCT, "ThanhTien", "Thành tiền", "N0");

            // ----- Tổng tiền + các nút -----
            var lt = new Label { Text = "Tổng tiền:", Left = 15, Top = 345, Width = 100, Font = new Font("Segoe UI", 12F, FontStyle.Bold) };
            Controls.Add(lt);
            lblTong = new Label { Left = 120, Top = 345, Width = 280, ForeColor = Color.Firebrick, Font = new Font("Segoe UI", 12F, FontStyle.Bold) };
            Controls.Add(lblTong);

            UI.Btn(this, "Hóa đơn mới", 420, 340, 120, (s, e) => HoaDonMoi());
            UI.Btn(this, "Lưu hóa đơn", 550, 340, 120, (s, e) => Luu());
            btnXoa = UI.Btn(this, "Xóa hóa đơn", 680, 340, 120, (s, e) => XoaHoaDon());
            UI.Btn(this, "In ra Excel", 810, 340, 120, (s, e) => InExcel());

            // ----- Danh sách hóa đơn đã lập -----
            var lds = new Label { Text = "Danh sách hóa đơn đã lập (bấm vào dòng để xem / sửa)", Left = 15, Top = 395, Width = 500, Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
            Controls.Add(lds);
            UI.Lbl(this, "Tìm", 600, 392, 40);
            txtTim = UI.Txt(this, 645, 392, 250);
            UI.Btn(this, "Tìm", 905, 390, 80, (s, e) => LoadDS());
            txtTim.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadDS(); };

            dgvDS = UI.Grid(this, 15, 430, 1120, 295);
            dgvDS.CellClick += (s, e) =>
            {
                if (e.RowIndex >= 0) MoHoaDon(Convert.ToString(dgvDS.Rows[e.RowIndex].Cells["MaHDBan"].Value));
            };

            Load += (s, e) =>
            {
                UI.BindCombo(cboNV, DBHelper.Query("SELECT MaNhanVien, TenNhanVien FROM tblNhanVien ORDER BY MaNhanVien"), "TenNhanVien", "MaNhanVien");
                UI.BindCombo(cboKhach, DBHelper.Query("SELECT MaKhach, TenKhach FROM tblKhach ORDER BY MaKhach"), "TenKhach", "MaKhach");
                LoadHang();
                LoadDS();
                HoaDonMoi();
            };
        }

        // ================== Dữ liệu danh mục ==================
        void LoadHang()
        {
            object keep = cboHang.SelectedValue;
            UI.BindCombo(cboHang,
                DBHelper.Query("SELECT MaHang, TenHang, MaHang + N' - ' + TenHang AS Hien, SoLuong, DonGiaBan FROM tblHang ORDER BY MaHang"),
                "Hien", "MaHang");
            if (keep != null) cboHang.SelectedValue = keep;
            HienThongTinHang();
        }

        void HienThongTinHang()
        {
            var drv = cboHang.SelectedItem as DataRowView;
            if (drv == null) { lblInfo.Text = ""; return; }
            lblInfo.Text = "Đơn giá: " + Convert.ToDouble(drv["DonGiaBan"]).ToString("N0") +
                           "  |  Tồn kho: " + Convert.ToInt32(drv["SoLuong"]);
        }

        void LoadDS()
        {
            string k = "%" + txtTim.Text.Trim() + "%";
            dgvDS.DataSource = DBHelper.Query(
                "SELECT b.MaHDBan, b.NgayBan, n.TenNhanVien, k.TenKhach, b.TongTien " +
                "FROM tblHDBan b " +
                "LEFT JOIN tblNhanVien n ON n.MaNhanVien = b.MaNhanVien " +
                "LEFT JOIN tblKhach k ON k.MaKhach = b.MaKhach " +
                "WHERE b.MaHDBan LIKE @k OR k.TenKhach LIKE @k OR n.TenNhanVien LIKE @k " +
                "ORDER BY b.NgayBan DESC, b.MaHDBan DESC",
                DBHelper.P("@k", k));
            UI.Col(dgvDS, "MaHDBan", "Mã HĐ");
            UI.Col(dgvDS, "NgayBan", "Ngày bán", "dd/MM/yyyy");
            UI.Col(dgvDS, "TenNhanVien", "Nhân viên");
            UI.Col(dgvDS, "TenKhach", "Khách hàng");
            UI.Col(dgvDS, "TongTien", "Tổng tiền", "N0");
        }

        // ================== Thao tác trên form ==================
        static double Tinh(double sl, double gia, double gg)
        {
            return sl * gia * (1 - gg / 100.0);
        }

        void RefreshTong()
        {
            double tong = 0;
            foreach (DataRow r in dtCT.Rows) tong += Convert.ToDouble(r["ThanhTien"]);
            lblTong.Text = tong.ToString("N0") + " VNĐ";
        }

        void HoaDonMoi()
        {
            txtMaHD.Text = UI.NextId("tblHDBan", "MaHDBan", "HD", 3);
            dtpNgay.Value = DateTime.Today;
            dtCT.Clear();
            oldQty.Clear();
            nudSL.Value = 1;
            nudGG.Value = 0;
            if (cboKhach.Items.Count > 0) cboKhach.SelectedIndex = 0;
            if (!Session.IsAdmin && Session.MaNhanVien != null)
                cboNV.SelectedValue = Session.MaNhanVien;           // nhân viên: tự điền chính mình
            else if (cboNV.Items.Count > 0)
                cboNV.SelectedIndex = 0;
            cboNV.Enabled = Session.IsAdmin;
            btnXoa.Enabled = false;
            LoadHang();                                             // cập nhật lại tồn kho mới nhất
            RefreshTong();
        }

        void ThemDong()
        {
            var drv = cboHang.SelectedItem as DataRowView;
            if (drv == null) { UI.Warn("Chưa chọn hàng!"); return; }

            string ma = drv["MaHang"].ToString();
            int ton = Convert.ToInt32(drv["SoLuong"]);
            double gia = Convert.ToDouble(drv["DonGiaBan"]);
            int sl = (int)nudSL.Value;
            double gg = (double)nudGG.Value;

            int cu;
            oldQty.TryGetValue(ma, out cu);       // nếu đang sửa HĐ cũ: kho đã trừ phần này rồi

            DataRow exist = null;
            foreach (DataRow r in dtCT.Rows)
                if (r["MaHang"].ToString() == ma) { exist = r; break; }

            int tongSL = sl + (exist == null ? 0 : Convert.ToInt32(exist["SoLuong"]));
            if (tongSL > ton + cu)
            {
                UI.Warn("Không đủ hàng! Số lượng có thể bán: " + (ton + cu));
                return;
            }

            if (exist == null)
            {
                dtCT.Rows.Add(ma, drv["TenHang"].ToString(), sl, gia, gg, Tinh(sl, gia, gg));
            }
            else
            {
                exist["SoLuong"] = tongSL;
                exist["DonGia"] = gia;
                exist["GiamGia"] = gg;
                exist["ThanhTien"] = Tinh(tongSL, gia, gg);
            }
            RefreshTong();
        }

        void XoaDong()
        {
            if (dgvCT.CurrentRow == null) { UI.Warn("Hãy chọn dòng cần xóa!"); return; }
            var drv = dgvCT.CurrentRow.DataBoundItem as DataRowView;
            if (drv == null) return;
            drv.Row.Delete();
            dtCT.AcceptChanges();
            RefreshTong();
        }

        void MoHoaDon(string ma)
        {
            DataTable h = DBHelper.Query("SELECT NgayBan, MaNhanVien, MaKhach FROM tblHDBan WHERE MaHDBan=@m", DBHelper.P("@m", ma));
            if (h.Rows.Count == 0) return;

            txtMaHD.Text = ma;
            dtpNgay.Value = Convert.ToDateTime(h.Rows[0]["NgayBan"]);
            cboNV.SelectedValue = h.Rows[0]["MaNhanVien"] is DBNull ? null : h.Rows[0]["MaNhanVien"];
            cboKhach.SelectedValue = h.Rows[0]["MaKhach"] is DBNull ? null : h.Rows[0]["MaKhach"];

            dtCT.Clear();
            oldQty.Clear();
            DataTable ct = DBHelper.Query(
                "SELECT c.MaHang, h.TenHang, c.SoLuong, h.DonGiaBan AS DonGia, c.GiamGia, c.ThanhTien " +
                "FROM tblChiTietHDBan c JOIN tblHang h ON h.MaHang = c.MaHang WHERE c.MaHDBan = @m",
                DBHelper.P("@m", ma));
            foreach (DataRow r in ct.Rows)
            {
                dtCT.Rows.Add(r["MaHang"], r["TenHang"], r["SoLuong"], r["DonGia"], r["GiamGia"], r["ThanhTien"]);
                oldQty[r["MaHang"].ToString()] = Convert.ToInt32(r["SoLuong"]);
            }
            btnXoa.Enabled = Session.IsAdmin;       // chỉ Admin được xóa hóa đơn
            RefreshTong();
        }

        // ================== Lưu hóa đơn (Transaction) ==================
        void Luu()
        {
            if (dtCT.Rows.Count == 0) { UI.Warn("Hóa đơn chưa có mặt hàng nào!"); return; }
            if (cboNV.SelectedValue == null) { UI.Warn("Chưa chọn nhân viên!"); return; }
            if (cboKhach.SelectedValue == null) { UI.Warn("Chưa chọn khách hàng!"); return; }

            string ma = txtMaHD.Text;
            double tong = 0;
            foreach (DataRow r in dtCT.Rows) tong += Convert.ToDouble(r["ThanhTien"]);

            try
            {
                using (var cn = new SqlConnection(DBHelper.ConnStr))
                {
                    cn.Open();
                    // Nếu có lỗi giữa chừng, không gọi Commit => using tự Rollback: không bị "nửa hóa đơn"
                    using (SqlTransaction tr = cn.BeginTransaction())
                    {
                        int count = Convert.ToInt32(DBHelper.ScalarTx(cn, tr,
                            "SELECT COUNT(*) FROM tblHDBan WHERE MaHDBan=@m", DBHelper.P("@m", ma)));

                        if (count > 0)
                        {
                            // Sửa HĐ cũ: hoàn lại kho phần đã bán, xóa chi tiết cũ, cập nhật đầu hóa đơn
                            DBHelper.ExecTx(cn, tr,
                                "UPDATE h SET h.SoLuong = h.SoLuong + c.SoLuong FROM tblHang h " +
                                "JOIN tblChiTietHDBan c ON c.MaHang = h.MaHang WHERE c.MaHDBan = @m",
                                DBHelper.P("@m", ma));
                            DBHelper.ExecTx(cn, tr, "DELETE FROM tblChiTietHDBan WHERE MaHDBan=@m", DBHelper.P("@m", ma));
                            DBHelper.ExecTx(cn, tr,
                                "UPDATE tblHDBan SET MaNhanVien=@nv, NgayBan=@ngay, MaKhach=@kh WHERE MaHDBan=@m",
                                DBHelper.P("@nv", cboNV.SelectedValue), DBHelper.P("@ngay", dtpNgay.Value.Date),
                                DBHelper.P("@kh", cboKhach.SelectedValue), DBHelper.P("@m", ma));
                        }
                        else
                        {
                            DBHelper.ExecTx(cn, tr,
                                "INSERT INTO tblHDBan(MaHDBan, MaNhanVien, NgayBan, MaKhach, TongTien) VALUES(@m, @nv, @ngay, @kh, 0)",
                                DBHelper.P("@m", ma), DBHelper.P("@nv", cboNV.SelectedValue),
                                DBHelper.P("@ngay", dtpNgay.Value.Date), DBHelper.P("@kh", cboKhach.SelectedValue));
                        }

                        foreach (DataRow r in dtCT.Rows)
                        {
                            string mh = r["MaHang"].ToString();
                            int sl = Convert.ToInt32(r["SoLuong"]);

                            // Trừ kho chỉ khi còn đủ hàng (điều kiện SoLuong >= @sl nằm ngay trong câu UPDATE)
                            int n = DBHelper.ExecTx(cn, tr,
                                "UPDATE tblHang SET SoLuong = SoLuong - @sl WHERE MaHang = @h AND SoLuong >= @sl",
                                DBHelper.P("@sl", sl), DBHelper.P("@h", mh));
                            if (n == 0)
                                throw new InvalidOperationException("Hàng '" + r["TenHang"] + "' không đủ số lượng tồn kho!");

                            DBHelper.ExecTx(cn, tr,
                                "INSERT INTO tblChiTietHDBan(MaHDBan, MaHang, SoLuong, GiamGia, ThanhTien) VALUES(@m, @h, @sl, @gg, @tt)",
                                DBHelper.P("@m", ma), DBHelper.P("@h", mh), DBHelper.P("@sl", sl),
                                DBHelper.P("@gg", Convert.ToDouble(r["GiamGia"])), DBHelper.P("@tt", Convert.ToDouble(r["ThanhTien"])));
                        }

                        DBHelper.ExecTx(cn, tr, "UPDATE tblHDBan SET TongTien=@t WHERE MaHDBan=@m",
                            DBHelper.P("@t", tong), DBHelper.P("@m", ma));

                        tr.Commit();
                    }
                }
                UI.Info("Đã lưu hóa đơn " + ma + ".");
                LoadDS();
                HoaDonMoi();
            }
            catch (InvalidOperationException ex) { UI.Warn(ex.Message); }
            catch (SqlException ex) { UI.ShowSql(ex, "lưu"); }
        }

        void XoaHoaDon()
        {
            if (!Session.IsAdmin) { UI.Warn("Chỉ Admin mới được xóa hóa đơn!"); return; }
            string ma = txtMaHD.Text;
            if (Convert.ToInt32(DBHelper.Scalar("SELECT COUNT(*) FROM tblHDBan WHERE MaHDBan=@m", DBHelper.P("@m", ma))) == 0)
            {
                UI.Warn("Hóa đơn này chưa được lưu!");
                return;
            }
            if (!UI.Confirm("Xóa hóa đơn " + ma + "? Số lượng hàng sẽ được hoàn lại kho.")) return;

            try
            {
                using (var cn = new SqlConnection(DBHelper.ConnStr))
                {
                    cn.Open();
                    using (SqlTransaction tr = cn.BeginTransaction())
                    {
                        DBHelper.ExecTx(cn, tr,
                            "UPDATE h SET h.SoLuong = h.SoLuong + c.SoLuong FROM tblHang h " +
                            "JOIN tblChiTietHDBan c ON c.MaHang = h.MaHang WHERE c.MaHDBan = @m", DBHelper.P("@m", ma));
                        DBHelper.ExecTx(cn, tr, "DELETE FROM tblChiTietHDBan WHERE MaHDBan=@m", DBHelper.P("@m", ma));
                        DBHelper.ExecTx(cn, tr, "DELETE FROM tblHDBan WHERE MaHDBan=@m", DBHelper.P("@m", ma));
                        tr.Commit();
                    }
                }
                UI.Info("Đã xóa hóa đơn.");
                LoadDS();
                HoaDonMoi();
            }
            catch (SqlException ex) { UI.ShowSql(ex, "xóa"); }
        }

        // ================== In hóa đơn ra Excel ==================
        void InExcel()
        {
            if (dtCT.Rows.Count == 0) { UI.Warn("Hóa đơn chưa có mặt hàng để in!"); return; }

            using (var dlg = new SaveFileDialog())
            {
                dlg.Filter = "Excel (*.xlsx)|*.xlsx";
                dlg.FileName = "HoaDon_" + txtMaHD.Text + ".xlsx";
                if (dlg.ShowDialog() != DialogResult.OK) return;

                try
                {
                    using (var wb = new XLWorkbook())
                    {
                        IXLWorksheet ws = wb.Worksheets.Add("HoaDon");

                        ws.Cell(1, 1).Value = "HÓA ĐƠN BÁN HÀNG";
                        ws.Range(1, 1, 1, 6).Merge();
                        ws.Cell(1, 1).Style.Font.Bold = true;
                        ws.Cell(1, 1).Style.Font.FontSize = 16;
                        ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        ws.Cell(3, 1).Value = "Mã hóa đơn:";   ws.Cell(3, 2).Value = txtMaHD.Text;
                        ws.Cell(4, 1).Value = "Ngày bán:";     ws.Cell(4, 2).Value = dtpNgay.Value.ToString("dd/MM/yyyy");
                        ws.Cell(5, 1).Value = "Nhân viên:";    ws.Cell(5, 2).Value = cboNV.Text;
                        ws.Cell(6, 1).Value = "Khách hàng:";   ws.Cell(6, 2).Value = cboKhach.Text;

                        string[] hdr = { "STT", "Mã hàng", "Tên hàng", "Số lượng", "Đơn giá", "Giảm giá (%)", "Thành tiền" };
                        int row = 8;
                        for (int c = 0; c < hdr.Length; c++)
                        {
                            ws.Cell(row, c + 1).Value = hdr[c];
                            ws.Cell(row, c + 1).Style.Font.Bold = true;
                            ws.Cell(row, c + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                        }

                        int stt = 1;
                        double tong = 0;
                        foreach (DataRow r in dtCT.Rows)
                        {
                            row++;
                            ws.Cell(row, 1).Value = stt++;
                            ws.Cell(row, 2).Value = r["MaHang"].ToString();
                            ws.Cell(row, 3).Value = r["TenHang"].ToString();
                            ws.Cell(row, 4).Value = Convert.ToInt32(r["SoLuong"]);
                            ws.Cell(row, 5).Value = Convert.ToDouble(r["DonGia"]);
                            ws.Cell(row, 6).Value = Convert.ToDouble(r["GiamGia"]);
                            ws.Cell(row, 7).Value = Convert.ToDouble(r["ThanhTien"]);
                            tong += Convert.ToDouble(r["ThanhTien"]);
                        }
                        row++;
                        ws.Cell(row, 6).Value = "TỔNG TIỀN:";
                        ws.Cell(row, 6).Style.Font.Bold = true;
                        ws.Cell(row, 7).Value = tong;
                        ws.Cell(row, 7).Style.Font.Bold = true;

                        ws.Range(9, 5, row, 5).Style.NumberFormat.Format = "#,##0";
                        ws.Range(9, 7, row, 7).Style.NumberFormat.Format = "#,##0";
                        ws.Range(8, 1, row - 1, 7).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        ws.Range(8, 1, row - 1, 7).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        ws.Columns().AdjustToContents();

                        wb.SaveAs(dlg.FileName);
                    }
                    UI.Info("Đã xuất hóa đơn ra Excel:\n" + dlg.FileName);
                }
                catch (Exception ex)
                {
                    UI.Err("Không thể xuất file (file đang mở trong Excel?).\n" + ex.Message);
                }
            }
        }
    }
}
