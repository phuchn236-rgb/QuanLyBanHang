using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace QuanLyBanHang
{
    public class frmNhanVien : Form
    {
        TextBox txtMa, txtTen, txtDiaChi, txtDienThoai, txtTim;
        ComboBox cboGT;
        DateTimePicker dtpNgaySinh;
        CheckBox chkTaoTK;
        DataGridView dgv;

        public frmNhanVien()
        {
            Text = "Quản lý nhân viên";
            Font = new Font("Segoe UI", 10F);
            ClientSize = new Size(960, 620);
            StartPosition = FormStartPosition.CenterParent;

            UI.Lbl(this, "Mã nhân viên", 15, 15, 95);
            txtMa = UI.Txt(this, 115, 15, 150);
            UI.Lbl(this, "Tên nhân viên", 300, 15, 100);
            txtTen = UI.Txt(this, 405, 15, 250);

            UI.Lbl(this, "Giới tính", 15, 55, 95);
            cboGT = UI.Cbo(this, 115, 55, 150);
            cboGT.Items.AddRange(new object[] { "Nam", "Nữ" });
            UI.Lbl(this, "Ngày sinh", 300, 55, 100);
            dtpNgaySinh = new DateTimePicker { Left = 405, Top = 55, Width = 150, Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy" };
            Controls.Add(dtpNgaySinh);

            UI.Lbl(this, "Địa chỉ", 15, 95, 95);
            txtDiaChi = UI.Txt(this, 115, 95, 150);
            UI.Lbl(this, "Điện thoại", 300, 95, 100);
            txtDienThoai = UI.Txt(this, 405, 95, 250);

            chkTaoTK = new CheckBox
            {
                Text = "Tạo tài khoản đăng nhập khi thêm (tên = mã NV, mật khẩu mặc định 123)",
                Left = 115, Top = 132, Width = 600, Checked = true
            };
            Controls.Add(chkTaoTK);

            UI.Btn(this, "Thêm", 15, 170, 95, (s, e) => Them());
            UI.Btn(this, "Sửa", 120, 170, 95, (s, e) => Sua());
            UI.Btn(this, "Xóa", 225, 170, 95, (s, e) => Xoa());
            UI.Btn(this, "Làm mới", 330, 170, 95, (s, e) => ClearInput());

            UI.Lbl(this, "Tìm kiếm", 15, 220, 90);
            txtTim = UI.Txt(this, 115, 220, 300);
            UI.Btn(this, "Tìm", 430, 218, 80, (s, e) => LoadGrid());
            txtTim.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadGrid(); };

            dgv = UI.Grid(this, 15, 265, 930, 340);
            dgv.CellClick += Dgv_CellClick;

            Load += (s, e) =>
            {
                if (!Session.IsAdmin)
                {
                    UI.Warn("Chỉ Admin mới được quản lý nhân viên!");
                    Close();
                    return;
                }
                LoadGrid();
                ClearInput();
            };
        }

        void LoadGrid()
        {
            string k = "%" + txtTim.Text.Trim() + "%";
            dgv.DataSource = DBHelper.Query(
                "SELECT MaNhanVien, TenNhanVien, GioiTinh, NgaySinh, DiaChi, DienThoai FROM tblNhanVien " +
                "WHERE MaNhanVien LIKE @k OR TenNhanVien LIKE @k OR DiaChi LIKE @k OR DienThoai LIKE @k ORDER BY MaNhanVien",
                DBHelper.P("@k", k));
            UI.Col(dgv, "MaNhanVien", "Mã NV");
            UI.Col(dgv, "TenNhanVien", "Tên nhân viên");
            UI.Col(dgv, "GioiTinh", "Giới tính");
            UI.Col(dgv, "NgaySinh", "Ngày sinh", "dd/MM/yyyy");
            UI.Col(dgv, "DiaChi", "Địa chỉ");
            UI.Col(dgv, "DienThoai", "Điện thoại");
        }

        void Dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            DataGridViewRow r = dgv.Rows[e.RowIndex];
            txtMa.Text = Convert.ToString(r.Cells["MaNhanVien"].Value);
            txtMa.Enabled = false;
            txtTen.Text = Convert.ToString(r.Cells["TenNhanVien"].Value);
            cboGT.SelectedItem = Convert.ToString(r.Cells["GioiTinh"].Value);
            txtDiaChi.Text = Convert.ToString(r.Cells["DiaChi"].Value);
            txtDienThoai.Text = Convert.ToString(r.Cells["DienThoai"].Value);
            DateTime d;
            if (DateTime.TryParse(Convert.ToString(r.Cells["NgaySinh"].Value), out d)) dtpNgaySinh.Value = d;
            chkTaoTK.Enabled = false;
        }

        void ClearInput()
        {
            txtMa.Enabled = true;
            chkTaoTK.Enabled = true;
            txtMa.Text = UI.NextId("tblNhanVien", "MaNhanVien", "NV", 2);
            txtTen.Clear();
            txtDiaChi.Clear();
            txtDienThoai.Clear();
            cboGT.SelectedIndex = 0;
            dtpNgaySinh.Value = new DateTime(2000, 1, 1);
            txtTen.Focus();
        }

        bool Check()
        {
            if (txtMa.Text.Trim() == "") { UI.Warn("Chưa nhập mã nhân viên!"); return false; }
            if (txtTen.Text.Trim() == "") { UI.Warn("Chưa nhập tên nhân viên!"); txtTen.Focus(); return false; }
            if (cboGT.SelectedIndex < 0) { UI.Warn("Chưa chọn giới tính!"); return false; }
            if (dtpNgaySinh.Value.Date > DateTime.Today.AddYears(-18))
            {
                UI.Warn("Nhân viên phải đủ 18 tuổi!");
                return false;
            }
            string sdt = txtDienThoai.Text.Trim();
            if (sdt != "" && !Regex.IsMatch(sdt, @"^0\d{9,10}$"))
            {
                UI.Warn("Số điện thoại phải bắt đầu bằng 0 và có 10-11 chữ số!");
                txtDienThoai.Focus();
                return false;
            }
            return true;
        }

        SqlParameter[] Params()
        {
            return new[]
            {
                DBHelper.P("@ma", txtMa.Text.Trim()),
                DBHelper.P("@ten", txtTen.Text.Trim()),
                DBHelper.P("@gt", cboGT.SelectedItem.ToString()),
                DBHelper.P("@dc", txtDiaChi.Text.Trim()),
                DBHelper.P("@sdt", txtDienThoai.Text.Trim()),
                DBHelper.P("@ns", dtpNgaySinh.Value.Date),
                DBHelper.P("@tk", chkTaoTK.Checked),
                DBHelper.P("@un", txtMa.Text.Trim().ToLower())
            };
        }

        void Them()
        {
            if (!Check()) return;
            try
            {
                // 1 batch + transaction: thêm nhân viên và (tùy chọn) tài khoản, hoặc cả hai cùng thất bại
                DBHelper.Exec(
                    "SET XACT_ABORT ON; BEGIN TRAN; " +
                    "INSERT INTO tblNhanVien(MaNhanVien, TenNhanVien, GioiTinh, DiaChi, DienThoai, NgaySinh) " +
                    "VALUES(@ma, @ten, @gt, @dc, @sdt, @ns); " +
                    "IF @tk = 1 AND NOT EXISTS (SELECT 1 FROM tblUser WHERE userName = @un) " +
                    "INSERT INTO tblUser(userName, passWord, Quyen, MaNhanVien) VALUES(@un, N'123', N'NhanVien', @ma); " +
                    "COMMIT;", Params());
                UI.Info("Thêm nhân viên thành công.");
                LoadGrid();
                ClearInput();
            }
            catch (SqlException ex) { UI.ShowSql(ex, "thêm"); }
        }

        void Sua()
        {
            if (txtMa.Enabled) { UI.Warn("Hãy chọn một nhân viên trong danh sách để sửa!"); return; }
            if (!Check()) return;
            try
            {
                DBHelper.Exec(
                    "UPDATE tblNhanVien SET TenNhanVien=@ten, GioiTinh=@gt, DiaChi=@dc, DienThoai=@sdt, NgaySinh=@ns " +
                    "WHERE MaNhanVien=@ma", Params());
                UI.Info("Cập nhật thành công.");
                LoadGrid();
            }
            catch (SqlException ex) { UI.ShowSql(ex, "sửa"); }
        }

        void Xoa()
        {
            if (txtMa.Enabled) { UI.Warn("Hãy chọn một nhân viên trong danh sách để xóa!"); return; }
            if (!UI.Confirm("Xóa nhân viên '" + txtTen.Text + "' (và tài khoản đăng nhập nếu có)?")) return;
            try
            {
                DBHelper.Exec(
                    "SET XACT_ABORT ON; BEGIN TRAN; " +
                    "DELETE FROM tblUser WHERE MaNhanVien = @ma; " +
                    "DELETE FROM tblNhanVien WHERE MaNhanVien = @ma; COMMIT;",
                    DBHelper.P("@ma", txtMa.Text.Trim()));
                UI.Info("Đã xóa.");
                LoadGrid();
                ClearInput();
            }
            catch (SqlException ex) { UI.ShowSql(ex, "xóa"); }
        }
    }
}
