using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace QuanLyBanHang
{
    public class frmKhach : Form
    {
        TextBox txtMa, txtTen, txtDiaChi, txtDienThoai, txtTim;
        DataGridView dgv;

        public frmKhach()
        {
            Text = "Quản lý khách hàng";
            Font = new Font("Segoe UI", 10F);
            ClientSize = new Size(900, 600);
            StartPosition = FormStartPosition.CenterParent;

            UI.Lbl(this, "Mã khách", 15, 15);
            txtMa = UI.Txt(this, 110, 15, 150);
            UI.Lbl(this, "Tên khách", 290, 15, 80);
            txtTen = UI.Txt(this, 380, 15, 250);

            UI.Lbl(this, "Địa chỉ", 15, 55);
            txtDiaChi = UI.Txt(this, 110, 55, 150);
            UI.Lbl(this, "Điện thoại", 290, 55, 80);
            txtDienThoai = UI.Txt(this, 380, 55, 250);

            UI.Btn(this, "Thêm", 15, 100, 95, (s, e) => Them());
            UI.Btn(this, "Sửa", 120, 100, 95, (s, e) => Sua());
            UI.Btn(this, "Xóa", 225, 100, 95, (s, e) => Xoa());
            UI.Btn(this, "Làm mới", 330, 100, 95, (s, e) => ClearInput());

            UI.Lbl(this, "Tìm kiếm", 15, 150, 90);
            txtTim = UI.Txt(this, 110, 150, 300);
            UI.Btn(this, "Tìm", 425, 148, 80, (s, e) => LoadGrid());
            txtTim.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadGrid(); };
            var gy = new Label { Text = "(tìm theo mã, tên, địa chỉ, số điện thoại)", Left = 520, Top = 154, Width = 350, ForeColor = Color.Gray };
            Controls.Add(gy);

            dgv = UI.Grid(this, 15, 195, 870, 390);
            dgv.CellClick += Dgv_CellClick;

            Load += (s, e) => { LoadGrid(); ClearInput(); };
        }

        void LoadGrid()
        {
            string k = "%" + txtTim.Text.Trim() + "%";
            dgv.DataSource = DBHelper.Query(
                "SELECT MaKhach, TenKhach, DiaChi, DienThoai FROM tblKhach " +
                "WHERE MaKhach LIKE @k OR TenKhach LIKE @k OR DiaChi LIKE @k OR DienThoai LIKE @k ORDER BY MaKhach",
                DBHelper.P("@k", k));
            UI.Col(dgv, "MaKhach", "Mã khách");
            UI.Col(dgv, "TenKhach", "Tên khách");
            UI.Col(dgv, "DiaChi", "Địa chỉ");
            UI.Col(dgv, "DienThoai", "Điện thoại");
        }

        void Dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            DataGridViewRow r = dgv.Rows[e.RowIndex];
            txtMa.Text = Convert.ToString(r.Cells["MaKhach"].Value);
            txtMa.Enabled = false;
            txtTen.Text = Convert.ToString(r.Cells["TenKhach"].Value);
            txtDiaChi.Text = Convert.ToString(r.Cells["DiaChi"].Value);
            txtDienThoai.Text = Convert.ToString(r.Cells["DienThoai"].Value);
        }

        void ClearInput()
        {
            txtMa.Enabled = true;
            txtMa.Text = UI.NextId("tblKhach", "MaKhach", "KH", 2);
            txtTen.Clear();
            txtDiaChi.Clear();
            txtDienThoai.Clear();
            txtTen.Focus();
        }

        bool Check()
        {
            if (txtMa.Text.Trim() == "") { UI.Warn("Chưa nhập mã khách!"); return false; }
            if (txtTen.Text.Trim() == "") { UI.Warn("Chưa nhập tên khách!"); txtTen.Focus(); return false; }
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
                DBHelper.P("@dc", txtDiaChi.Text.Trim()),
                DBHelper.P("@sdt", txtDienThoai.Text.Trim())
            };
        }

        void Them()
        {
            if (!Check()) return;
            try
            {
                DBHelper.Exec("INSERT INTO tblKhach(MaKhach, TenKhach, DiaChi, DienThoai) VALUES(@ma, @ten, @dc, @sdt)", Params());
                UI.Info("Thêm khách hàng thành công.");
                LoadGrid();
                ClearInput();
            }
            catch (SqlException ex) { UI.ShowSql(ex, "thêm"); }
        }

        void Sua()
        {
            if (txtMa.Enabled) { UI.Warn("Hãy chọn một khách hàng trong danh sách để sửa!"); return; }
            if (!Check()) return;
            try
            {
                DBHelper.Exec("UPDATE tblKhach SET TenKhach=@ten, DiaChi=@dc, DienThoai=@sdt WHERE MaKhach=@ma", Params());
                UI.Info("Cập nhật thành công.");
                LoadGrid();
            }
            catch (SqlException ex) { UI.ShowSql(ex, "sửa"); }
        }

        void Xoa()
        {
            if (txtMa.Enabled) { UI.Warn("Hãy chọn một khách hàng trong danh sách để xóa!"); return; }
            if (!UI.Confirm("Bạn có chắc muốn xóa khách '" + txtTen.Text + "'?")) return;
            try
            {
                DBHelper.Exec("DELETE FROM tblKhach WHERE MaKhach=@ma", DBHelper.P("@ma", txtMa.Text.Trim()));
                UI.Info("Đã xóa.");
                LoadGrid();
                ClearInput();
            }
            catch (SqlException ex) { UI.ShowSql(ex, "xóa"); }
        }
    }
}
