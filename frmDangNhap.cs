using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace QuanLyBanHang
{
    public class frmDangNhap : Form
    {
        TextBox txtUser, txtPass;
        RadioButton rdoAdmin, rdoNV;
        Button btnDN, btnThoat;

        public frmDangNhap()
        {
            Text = "Đăng nhập hệ thống";
            Font = new Font("Segoe UI", 10F);
            ClientSize = new Size(400, 320);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            var title = new Label
            {
                Text = "QUẢN LÝ BÁN HÀNG",
                Font = new Font("Segoe UI", 16F, FontStyle.Bold),
                ForeColor = Color.FromArgb(40, 30, 120),
                TextAlign = ContentAlignment.MiddleCenter,
                Left = 0, Top = 20, Width = 400, Height = 40
            };
            Controls.Add(title);

            var grp = new GroupBox { Text = "Đăng nhập với vai trò", Left = 30, Top = 75, Width = 340, Height = 60 };
            rdoAdmin = new RadioButton { Text = "Admin", Left = 40, Top = 25, Width = 100 };
            rdoNV = new RadioButton { Text = "Nhân viên", Left = 170, Top = 25, Width = 120, Checked = true };
            grp.Controls.Add(rdoAdmin);
            grp.Controls.Add(rdoNV);
            Controls.Add(grp);

            UI.Lbl(this, "Tài khoản", 30, 155, 90);
            txtUser = UI.Txt(this, 130, 155, 240);
            UI.Lbl(this, "Mật khẩu", 30, 195, 90);
            txtPass = UI.Txt(this, 130, 195, 240);
            txtPass.UseSystemPasswordChar = true;

            btnDN = UI.Btn(this, "Đăng nhập", 80, 250, 120, (s, e) => DangNhap());
            btnThoat = UI.Btn(this, "Thoát", 220, 250, 100, (s, e) => { DialogResult = DialogResult.Cancel; Close(); });

            AcceptButton = btnDN;      // nhấn Enter = Đăng nhập
            CancelButton = btnThoat;   // nhấn Esc = Thoát
        }

        void DangNhap()
        {
            string user = txtUser.Text.Trim();
            string pass = txtPass.Text;
            string quyen = rdoAdmin.Checked ? "Admin" : "NhanVien";

            if (user == "" || pass == "")
            {
                UI.Warn("Vui lòng nhập đầy đủ tài khoản và mật khẩu!");
                return;
            }

            try
            {
                // COLLATE ...CS_AS: so sánh mật khẩu phân biệt hoa/thường
                DataTable dt = DBHelper.Query(
                    "SELECT userName, Quyen, MaNhanVien FROM tblUser " +
                    "WHERE userName = @u AND passWord = @p COLLATE Latin1_General_CS_AS AND Quyen = @q",
                    DBHelper.P("@u", user), DBHelper.P("@p", pass), DBHelper.P("@q", quyen));

                if (dt.Rows.Count == 0)
                {
                    UI.Warn("Sai tài khoản, mật khẩu hoặc vai trò!");
                    txtPass.Clear();
                    txtPass.Focus();
                    return;
                }

                Session.UserName = dt.Rows[0]["userName"].ToString();
                Session.Quyen = dt.Rows[0]["Quyen"].ToString();
                Session.MaNhanVien = dt.Rows[0]["MaNhanVien"] as string;   // admin => null

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (SqlException ex)
            {
                UI.Err("Không kết nối được cơ sở dữ liệu.\nHãy kiểm tra chuỗi kết nối trong App.config và đã chạy file SQL chưa.\n\nChi tiết: " + ex.Message);
            }
        }
    }
}
