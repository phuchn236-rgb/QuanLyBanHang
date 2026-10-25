using System;
using System.Windows.Forms;

namespace QuanLyBanHang
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Vòng lặp: Đăng nhập -> Giao diện chính -> (Đăng xuất thì quay lại Đăng nhập)
            while (true)
            {
                using (var login = new frmDangNhap())
                {
                    if (login.ShowDialog() != DialogResult.OK) break;   // bấm Thoát
                }

                var main = new frmMain();
                Application.Run(main);

                if (!main.Logout) break;   // đóng cửa sổ chính = thoát chương trình
                Session.Clear();
            }
        }
    }
}
