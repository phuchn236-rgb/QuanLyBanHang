using System;
using System.Drawing;
using System.Windows.Forms;

namespace QuanLyBanHang
{
    public class frmMain : Form
    {
        public bool Logout;   // true nếu người dùng chọn "Đăng xuất"

        public frmMain()
        {
            Text = "Quản lý bán hàng";
            Font = new Font("Segoe UI", 10F);
            IsMdiContainer = true;
            WindowState = FormWindowState.Maximized;

            // ---------- Thanh trạng thái ----------
            var ss = new StatusStrip();
            ss.Items.Add(new ToolStripStatusLabel(
                "Người dùng: " + Session.UserName +
                "   |   Quyền: " + (Session.IsAdmin ? "Quản trị (Admin)" : "Nhân viên") +
                "   |   Ngày: " + DateTime.Today.ToString("dd/MM/yyyy")));

            // ---------- Menu ----------
            var ms = new MenuStrip();

            var mHeThong = new ToolStripMenuItem("Hệ thống");
            mHeThong.DropDownItems.Add("Đăng xuất", null, (s, e) =>
            {
                if (UI.Confirm("Bạn có chắc muốn đăng xuất?")) { Logout = true; Close(); }
            });
            mHeThong.DropDownItems.Add("Thoát", null, (s, e) => Close());

            var mDanhMuc = new ToolStripMenuItem("Danh mục");
            mDanhMuc.DropDownItems.Add("Hàng hóa", null, (s, e) => OpenChild<frmHang>());
            mDanhMuc.DropDownItems.Add("Khách hàng", null, (s, e) => OpenChild<frmKhach>());
            var miNhanVien = new ToolStripMenuItem("Nhân viên", null, (s, e) => OpenChild<frmNhanVien>());
            miNhanVien.Visible = Session.IsAdmin;          // chỉ Admin
            mDanhMuc.DropDownItems.Add(miNhanVien);

            var mBanHang = new ToolStripMenuItem("Bán hàng");
            mBanHang.DropDownItems.Add("Lập hóa đơn", null, (s, e) => OpenChild<frmHoaDon>());

            var mThongKe = new ToolStripMenuItem("Thống kê & Báo cáo", null, (s, e) => OpenChild<frmThongKe>());
            mThongKe.Visible = Session.IsAdmin;            // chỉ Admin

            var mCuaSo = new ToolStripMenuItem("Cửa sổ");
            mCuaSo.DropDownItems.Add("Xếp chồng", null, (s, e) => LayoutMdi(MdiLayout.Cascade));
            mCuaSo.DropDownItems.Add("Xếp ngang", null, (s, e) => LayoutMdi(MdiLayout.TileHorizontal));
            mCuaSo.DropDownItems.Add("Xếp dọc", null, (s, e) => LayoutMdi(MdiLayout.TileVertical));
            mCuaSo.DropDownItems.Add("Đóng tất cả", null, (s, e) =>
            {
                foreach (Form f in MdiChildren) f.Close();
            });

            ms.Items.AddRange(new ToolStripItem[] { mHeThong, mDanhMuc, mBanHang, mThongKe, mCuaSo });
            MainMenuStrip = ms;

            // ---------- Thanh công cụ (truy cập nhanh) ----------
            var ts = new ToolStrip { ImageScalingSize = new Size(24, 24) };
            ts.Items.Add(new ToolStripButton("Hàng hóa", null, (s, e) => OpenChild<frmHang>()));
            ts.Items.Add(new ToolStripButton("Khách hàng", null, (s, e) => OpenChild<frmKhach>()));
            if (Session.IsAdmin)
                ts.Items.Add(new ToolStripButton("Nhân viên", null, (s, e) => OpenChild<frmNhanVien>()));
            ts.Items.Add(new ToolStripButton("Lập hóa đơn", null, (s, e) => OpenChild<frmHoaDon>()));
            if (Session.IsAdmin)
                ts.Items.Add(new ToolStripButton("Thống kê", null, (s, e) => OpenChild<frmThongKe>()));

            // Thứ tự Add: control thêm sau nằm trên cùng
            Controls.Add(ss);
            Controls.Add(ts);
            Controls.Add(ms);
        }

        /// <summary>Mở form con; nếu đã mở rồi thì chỉ kích hoạt lại, không mở trùng.</summary>
        void OpenChild<T>() where T : Form, new()
        {
            foreach (Form f in MdiChildren)
            {
                if (f is T) { f.Activate(); return; }
            }
            var frm = new T();
            frm.MdiParent = this;
            frm.Show();
        }
    }
}
