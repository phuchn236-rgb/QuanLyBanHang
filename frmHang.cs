using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace QuanLyBanHang
{
    public class frmHang : Form
    {
        TextBox txtMa, txtTen, txtGiaNhap, txtGiaBan, txtAnh, txtGhiChu, txtTim;
        ComboBox cboLoai, cboLoc;
        NumericUpDown nudSL;
        PictureBox pic;
        DataGridView dgv;
        Button btnThem, btnSua, btnXoa, btnMoi, btnAnh, btnTim;

        // Ảnh được sao chép vào thư mục Images cạnh file .exe; CSDL chỉ lưu TÊN FILE
        static readonly string ImgDir = Path.Combine(Application.StartupPath, "Images");

        public frmHang()
        {
            Text = "Quản lý hàng hóa";
            Font = new Font("Segoe UI", 10F);
            ClientSize = new Size(1100, 680);
            StartPosition = FormStartPosition.CenterParent;

            UI.Lbl(this, "Mã hàng", 15, 15);
            txtMa = UI.Txt(this, 110, 15, 150);
            UI.Lbl(this, "Tên hàng", 290, 15, 80);
            txtTen = UI.Txt(this, 380, 15, 230);

            UI.Lbl(this, "Loại hàng", 15, 55);
            cboLoai = UI.Cbo(this, 110, 55, 150);
            UI.Lbl(this, "Số lượng", 290, 55, 80);
            nudSL = new NumericUpDown { Left = 380, Top = 55, Width = 100, Maximum = 1000000 };
            Controls.Add(nudSL);

            UI.Lbl(this, "Đơn giá nhập", 15, 95);
            txtGiaNhap = UI.Txt(this, 110, 95, 150);
            UI.Lbl(this, "Đơn giá bán", 290, 95, 90);
            txtGiaBan = UI.Txt(this, 380, 95, 150);

            UI.Lbl(this, "Ảnh (tên file)", 15, 135);
            txtAnh = UI.Txt(this, 110, 135, 150);
            txtAnh.ReadOnly = true;
            btnAnh = UI.Btn(this, "Chọn ảnh...", 270, 133, 100, (s, e) => ChonAnh());

            UI.Lbl(this, "Ghi chú", 15, 175);
            txtGhiChu = UI.Txt(this, 110, 175, 500);

            pic = new PictureBox
            {
                Left = 660, Top = 15, Width = 160, Height = 160,
                BorderStyle = BorderStyle.FixedSingle, SizeMode = PictureBoxSizeMode.Zoom
            };
            Controls.Add(pic);

            btnThem = UI.Btn(this, "Thêm", 15, 220, 95, (s, e) => Them());
            btnSua = UI.Btn(this, "Sửa", 120, 220, 95, (s, e) => Sua());
            btnXoa = UI.Btn(this, "Xóa", 225, 220, 95, (s, e) => Xoa());
            btnMoi = UI.Btn(this, "Làm mới", 330, 220, 95, (s, e) => ClearInput());

            UI.Lbl(this, "Tìm kiếm", 15, 270, 90);
            txtTim = UI.Txt(this, 110, 270, 300);
            UI.Lbl(this, "Loại", 430, 270, 40);
            cboLoc = UI.Cbo(this, 475, 270, 170);
            btnTim = UI.Btn(this, "Tìm", 660, 268, 80, (s, e) => LoadGrid());
            var lblGoiY = new Label
            {
                Text = "(tìm theo mã, tên, ghi chú: công dụng / thành phần / hạn sử dụng)",
                Left = 15, Top = 305, Width = 700, ForeColor = Color.Gray
            };
            Controls.Add(lblGoiY);
            txtTim.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) LoadGrid(); };
            cboLoc.SelectionChangeCommitted += (s, e) => LoadGrid();

            dgv = UI.Grid(this, 15, 335, 1070, 330);
            dgv.CellClick += Dgv_CellClick;

            // Nhân viên chỉ được xem và tìm kiếm
            if (!Session.IsAdmin)
            {
                Text += " (chỉ xem)";
                txtMa.ReadOnly = txtTen.ReadOnly = txtGiaNhap.ReadOnly = txtGiaBan.ReadOnly = txtGhiChu.ReadOnly = true;
                cboLoai.Enabled = false;
                nudSL.Enabled = false;
                btnThem.Enabled = btnSua.Enabled = btnXoa.Enabled = btnAnh.Enabled = false;
            }

            Load += (s, e) =>
            {
                LoadLoai();
                LoadGrid();
                ClearInput();
            };
        }

        void LoadLoai()
        {
            DataTable dt = DBHelper.Query("SELECT MaLoaiHang, TenLoaiHang FROM tblLoaiHang ORDER BY MaLoaiHang");
            UI.BindCombo(cboLoai, dt, "TenLoaiHang", "MaLoaiHang");

            DataTable dt2 = dt.Copy();
            DataRow all = dt2.NewRow();
            all["MaLoaiHang"] = "";
            all["TenLoaiHang"] = "(Tất cả)";
            dt2.Rows.InsertAt(all, 0);
            UI.BindCombo(cboLoc, dt2, "TenLoaiHang", "MaLoaiHang");
        }

        void LoadGrid()
        {
            string k = "%" + txtTim.Text.Trim() + "%";
            string loai = cboLoc.SelectedValue == null ? "" : cboLoc.SelectedValue.ToString();

            DataTable dt = DBHelper.Query(
                "SELECT h.MaHang, h.TenHang, l.TenLoaiHang, h.MaLoaiHang, h.SoLuong, h.DonGiaNhap, h.DonGiaBan, h.Anh, h.GhiChu " +
                "FROM tblHang h LEFT JOIN tblLoaiHang l ON l.MaLoaiHang = h.MaLoaiHang " +
                "WHERE (h.MaHang LIKE @k OR h.TenHang LIKE @k OR h.GhiChu LIKE @k) " +
                "AND (@loai = '' OR h.MaLoaiHang = @loai) " +
                "ORDER BY h.MaHang",
                DBHelper.P("@k", k), DBHelper.P("@loai", loai));

            dgv.DataSource = dt;
            dgv.Columns["MaLoaiHang"].Visible = false;
            UI.Col(dgv, "MaHang", "Mã hàng");
            UI.Col(dgv, "TenHang", "Tên hàng");
            UI.Col(dgv, "TenLoaiHang", "Loại hàng");
            UI.Col(dgv, "SoLuong", "Số lượng", "N0");
            UI.Col(dgv, "DonGiaNhap", "Đơn giá nhập", "N0");
            UI.Col(dgv, "DonGiaBan", "Đơn giá bán", "N0");
            UI.Col(dgv, "Anh", "Ảnh");
            UI.Col(dgv, "GhiChu", "Ghi chú");
        }

        void Dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            DataGridViewRow r = dgv.Rows[e.RowIndex];

            txtMa.Text = Convert.ToString(r.Cells["MaHang"].Value);
            txtMa.Enabled = false;                       // khóa chính không cho sửa
            txtTen.Text = Convert.ToString(r.Cells["TenHang"].Value);
            cboLoai.SelectedValue = Convert.ToString(r.Cells["MaLoaiHang"].Value);

            decimal sl;
            decimal.TryParse(Convert.ToString(r.Cells["SoLuong"].Value), out sl);
            nudSL.Value = Math.Min(Math.Max(sl, nudSL.Minimum), nudSL.Maximum);

            txtGiaNhap.Text = Convert.ToString(r.Cells["DonGiaNhap"].Value);
            txtGiaBan.Text = Convert.ToString(r.Cells["DonGiaBan"].Value);
            txtAnh.Text = Convert.ToString(r.Cells["Anh"].Value);
            txtGhiChu.Text = Convert.ToString(r.Cells["GhiChu"].Value);
            ShowPic(txtAnh.Text);
        }

        void ClearInput()
        {
            txtMa.Enabled = true;
            txtMa.Text = Session.IsAdmin ? UI.NextId("tblHang", "MaHang", "H", 2) : "";
            txtTen.Clear();
            txtGiaNhap.Clear();
            txtGiaBan.Clear();
            txtAnh.Clear();
            txtGhiChu.Clear();
            nudSL.Value = 0;
            if (cboLoai.Items.Count > 0) cboLoai.SelectedIndex = 0;
            ShowPic("");
            txtTen.Focus();
        }

        void ShowPic(string fileName)
        {
            if (pic.Image != null) { pic.Image.Dispose(); pic.Image = null; }
            if (string.IsNullOrWhiteSpace(fileName)) return;

            string path = Path.Combine(ImgDir, fileName);
            if (!File.Exists(path)) return;

            // Copy sang Bitmap mới để không bị khóa file ảnh
            using (var tmp = Image.FromFile(path))
            {
                pic.Image = new Bitmap(tmp);
            }
        }

        void ChonAnh()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Ảnh|*.jpg;*.jpeg;*.png;*.bmp;*.gif";
                if (dlg.ShowDialog() != DialogResult.OK) return;

                Directory.CreateDirectory(ImgDir);
                string fileName = Path.GetFileName(dlg.FileName);
                string dest = Path.Combine(ImgDir, fileName);
                if (!string.Equals(Path.GetFullPath(dlg.FileName), Path.GetFullPath(dest), StringComparison.OrdinalIgnoreCase))
                    File.Copy(dlg.FileName, dest, true);

                txtAnh.Text = fileName;       // chỉ lưu tên file, không lưu đường dẫn tuyệt đối
                ShowPic(fileName);
            }
        }

        bool Check(out double gn, out double gb)
        {
            gn = gb = 0;
            if (txtMa.Text.Trim() == "") { UI.Warn("Chưa nhập mã hàng!"); txtMa.Focus(); return false; }
            if (txtTen.Text.Trim() == "") { UI.Warn("Chưa nhập tên hàng!"); txtTen.Focus(); return false; }
            if (cboLoai.SelectedValue == null) { UI.Warn("Chưa chọn loại hàng!"); return false; }
            if (!UI.ParseMoney(txtGiaNhap.Text, out gn)) { UI.Warn("Đơn giá nhập không hợp lệ!"); txtGiaNhap.Focus(); return false; }
            if (!UI.ParseMoney(txtGiaBan.Text, out gb)) { UI.Warn("Đơn giá bán không hợp lệ!"); txtGiaBan.Focus(); return false; }
            return true;
        }

        SqlParameter[] Params(double gn, double gb)
        {
            return new[]
            {
                DBHelper.P("@ma", txtMa.Text.Trim()),
                DBHelper.P("@ten", txtTen.Text.Trim()),
                DBHelper.P("@loai", cboLoai.SelectedValue),
                DBHelper.P("@sl", (int)nudSL.Value),
                DBHelper.P("@gn", gn),
                DBHelper.P("@gb", gb),
                DBHelper.P("@anh", string.IsNullOrWhiteSpace(txtAnh.Text) ? null : txtAnh.Text.Trim()),
                DBHelper.P("@gc", string.IsNullOrWhiteSpace(txtGhiChu.Text) ? null : txtGhiChu.Text.Trim())
            };
        }

        void Them()
        {
            double gn, gb;
            if (!Check(out gn, out gb)) return;
            try
            {
                DBHelper.Exec(
                    "INSERT INTO tblHang(MaHang, TenHang, MaLoaiHang, SoLuong, DonGiaNhap, DonGiaBan, Anh, GhiChu) " +
                    "VALUES(@ma, @ten, @loai, @sl, @gn, @gb, @anh, @gc)", Params(gn, gb));
                UI.Info("Thêm hàng thành công.");
                LoadGrid();
                ClearInput();
            }
            catch (SqlException ex) { UI.ShowSql(ex, "thêm"); }
        }

        void Sua()
        {
            if (txtMa.Enabled) { UI.Warn("Hãy chọn một hàng trong danh sách để sửa!"); return; }
            double gn, gb;
            if (!Check(out gn, out gb)) return;
            try
            {
                int n = DBHelper.Exec(
                    "UPDATE tblHang SET TenHang=@ten, MaLoaiHang=@loai, SoLuong=@sl, DonGiaNhap=@gn, DonGiaBan=@gb, " +
                    "Anh=@anh, GhiChu=@gc WHERE MaHang=@ma", Params(gn, gb));
                if (n == 0) { UI.Warn("Không tìm thấy hàng cần sửa!"); return; }
                UI.Info("Cập nhật thành công.");
                LoadGrid();
            }
            catch (SqlException ex) { UI.ShowSql(ex, "sửa"); }
        }

        void Xoa()
        {
            if (txtMa.Enabled) { UI.Warn("Hãy chọn một hàng trong danh sách để xóa!"); return; }
            if (!UI.Confirm("Bạn có chắc muốn xóa hàng '" + txtTen.Text + "'?")) return;
            try
            {
                DBHelper.Exec("DELETE FROM tblHang WHERE MaHang=@ma", DBHelper.P("@ma", txtMa.Text.Trim()));
                UI.Info("Đã xóa.");
                LoadGrid();
                ClearInput();
            }
            catch (SqlException ex) { UI.ShowSql(ex, "xóa"); }
        }
    }
}
