using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace QuanLyBanHang
{
    /// <summary>Các hàm tiện ích tạo control bằng code và hiển thị thông báo.</summary>
    public static class UI
    {
        public static Label Lbl(Control p, string text, int x, int y, int w = 100)
        {
            var l = new Label { Text = text, Left = x, Top = y + 4, Width = w, Height = 24, AutoSize = false };
            p.Controls.Add(l);
            return l;
        }

        public static TextBox Txt(Control p, int x, int y, int w = 180)
        {
            var t = new TextBox { Left = x, Top = y, Width = w };
            p.Controls.Add(t);
            return t;
        }

        public static ComboBox Cbo(Control p, int x, int y, int w = 180)
        {
            var c = new ComboBox { Left = x, Top = y, Width = w, DropDownStyle = ComboBoxStyle.DropDownList };
            p.Controls.Add(c);
            return c;
        }

        public static Button Btn(Control p, string text, int x, int y, int w, EventHandler click)
        {
            var b = new Button { Text = text, Left = x, Top = y, Width = w, Height = 32 };
            if (click != null) b.Click += click;
            p.Controls.Add(b);
            return b;
        }

        public static DataGridView Grid(Control p, int x, int y, int w, int h)
        {
            var g = new DataGridView
            {
                Left = x, Top = y, Width = w, Height = h,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            p.Controls.Add(g);
            return g;
        }

        /// <summary>Đặt tiêu đề cột (và định dạng số/ngày) cho DataGridView sau khi gán DataSource.</summary>
        public static void Col(DataGridView g, string name, string header, string fmt = null)
        {
            if (!g.Columns.Contains(name)) return;
            var c = g.Columns[name];
            c.HeaderText = header;
            if (fmt != null)
            {
                c.DefaultCellStyle.Format = fmt;
                if (fmt.StartsWith("N")) c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        public static void BindCombo(ComboBox c, DataTable dt, string display, string value)
        {
            c.DataSource = null;
            c.DisplayMember = display;
            c.ValueMember = value;
            c.DataSource = dt;
        }

        /// <summary>Đọc tiền: chấp nhận "100000", "100.000", "100,000".</summary>
        public static bool ParseMoney(string s, out double v)
        {
            s = (s ?? "").Replace(".", "").Replace(",", "").Replace(" ", "");
            return double.TryParse(s, out v) && v >= 0;
        }

        /// <summary>Sinh mã tiếp theo, ví dụ H10 -> H11 (prefix "H", digits 2).</summary>
        public static string NextId(string table, string col, string prefix, int digits)
        {
            string sql = "SELECT ISNULL(MAX(TRY_CAST(SUBSTRING(" + col + ", @n, 20) AS int)), 0) + 1 FROM " + table +
                         " WHERE " + col + " LIKE @pat";
            object o = DBHelper.Scalar(sql,
                DBHelper.P("@n", prefix.Length + 1),
                DBHelper.P("@pat", prefix + "[0-9]%"));
            return prefix + Convert.ToInt32(o).ToString("D" + digits);
        }

        public static void Info(string msg)
        {
            MessageBox.Show(msg, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void Warn(string msg)
        {
            MessageBox.Show(msg, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static void Err(string msg)
        {
            MessageBox.Show(msg, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static bool Confirm(string msg)
        {
            return MessageBox.Show(msg, "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        /// <summary>Dịch lỗi SQL Server thành thông báo dễ hiểu (2627/2601: trùng khóa, 547: vi phạm khóa ngoại).</summary>
        public static void ShowSql(SqlException ex, string action)
        {
            if (ex.Number == 2627 || ex.Number == 2601)
                Warn("Mã đã tồn tại, vui lòng nhập mã khác.");
            else if (ex.Number == 547)
                Warn("Không thể " + action + " vì dữ liệu đang liên quan đến bảng khác (hóa đơn, loại hàng...).");
            else
                Err("Lỗi CSDL: " + ex.Message);
        }
    }
}
