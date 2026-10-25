using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using ClosedXML.Excel;

namespace QuanLyBanHang
{
    public class frmThongKe : Form
    {
        ComboBox cboKieu;
        NumericUpDown nudNam;
        DataGridView dgv;
        Chart chart;
        Label lblTong;
        DataTable dtKetQua;

        public frmThongKe()
        {
            Text = "Thống kê & Báo cáo doanh thu - lãi/lỗ";
            Font = new Font("Segoe UI", 10F);
            ClientSize = new Size(1100, 720);
            StartPosition = FormStartPosition.CenterParent;

            UI.Lbl(this, "Thống kê theo", 15, 15, 110);
            cboKieu = UI.Cbo(this, 130, 15, 180);
            cboKieu.Items.AddRange(new object[] { "Theo tháng", "Theo quý", "Theo năm", "Theo mặt hàng" });
            cboKieu.SelectedIndex = 0;

            UI.Lbl(this, "Năm", 330, 15, 40);
            nudNam = new NumericUpDown { Left = 375, Top = 15, Width = 90, Minimum = 2000, Maximum = 2100, Value = DateTime.Today.Year };
            Controls.Add(nudNam);

            UI.Btn(this, "Xem báo cáo", 490, 12, 120, (s, e) => Xem());
            UI.Btn(this, "Xuất Excel", 620, 12, 110, (s, e) => XuatExcel());
            cboKieu.SelectedIndexChanged += (s, e) => nudNam.Enabled = cboKieu.SelectedIndex != 2; // "Theo năm" xem tất cả các năm

            dgv = UI.Grid(this, 15, 60, 1070, 220);
            dgv.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            lblTong = new Label { Left = 15, Top = 290, Width = 1070, Font = new Font("Segoe UI", 11F, FontStyle.Bold), ForeColor = Color.Firebrick };
            Controls.Add(lblTong);

            chart = new Chart { Left = 15, Top = 325, Width = 1070, Height = 380, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom };
            chart.ChartAreas.Add(new ChartArea("area"));
            chart.Legends.Add(new Legend("legend"));
            Controls.Add(chart);

            Load += (s, e) =>
            {
                if (!Session.IsAdmin)
                {
                    UI.Warn("Chỉ Admin mới được xem thống kê!");
                    Close();
                    return;
                }
                Xem();
            };
        }

        void Xem()
        {
            string keyExpr, sortExpr, where;
            switch (cboKieu.SelectedIndex)
            {
                case 0:   // theo tháng
                    keyExpr = "N'Tháng ' + RIGHT('0' + CAST(MONTH(b.NgayBan) AS nvarchar(2)), 2)";
                    sortExpr = "MONTH(b.NgayBan)";
                    where = "YEAR(b.NgayBan) = @y";
                    break;
                case 1:   // theo quý
                    keyExpr = "N'Quý ' + CAST(DATEPART(QUARTER, b.NgayBan) AS nvarchar(1))";
                    sortExpr = "DATEPART(QUARTER, b.NgayBan)";
                    where = "YEAR(b.NgayBan) = @y";
                    break;
                case 2:   // theo năm
                    keyExpr = "N'Năm ' + CAST(YEAR(b.NgayBan) AS nvarchar(4))";
                    sortExpr = "YEAR(b.NgayBan)";
                    where = "1 = 1";
                    break;
                default:  // theo mặt hàng
                    keyExpr = "h.TenHang";
                    sortExpr = "h.MaHang";
                    where = "YEAR(b.NgayBan) = @y";
                    break;
            }

            // Doanh thu = tổng Thành tiền; Vốn = SL x Đơn giá nhập; Lãi/Lỗ = Doanh thu - Vốn
            string sql =
                "SELECT x.Ky AS [Kỳ], SUM(x.ThanhTien) AS [Doanh thu], SUM(x.Von) AS [Vốn], " +
                "SUM(x.ThanhTien - x.Von) AS [Lãi/Lỗ], " +
                "CASE WHEN SUM(x.ThanhTien - x.Von) >= 0 THEN N'Lãi' ELSE N'Lỗ' END AS [Kết quả] " +
                "FROM (SELECT " + keyExpr + " AS Ky, " + sortExpr + " AS Sx, ct.ThanhTien, ct.SoLuong * h.DonGiaNhap AS Von " +
                "      FROM tblHDBan b " +
                "      JOIN tblChiTietHDBan ct ON ct.MaHDBan = b.MaHDBan " +
                "      JOIN tblHang h ON h.MaHang = ct.MaHang " +
                "      WHERE " + where + ") x " +
                "GROUP BY x.Ky, x.Sx ORDER BY x.Sx";

            try
            {
                dtKetQua = DBHelper.Query(sql, DBHelper.P("@y", (int)nudNam.Value));
            }
            catch (SqlException ex)
            {
                UI.Err("Lỗi CSDL: " + ex.Message);
                return;
            }

            dgv.DataSource = dtKetQua;
            UI.Col(dgv, "Doanh thu", "Doanh thu", "N0");
            UI.Col(dgv, "Vốn", "Vốn", "N0");
            UI.Col(dgv, "Lãi/Lỗ", "Lãi/Lỗ", "N0");

            double dt = 0, lai = 0;
            foreach (DataRow r in dtKetQua.Rows)
            {
                dt += Convert.ToDouble(r["Doanh thu"]);
                lai += Convert.ToDouble(r["Lãi/Lỗ"]);
            }
            lblTong.Text = "Tổng doanh thu: " + dt.ToString("N0") + " VNĐ     |     Tổng " +
                           (lai >= 0 ? "lãi: " : "lỗ: ") + Math.Abs(lai).ToString("N0") + " VNĐ";

            VeBieuDo();
        }

        void VeBieuDo()
        {
            chart.Series.Clear();
            chart.Titles.Clear();
            chart.Titles.Add("Biểu đồ doanh thu và lãi/lỗ - " + cboKieu.Text.ToLower());

            // Theo thời gian: biểu đồ đường (thể hiện xu hướng tăng trưởng); theo mặt hàng: biểu đồ cột
            SeriesChartType kieu = cboKieu.SelectedIndex == 3 ? SeriesChartType.Column : SeriesChartType.Line;

            var s1 = new Series("Doanh thu") { ChartType = kieu, ChartArea = "area", Legend = "legend", BorderWidth = 3 };
            var s2 = new Series("Lãi/Lỗ") { ChartType = kieu, ChartArea = "area", Legend = "legend", BorderWidth = 3 };
            if (kieu == SeriesChartType.Line)
            {
                s1.MarkerStyle = MarkerStyle.Circle; s1.MarkerSize = 8;
                s2.MarkerStyle = MarkerStyle.Diamond; s2.MarkerSize = 8;
            }

            foreach (DataRow r in dtKetQua.Rows)
            {
                s1.Points.AddXY(r["Kỳ"].ToString(), Convert.ToDouble(r["Doanh thu"]));
                s2.Points.AddXY(r["Kỳ"].ToString(), Convert.ToDouble(r["Lãi/Lỗ"]));
            }
            chart.Series.Add(s1);
            chart.Series.Add(s2);
            chart.ChartAreas["area"].AxisY.LabelStyle.Format = "N0";
            chart.ChartAreas["area"].AxisX.Interval = 1;
        }

        void XuatExcel()
        {
            if (dtKetQua == null || dtKetQua.Rows.Count == 0) { UI.Warn("Không có dữ liệu để xuất!"); return; }

            using (var dlg = new SaveFileDialog())
            {
                dlg.Filter = "Excel (*.xlsx)|*.xlsx";
                dlg.FileName = "BaoCao_" + cboKieu.Text.Replace(" ", "") + "_" + DateTime.Today.ToString("yyyyMMdd") + ".xlsx";
                if (dlg.ShowDialog() != DialogResult.OK) return;

                try
                {
                    using (var wb = new XLWorkbook())
                    {
                        IXLWorksheet ws = wb.Worksheets.Add("BaoCao");
                        ws.Cell(1, 1).Value = "BÁO CÁO DOANH THU - LÃI/LỖ (" + cboKieu.Text.ToUpper() + ")";
                        ws.Cell(1, 1).Style.Font.Bold = true;
                        ws.Cell(1, 1).Style.Font.FontSize = 14;
                        ws.Cell(2, 1).Value = "Ngày lập: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                        // Bảng dữ liệu bắt đầu từ ô A4
                        ws.Cell(4, 1).InsertTable(dtKetQua, "tblBaoCao", true);
                        int last = 4 + dtKetQua.Rows.Count;
                        ws.Range(5, 2, last, 4).Style.NumberFormat.Format = "#,##0";
                        ws.Columns().AdjustToContents();
                        wb.SaveAs(dlg.FileName);
                    }
                    UI.Info("Đã xuất báo cáo ra Excel:\n" + dlg.FileName);
                }
                catch (Exception ex)
                {
                    UI.Err("Không thể xuất file (file đang mở trong Excel?).\n" + ex.Message);
                }
            }
        }
    }
}
