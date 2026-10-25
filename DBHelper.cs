using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace QuanLyBanHang
{
    /// <summary>Lớp dùng chung để làm việc với SQL Server (ADO.NET). Mọi câu lệnh đều dùng tham số (chống SQL Injection).</summary>
    public static class DBHelper
    {
        public static string ConnStr
        {
            get { return ConfigurationManager.ConnectionStrings["HangHoa"].ConnectionString; }
        }

        /// <summary>Tạo tham số; null được đổi thành DBNull.</summary>
        public static SqlParameter P(string name, object value)
        {
            return new SqlParameter(name, value ?? DBNull.Value);
        }

        /// <summary>SELECT -> DataTable (mô hình ngắt kết nối: SqlDataAdapter.Fill).</summary>
        public static DataTable Query(string sql, params SqlParameter[] ps)
        {
            var dt = new DataTable();
            using (var cn = new SqlConnection(ConnStr))
            using (var da = new SqlDataAdapter(sql, cn))
            {
                if (ps != null) da.SelectCommand.Parameters.AddRange(ps);
                da.Fill(dt);
            }
            return dt;
        }

        /// <summary>INSERT / UPDATE / DELETE -> số dòng bị ảnh hưởng.</summary>
        public static int Exec(string sql, params SqlParameter[] ps)
        {
            using (var cn = new SqlConnection(ConnStr))
            using (var cmd = new SqlCommand(sql, cn))
            {
                if (ps != null) cmd.Parameters.AddRange(ps);
                cn.Open();
                return cmd.ExecuteNonQuery();
            }
        }

        /// <summary>Trả về giá trị đầu tiên (COUNT, MAX, ...).</summary>
        public static object Scalar(string sql, params SqlParameter[] ps)
        {
            using (var cn = new SqlConnection(ConnStr))
            using (var cmd = new SqlCommand(sql, cn))
            {
                if (ps != null) cmd.Parameters.AddRange(ps);
                cn.Open();
                return cmd.ExecuteScalar();
            }
        }

        // ---- Các hàm dùng trong Transaction (dùng chung 1 connection) ----
        public static int ExecTx(SqlConnection cn, SqlTransaction tr, string sql, params SqlParameter[] ps)
        {
            using (var cmd = new SqlCommand(sql, cn, tr))
            {
                if (ps != null) cmd.Parameters.AddRange(ps);
                return cmd.ExecuteNonQuery();
            }
        }

        public static object ScalarTx(SqlConnection cn, SqlTransaction tr, string sql, params SqlParameter[] ps)
        {
            using (var cmd = new SqlCommand(sql, cn, tr))
            {
                if (ps != null) cmd.Parameters.AddRange(ps);
                return cmd.ExecuteScalar();
            }
        }
    }

    /// <summary>Lưu thông tin người đăng nhập hiện tại để phân quyền ở các form.</summary>
    public static class Session
    {
        public static string UserName;
        public static string Quyen;          // "Admin" hoặc "NhanVien"
        public static string MaNhanVien;     // null nếu là admin

        public static bool IsAdmin
        {
            get { return Quyen == "Admin"; }
        }

        public static void Clear()
        {
            UserName = null; Quyen = null; MaNhanVien = null;
        }
    }
}
