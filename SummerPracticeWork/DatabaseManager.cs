using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace SummerPractice
{
    public static class DatabaseManager
    {
        private static string GetConnectionString()
        {
            string dbName = "CourseWork.accdb";
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string dbPath = Path.Combine(basePath, dbName);

            if (!File.Exists(dbPath))
                throw new FileNotFoundException($"Файл БД не найден: {dbPath}. Убедитесь, что CourseWork.accdb лежит в bin\\Debug\\...");

            return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";
        }

        private static string HashPassword(string password)
        {
            using var sha256 = SHA256.Create();
            byte[] sourceBytes = Encoding.UTF8.GetBytes(password);
            byte[] hashBytes = sha256.ComputeHash(sourceBytes);

            StringBuilder sb = new StringBuilder();
            foreach (byte b in hashBytes)
                sb.Append(b.ToString("x2"));
            return sb.ToString();
        }

        public static void RegisterGuest(string login, string pass)
        {
            using var conn = new OleDbConnection(GetConnectionString());
            conn.Open();

            string checkSql = "SELECT COUNT(*) FROM Пользователи WHERE Логин = ?";
            using var checkCmd = new OleDbCommand(checkSql, conn);
            checkCmd.Parameters.AddWithValue("?", login);
            int count = (int)checkCmd.ExecuteScalar();
            if (count > 0)
                throw new Exception("Такой логин уже существует.");

            string hash = HashPassword(pass);
            string insertSql = "INSERT INTO Пользователи (Логин, Пароль, Роль) VALUES (?, ?, ?)";
            using var insertCmd = new OleDbCommand(insertSql, conn);
            insertCmd.Parameters.AddWithValue("?", login);
            insertCmd.Parameters.AddWithValue("?", hash);
            insertCmd.Parameters.AddWithValue("?", "Гость");

            insertCmd.ExecuteNonQuery();
        }

        public static void EnsureAdminCreated()
        {
            using var conn = new OleDbConnection(GetConnectionString());
            conn.Open();

            string adminLogin = "AdminArty";

            string checkSql = "SELECT COUNT(*) FROM Пользователи WHERE Логин = ? AND Роль = ?";
            using var checkCmd = new OleDbCommand(checkSql, conn);
            checkCmd.Parameters.AddWithValue("?", adminLogin);
            checkCmd.Parameters.AddWithValue("?", "Администратор");
            int count = (int)checkCmd.ExecuteScalar();

            if (count == 0)
            {
                string hash = HashPassword("24062007");
                string insertSql = "INSERT INTO Пользователи (Логин, Пароль, Роль) VALUES (?, ?, ?)";
                using var insertCmd = new OleDbCommand(insertSql, conn);
                insertCmd.Parameters.AddWithValue("?", adminLogin);
                insertCmd.Parameters.AddWithValue("?", hash);
                insertCmd.Parameters.AddWithValue("?", "Администратор");
                insertCmd.ExecuteNonQuery();
            }
        }

        public static DataTable GetUserByLogin(string login)
        {
            var result = new DataTable();
            using var conn = new OleDbConnection(GetConnectionString());

            string sql = "SELECT Логин, Пароль, Роль FROM Пользователи WHERE Логин = ?";
            using var cmd = new OleDbCommand(sql, conn);
            cmd.Parameters.AddWithValue("?", login);

            conn.Open();
            using var adapter = new OleDbDataAdapter(cmd);
            adapter.Fill(result);
            return result;
        }

        public static bool VerifyPassword(string inputPassword, string storedHash)
        {
            try
            {
                string hash = HashPassword(inputPassword);
                return hash == storedHash;
            }
            catch
            {
                return false;
            }
        }
    }
}
