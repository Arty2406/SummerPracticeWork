using System;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

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

        /// <summary>
        /// Хеширует пароль через SHA256 и возвращает HEX строку (64 символа).
        /// </summary>
        private static string HashPassword(string password)
        {
            using var sha256 = SHA256.Create();
            byte[] sourceBytes = Encoding.UTF8.GetBytes(password);
            byte[] hashBytes = sha256.ComputeHash(sourceBytes);

            StringBuilder sb = new StringBuilder();
            foreach (byte b in hashBytes)
                sb.Append(b.ToString("x2")); // HEX строка, 64 символа
            return sb.ToString();
        }

        /// <summary>
        /// Регистрирует пользователя: хеширует пароль и сохраняет в поле "Пароль".
        /// </summary>
        public static void RegisterUser(string login, string pass)
        {
            using var conn = new OleDbConnection(GetConnectionString());
            conn.Open();

            // Проверка на дубликат логина
            string checkSql = "SELECT COUNT(*) FROM Пользователи WHERE Логин = @login";
            using var checkCmd = new OleDbCommand(checkSql, conn);
            checkCmd.Parameters.AddWithValue("@login", login);
            int count = (int)checkCmd.ExecuteScalar();
            if (count > 0)
                throw new Exception("Такой логин уже существует.");

            // Хешируем пароль
            string hash = HashPassword(pass);

            // Вставляем в БД: только Логин и Хеш в поле "Пароль"
            string insertSql = "INSERT INTO Пользователи (Логин, Пароль) VALUES (@login, @hash)";
            using var insertCmd = new OleDbCommand(insertSql, conn);
            insertCmd.Parameters.AddWithValue("@login", login);
            insertCmd.Parameters.AddWithValue("@hash", hash);

            insertCmd.ExecuteNonQuery();
        }

        /// <summary>
        /// Получает пользователя по логину. Возвращает DataTable с колонками: Логин, Пароль
        /// </summary>
        public static DataTable GetUserByLogin(string login)
        {
            var result = new DataTable();
            using var conn = new OleDbConnection(GetConnectionString());

            string sql = "SELECT Логин, Пароль FROM Пользователи WHERE Логин = @login";
            using var cmd = new OleDbCommand(sql, conn);
            cmd.Parameters.AddWithValue("@login", login);

            conn.Open();
            using var adapter = new OleDbDataAdapter(cmd);
            adapter.Fill(result);

            return result;
        }

        /// <summary>
        /// Проверяет пароль: хеширует введённый и сравнивает с сохранённым хешем.
        /// </summary>
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
