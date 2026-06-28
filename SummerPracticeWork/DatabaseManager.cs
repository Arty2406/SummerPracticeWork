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

            // Поиск файла в нескольких местах
            string[] paths = {
                Path.Combine(basePath, dbName),
                Path.Combine(basePath, "..", "..", "..", dbName),
                Path.Combine(basePath, "..", "..", dbName),
                Path.Combine(basePath, "..", dbName),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, dbName)
            };

            string dbPath = null;
            foreach (var path in paths)
            {
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    dbPath = fullPath;
                    break;
                }
            }

            if (string.IsNullOrEmpty(dbPath))
                throw new FileNotFoundException($"Файл БД не найден. Искали в:\n{string.Join("\n", paths)}");

            return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Persist Security Info=False;OLE DB Services = -4;";
        }

        private static string HashPassword(string password)
        {
            using var sha256 = SHA256.Create();
            byte[] sourceBytes = Encoding.UTF8.GetBytes(password);
            byte[] hashBytes = sha256.ComputeHash(sourceBytes);

            var sb = new StringBuilder(hashBytes.Length * 2);
            foreach (byte b in hashBytes)
                sb.Append(b.ToString("x2"));
            return sb.ToString();
        }

        public static void RegisterGuest(string login, string pass)
        {
            if (string.IsNullOrWhiteSpace(login) || string.IsNullOrWhiteSpace(pass))
                throw new ArgumentException("Логин и пароль не могут быть пустыми.");

            string connStr = GetConnectionString();

            // Проверка существования пользователя
            string checkSql = "SELECT COUNT(*) FROM Пользователи WHERE Логин = ?";
            var result = SafeDatabaseHelper.ExecuteQuery(connStr, checkSql,
                new[] { new OleDbParameter("?", login) });

            if (result.Rows.Count > 0 && Convert.ToInt32(result.Rows[0][0]) > 0)
                throw new Exception("Такой логин уже существует.");

            // Вставка нового пользователя
            string hash = HashPassword(pass);
            string insertSql = "INSERT INTO Пользователи (Логин, Пароль, Роль) VALUES (?, ?, ?)";

            SafeDatabaseHelper.ExecuteNonQuery(connStr, insertSql,
                new[] {
                    new OleDbParameter("?", login),
                    new OleDbParameter("?", hash),
                    new OleDbParameter("?", "Гость")
                });
        }

        public static void EnsureAdminCreated()
        {
            try
            {
                string connStr = GetConnectionString();
                string adminLogin = "AdminArty";

                // Проверяем существование пользователя
                string checkSql = "SELECT Роль FROM Пользователи WHERE Логин = ?";
                var result = SafeDatabaseHelper.ExecuteQuery(connStr, checkSql,
                    new[] { new OleDbParameter("?", adminLogin) });

                if (result.Rows.Count == 0)
                {
                    // Создаем администратора
                    string hash = HashPassword("24062007");
                    string insertSql = "INSERT INTO Пользователи (Логин, Пароль, Роль) VALUES (?, ?, ?)";
                    SafeDatabaseHelper.ExecuteNonQuery(connStr, insertSql,
                        new[] {
                            new OleDbParameter("?", adminLogin),
                            new OleDbParameter("?", hash),
                            new OleDbParameter("?", "Администратор")
                        });
                }
                else
                {
                    string currentRole = result.Rows[0]["Роль"]?.ToString();
                    if (currentRole != "Администратор")
                    {
                        string updateSql = "UPDATE Пользователи SET Роль = ? WHERE Логин = ?";
                        SafeDatabaseHelper.ExecuteNonQuery(connStr, updateSql,
                            new[] {
                                new OleDbParameter("?", "Администратор"),
                                new OleDbParameter("?", adminLogin)
                            });
                    }
                }
            }
            catch (AccessViolationException ex)
            {
                System.Diagnostics.Debug.WriteLine($"AccessViolation in EnsureAdminCreated: {ex.Message}");
                // Не выбрасываем исключение, чтобы приложение не падало
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"EnsureAdminCreated error: {ex.Message}");
            }
        }

        public static DataTable GetUserByLogin(string login)
        {
            try
            {
                string connStr = GetConnectionString();
                string sql = "SELECT Логин, Пароль, Роль FROM Пользователи WHERE Логин = ?";
                return SafeDatabaseHelper.ExecuteQuery(connStr, sql,
                    new[] { new OleDbParameter("?", login) });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetUserByLogin error: {ex.Message}");
                return new DataTable();
            }
        }

        public static bool VerifyPassword(string inputPassword, string storedHash)
        {
            if (string.IsNullOrEmpty(inputPassword) || string.IsNullOrEmpty(storedHash))
                return false;

            try
            {
                string hash = HashPassword(inputPassword);
                return string.Equals(hash, storedHash, StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }
    }
}