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
            const string dbName = "CourseWork.accdb";
            string basePath = AppDomain.CurrentDomain.BaseDirectory;

            // Поиск файла в нескольких местах (отладочные и релизные папки)
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
                throw new FileNotFoundException($"Файл базы данных не найден. Пути поиска:\n{string.Join("\n", paths)}");

            return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Persist Security Info=False;OLE DB Services = -4;";
        }

        private static string HashPassword(string password)
        {
            if (password == null) return string.Empty;

            using (var sha256 = SHA256.Create())
            {
                byte[] sourceBytes = Encoding.UTF8.GetBytes(password);
                byte[] hashBytes = sha256.ComputeHash(sourceBytes);

                var sb = new StringBuilder(hashBytes.Length * 2);
                foreach (byte b in hashBytes)
                    sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        public static void RegisterGuest(string login, string pass)
        {
            if (string.IsNullOrWhiteSpace(login) || string.IsNullOrWhiteSpace(pass))
                throw new ArgumentException("Логин и пароль не могут быть пустыми или состоять из пробелов.");

            string connStr = GetConnectionString();
            string trimmedLogin = login.Trim();

            // 1. Проверяем, существует ли уже такой пользователь
            string checkSql = "SELECT COUNT(*) FROM [Пользователи] WHERE [Логин] = ?";

            using (var resultTable = SafeDatabaseHelper.ExecuteQuery(connStr, checkSql, new[] { new OleDbParameter("?", trimmedLogin) }))
            {
                if (resultTable != null && resultTable.Rows.Count > 0)
                {
                    int count = Convert.ToInt32(resultTable.Rows[0][0]);
                    if (count > 0)
                        throw new InvalidOperationException("Пользователь с таким логином уже зарегистрирован.");
                }
            }

            // 2. Вставляем нового пользователя (строго соблюдая порядок позиционных параметров для OLE DB)
            string hash = HashPassword(pass);
            string insertSql = "INSERT INTO [Пользователи] ([Логин], [Пароль], [Роль]) VALUES (?, ?, ?)";

            SafeDatabaseHelper.ExecuteNonQuery(connStr, insertSql,
                new[] {
                    new OleDbParameter("?", trimmedLogin),
                    new OleDbParameter("?", hash),
                    new OleDbParameter("?", "Гость")
                });
        }

        public static void EnsureAdminCreated()
        {
            try
            {
                string connStr = GetConnectionString();
                const string adminLogin = "AdminArty";

                // Проверяем существование администратора по логину
                string checkSql = "SELECT [Роль] FROM [Пользователи] WHERE [Логин] = ?";

                using (var result = SafeDatabaseHelper.ExecuteQuery(connStr, checkSql, new[] { new OleDbParameter("?", adminLogin) }))
                {
                    if (result == null || result.Rows.Count == 0)
                    {
                        // Администратора нет — создаем новую запись
                        string hash = HashPassword("24062007");
                        string insertSql = "INSERT INTO [Пользователи] ([Логин], [Пароль], [Роль]) VALUES (?, ?, ?)";

                        SafeDatabaseHelper.ExecuteNonQuery(connStr, insertSql,
                            new[] {
                                new OleDbParameter("?", adminLogin),
                                new OleDbParameter("?", hash),
                                new OleDbParameter("?", "Администратор")
                            });
                    }
                    else
                    {
                        // Пользователь есть. Проверяем, назначена ли ему роль Администратора
                        string currentRole = result.Rows[0]["Роль"]?.ToString();
                        if (currentRole != "Администратор")
                        {
                            string updateSql = "UPDATE [Пользователи] SET [Роль] = ? WHERE [Логин] = ?";

                            SafeDatabaseHelper.ExecuteNonQuery(connStr, updateSql,
                                new[] {
                                    new OleDbParameter("?", "Администратор"),
                                    new OleDbParameter("?", adminLogin)
                                });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Логируем ошибку, чтобы приложение не упало при первом старте из-за проблем с БД
                System.Diagnostics.Debug.WriteLine($"[DatabaseManager.EnsureAdminCreated] Ошибка инициализации администратора: {ex.Message}");
            }
        }

        public static DataTable GetUserByLogin(string login)
        {
            if (string.IsNullOrWhiteSpace(login))
                return new DataTable();

            try
            {
                string connStr = GetConnectionString();
                string sql = "SELECT [Логин], [Пароль], [Роль] FROM [Пользователи] WHERE [Логин] = ?";

                return SafeDatabaseHelper.ExecuteQuery(connStr, sql, new[] { new OleDbParameter("?", login.Trim()) });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DatabaseManager.GetUserByLogin] Ошибка поиска пользователя '{login}': {ex.Message}");
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
                // Используем StringComparison.Ordinal для безопасного побайтового сравнения хэш-строк
                return string.Equals(hash, storedHash, StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }
    }
}