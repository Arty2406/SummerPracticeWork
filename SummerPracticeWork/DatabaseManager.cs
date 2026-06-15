using System;
using System.Data;
using System.Data.OleDb;
using System.Security.Cryptography;
using System.Text;
using SummerPractice.DataAccess;

namespace SummerPractice
{
    // Аргумент выбора: Вынос логики безопасности и инициализации БД в отдельный класс 
    // предотвращает случайное сохранение паролей в открытом виде в коде форм.
    public static class DatabaseManager
    {
        private static readonly DataAccess.DataAccess dataAccess;

        static DatabaseManager()
        {
            // Динамический путь к БД (как мы обсуждали ранее)
            string dbName = "CourseWork.accdb";
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string dbPath = System.IO.Path.Combine(basePath, dbName);

            // Проверка наличия файла перед попыткой подключения
            if (!System.IO.File.Exists(dbPath))
                throw new FileNotFoundException($"Файл базы данных не найден по пути: {dbPath}. Проверьте настройки копирования файла в проект.");

            string connStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";

            // ПРОВЕРКА ДРАЙВЕРА: Попытка открыть соединение сразу при старте класса.
            // Если драйвера нет, вылетит исключение, которое мы поймаем в форме входа.
            try
            {
                using var testConn = new OleDbConnection(connStr);
                testConn.Open();
            }
            catch (Exception ex) when (ex.Message.Contains("Provider"))
            {
                throw new Exception("Не установлен драйвер Microsoft ACE OLEDB. Пожалуйста, установите Access Database Engine.", ex);
            }

            dataAccess = new DataAccess.DataAccess(connStr);
        }

        public static string ConnectionString => dataAccess.GetType().GetField("_connectionString", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.GetValue(dataAccess).ToString();

        /// <summary>
        /// Хеширование пароля с солью (Salt) используя SHA256.
        /// Аргумент: Простой MD5 или хранение пароля в открытом виде ненадежны. 
        /// SHA256 является стандартом для таких задач в .NET. Соль делает невозможным подбор по радужным таблицам.
        /// </summary>
        public static (string Hash, string Salt) HashPassword(string password)
        {
            byte[] saltBytes = new byte[16];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(saltBytes);
            }

            byte[] passwordBytes = Encoding.UTF8.GetBytes(password);
            byte[] combinedBytes = new byte[saltBytes.Length + passwordBytes.Length];

            Buffer.BlockCopy(saltBytes, 0, combinedBytes, 0, saltBytes.Length);
            Buffer.BlockCopy(passwordBytes, 0, combinedBytes, saltBytes.Length, passwordBytes.Length);

            using (var sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(combinedBytes);
                return (Convert.ToBase64String(hashBytes), Convert.ToBase64String(saltBytes));
            }
        }

        /// <summary>
        /// Проверка пароля: берет сохраненную соль, добавляет введенный пароль, хеширует и сравнивает.
        /// </summary>
        public static bool VerifyPassword(string inputPassword, string storedHash, string storedSalt)
        {
            byte[] saltBytes = Convert.FromBase64String(storedSalt);
            byte[] passwordBytes = Encoding.UTF8.GetBytes(inputPassword);
            byte[] combinedBytes = new byte[saltBytes.Length + passwordBytes.Length];

            Buffer.BlockCopy(saltBytes, 0, combinedBytes, 0, saltBytes.Length);
            Buffer.BlockCopy(passwordBytes, 0, combinedBytes, saltBytes.Length, passwordBytes.Length);

            using (var sha256 = SHA256.Create())
            {
                byte[] computedHash = sha256.ComputeHash(combinedBytes);
                string computedHashString = Convert.ToBase64String(computedHash);
                return computedHashString == storedHash;
            }
        }

        public static DataTable GetTableNames()
        {
            // Используем системную таблицу MSysObjects для надежности, если GetOleDbSchemaTable падает
            // Аргумент: Это более универсальный способ получения имен таблиц в Access, работающий даже в старых версиях драйверов.
            string sql = "SELECT Name FROM MSysObjects WHERE Type = 1 AND Flags = 0 ORDER BY Name";
            return dataAccess.GetData(sql);
        }

        public static DataTable GetUserByLogin(string login)
        {
            string sql = "SELECT * FROM Users WHERE Login = ?";
            var param = new OleDbParameter("?", login);
            return dataAccess.GetData(sql, new[] { param });
        }

        public static void RegisterUser(string login, string password)
        {
            var (hash, salt) = HashPassword(password);
            string sql = "INSERT INTO Users (Login, PasswordHash, Salt) VALUES (?, ?, ?)";
            var paramsArr = new[]
            {
                new OleDbParameter("?", login),
                new OleDbParameter("?", hash),
                new OleDbParameter("?", salt)
            };
            dataAccess.ExecuteNonQuery(sql, paramsArr);
        }

        // Метод для получения структуры таблицы (имен колонок) для динамической формы
        public static DataTable GetTableSchema(string tableName)
        {
            // SELECT TOP 0 возвращает только структуру (колонки), не грузя данные. Это быстро.
            // Аргумент: Лучше чем SELECT *, так как не тратит время на чтение миллионов строк.
            string sql = $"SELECT TOP 0 * FROM [{tableName}]";
            return dataAccess.GetData(sql);
        }
    }
}
