using System;
using System.Data;
using System.Data.OleDb;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace SummerPractice
{
    public static class SafeDatabaseHelper
    {
        private static readonly object _lock = new object();
        private const int MaxRetries = 3;
        private const int RetryDelayMs = 250;

        public static DataTable ExecuteQuery(string connectionString, string sql, OleDbParameter[] parameters = null)
        {
            for (int attempt = 1; attempt <= MaxRetries; attempt++)
            {
                OleDbCommand cmd = null;
                try
                {
                    lock (_lock)
                    {
                        using var conn = new OleDbConnection(connectionString);
                        conn.Open();

                        cmd = new OleDbCommand(sql, conn);
                        cmd.CommandTimeout = 60;

                        if (parameters != null)
                        {
                            // Клонируем параметры, чтобы у каждой попытки были свои чистые объекты в памяти
                            foreach (var p in parameters)
                            {
                                cmd.Parameters.Add(((ICloneable)p).Clone());
                            }
                        }

                        using var adapter = new OleDbDataAdapter(cmd);
                        var dt = new DataTable();
                        adapter.Fill(dt);

                        // Явно очищаем параметры перед успешным выходом
                        cmd.Parameters.Clear();
                        return dt;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Query Error on attempt {attempt}: {ex.Message}");

                    // Безопасная очистка в случае сбоя
                    try { cmd?.Parameters.Clear(); } catch { }

                    if (attempt == MaxRetries)
                        throw; // Пробрасываем дальше, если это не случайный сбой

                    Thread.Sleep(RetryDelayMs * attempt);
                }
            }

            return new DataTable();
        }

        public static int ExecuteNonQuery(string connectionString, string sql, OleDbParameter[] parameters = null)
        {
            for (int attempt = 1; attempt <= MaxRetries; attempt++)
            {
                OleDbCommand cmd = null;
                try
                {
                    lock (_lock)
                    {
                        using var conn = new OleDbConnection(connectionString);
                        conn.Open();

                        cmd = new OleDbCommand(sql, conn);
                        cmd.CommandTimeout = 60;

                        if (parameters != null)
                        {
                            foreach (var p in parameters)
                            {
                                cmd.Parameters.Add(((ICloneable)p).Clone());
                            }
                        }

                        int result = cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                        return result;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"NonQuery Error on attempt {attempt}: {ex.Message}");

                    try { cmd?.Parameters.Clear(); } catch { }

                    if (attempt == MaxRetries)
                        throw;

                    Thread.Sleep(RetryDelayMs * attempt);
                }
            }

            return 0;
        }

        public static DataTable GetSchemaTable(string connectionString, Guid schemaGuid, object[] restrictions)
        {
            for (int attempt = 1; attempt <= MaxRetries; attempt++)
            {
                try
                {
                    lock (_lock)
                    {
                        using var conn = new OleDbConnection(connectionString);
                        conn.Open();

                        // Метод GetOleDbSchemaTable капризен к многопоточности, lock(_lock) здесь обязателен
                        return conn.GetOleDbSchemaTable(schemaGuid, restrictions);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Schema Error on attempt {attempt}: {ex.Message}");

                    if (attempt == MaxRetries)
                        return new DataTable();

                    Thread.Sleep(RetryDelayMs * attempt);
                }
            }

            return new DataTable();
        }
    }
}