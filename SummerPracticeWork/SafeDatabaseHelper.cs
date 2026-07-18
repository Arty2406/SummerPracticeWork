using System;
using System.Data;
using System.Data.OleDb;
using System.Threading;

namespace SummerPractice
{
    public static class SafeDatabaseHelper
    {
        // Глобальный лок для синхронизации доступа к однопоточному движку MS Access (Jet/ACE)
        public static readonly object AceLock = new object();

        private const int MaxRetries = 3;
        private const int RetryDelayMs = 150; // Уменьшим задержку, чтобы интерфейс не зависал надолго

        public static DataTable ExecuteQuery(
            string connectionString, string sql, OleDbParameter[] parameters = null)
        {
            Exception lastEx = null;

            for (int attempt = 1; attempt <= MaxRetries; attempt++)
            {
                lock (AceLock)
                {
                    try
                    {
                        using (var conn = new OleDbConnection(connectionString))
                        {
                            conn.Open();

                            using (var cmd = new OleDbCommand(sql, conn))
                            {
                                cmd.CommandTimeout = 15; // 15 секунд более чем достаточно

                                // Решаем проблему повторного использования параметров:
                                // Создаем новые чистые параметры для каждой попытки
                                if (parameters != null)
                                {
                                    foreach (var p in parameters)
                                    {
                                        // Копируем только имя, тип и значение, чтобы избежать привязки к старым командам
                                        var newParam = new OleDbParameter(p.ParameterName, p.OleDbType)
                                        {
                                            Value = p.Value ?? DBNull.Value
                                        };
                                        cmd.Parameters.Add(newParam);
                                    }
                                }

                                using (var adapter = new OleDbDataAdapter(cmd))
                                {
                                    var dt = new DataTable();
                                    adapter.Fill(dt);
                                    return dt;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                        System.Diagnostics.Debug.WriteLine($"[ACE] Query attempt {attempt} failed: {ex.Message}");
                    }
                }

                // Задержка между попытками вне блока lock! 
                // Иначе мы держим блокировку во время сна, не давая другим потокам закрыть соединения.
                if (attempt < MaxRetries)
                {
                    Thread.Sleep(RetryDelayMs * attempt);
                }
            }

            throw lastEx ?? new Exception("ExecuteQuery failed after retries.");
        }

        public static int ExecuteNonQuery(
            string connectionString, string sql, OleDbParameter[] parameters = null)
        {
            Exception lastEx = null;

            for (int attempt = 1; attempt <= MaxRetries; attempt++)
            {
                lock (AceLock)
                {
                    try
                    {
                        using (var conn = new OleDbConnection(connectionString))
                        {
                            conn.Open();

                            using (var cmd = new OleDbCommand(sql, conn))
                            {
                                cmd.CommandTimeout = 15;

                                // Безопасное копирование параметров для каждой новой попытки
                                if (parameters != null)
                                {
                                    foreach (var p in parameters)
                                    {
                                        var newParam = new OleDbParameter(p.ParameterName, p.OleDbType)
                                        {
                                            Value = p.Value ?? DBNull.Value
                                        };
                                        cmd.Parameters.Add(newParam);
                                    }
                                }

                                return cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                        System.Diagnostics.Debug.WriteLine($"[ACE] NonQuery attempt {attempt} failed: {ex.Message}");
                    }
                }

                if (attempt < MaxRetries)
                {
                    Thread.Sleep(RetryDelayMs * attempt);
                }
            }

            throw lastEx ?? new Exception("ExecuteNonQuery failed after retries.");
        }

        public static DataTable GetSchemaTable(
            string connectionString, Guid schemaGuid, object[] restrictions)
        {
            lock (AceLock)
            {
                try
                {
                    using (var conn = new OleDbConnection(connectionString))
                    {
                        conn.Open();
                        return conn.GetOleDbSchemaTable(schemaGuid, restrictions) ?? new DataTable();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ACE] Schema error: {ex.Message}");
                    return new DataTable();
                }
            }
        }
    }
}