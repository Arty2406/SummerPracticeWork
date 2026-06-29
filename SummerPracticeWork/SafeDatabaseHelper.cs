using System;
using System.Data;
using System.Data.OleDb;
using System.Threading;

namespace SummerPractice
{
    /// <summary>
    /// Все обращения к Microsoft.ACE.OLEDB идут через один статический лок.
    /// ACE OLEDB — однопоточная COM-библиотека (STA), параллельные вызовы
    /// приводят к AccessViolationException на уровне нативного кода.
    /// </summary>
    public static class SafeDatabaseHelper
    {
        // Единственный лок для всего приложения — ACE OLEDB не thread-safe
        public static readonly object AceLock = new object();

        private const int MaxRetries = 3;
        private const int RetryDelayMs = 300;

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
                        using var conn = new OleDbConnection(connectionString);
                        conn.Open();

                        using var cmd = new OleDbCommand(sql, conn);
                        cmd.CommandTimeout = 30;

                        if (parameters != null)
                            foreach (var p in parameters)
                                cmd.Parameters.Add(new OleDbParameter(p.ParameterName, p.Value));

                        using var adapter = new OleDbDataAdapter(cmd);
                        var dt = new DataTable();
                        adapter.Fill(dt);
                        return dt;
                    }
                    catch (AccessViolationException ex)
                    {
                        // ACE OLEDB повредил память — больше не повторяем
                        System.Diagnostics.Debug.WriteLine($"[ACE] AccessViolation: {ex.Message}");
                        return new DataTable();
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                        System.Diagnostics.Debug.WriteLine(
                            $"[ACE] Query attempt {attempt} failed: {ex.Message}");
                    }
                }

                if (attempt < MaxRetries)
                    Thread.Sleep(RetryDelayMs * attempt);
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
                        using var conn = new OleDbConnection(connectionString);
                        conn.Open();

                        using var cmd = new OleDbCommand(sql, conn);
                        cmd.CommandTimeout = 30;

                        if (parameters != null)
                            foreach (var p in parameters)
                                cmd.Parameters.Add(new OleDbParameter(p.ParameterName, p.Value));

                        return cmd.ExecuteNonQuery();
                    }
                    catch (AccessViolationException ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ACE] AccessViolation: {ex.Message}");
                        return 0;
                    }
                    catch (Exception ex)
                    {
                        lastEx = ex;
                        System.Diagnostics.Debug.WriteLine(
                            $"[ACE] NonQuery attempt {attempt} failed: {ex.Message}");
                    }
                }

                if (attempt < MaxRetries)
                    Thread.Sleep(RetryDelayMs * attempt);
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
                    using var conn = new OleDbConnection(connectionString);
                    conn.Open();
                    return conn.GetOleDbSchemaTable(schemaGuid, restrictions)
                           ?? new DataTable();
                }
                catch (AccessViolationException ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ACE] Schema AccessViolation: {ex.Message}");
                    return new DataTable();
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