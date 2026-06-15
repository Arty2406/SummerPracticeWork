using System.Data;
using System.Data.OleDb;

namespace SummerPractice.DataAccess
{
    public class DataAccess
    {
        private readonly string _connectionString;

        public DataAccess(string connectionString)
        {
            _connectionString = connectionString;
        }

        // Универсальный метод для получения данных (SELECT)
        // Аргумент: использование этого метода вместо прямого new OleDbDataAdapter везде снижает дублирование кода.
        public DataTable GetData(string sql, OleDbParameter[] parameters = null)
        {
            using var conn = new OleDbConnection(_connectionString);
            using var cmd = new OleDbCommand(sql, conn);

            if (parameters != null)
                cmd.Parameters.AddRange(parameters);

            var adapter = new OleDbDataAdapter(cmd);
            var dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }

        // Универсальный метод для изменения данных (INSERT, UPDATE, DELETE)
        public int ExecuteNonQuery(string sql, OleDbParameter[] parameters = null)
        {
            using var conn = new OleDbConnection(_connectionString);
            using var cmd = new OleDbCommand(sql, conn);

            if (parameters != null)
                cmd.Parameters.AddRange(parameters);

            conn.Open();
            return cmd.ExecuteNonQuery();
        }
    }
}
