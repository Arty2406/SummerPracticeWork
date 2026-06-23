using System.Data;
using System.Data.OleDb;

namespace SummerPractice.DataAccess
{
    public class DataAccess
    {
        private readonly string ConnectionString;

        public DataAccess(string connectionString)
        {
            ConnectionString = connectionString;
        }

        public DataTable GetData(string sql, OleDbParameter[] parameters = null)
        {
            using var conn = new OleDbConnection(ConnectionString);
            using var cmd = new OleDbCommand(sql, conn);

            if (parameters != null)
                cmd.Parameters.AddRange(parameters);

            var adapter = new OleDbDataAdapter(cmd);
            var dt = new DataTable();
            adapter.Fill(dt);
            return dt;
        }

        public int ExecuteNonQuery(string sql, OleDbParameter[] parameters = null)
        {
            using var conn = new OleDbConnection(ConnectionString);
            using var cmd = new OleDbCommand(sql, conn);

            if (parameters != null)
                cmd.Parameters.AddRange(parameters);

            conn.Open();
            return cmd.ExecuteNonQuery();
        }
    }
}
