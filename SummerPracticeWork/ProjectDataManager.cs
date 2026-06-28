using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace SummerPractice
{
    public class ForeignKeyInfo
    {
        public string ParentTable { get; set; }
        public string ParentColumn { get; set; }
        public string ChildTable { get; set; }
        public string ChildColumn { get; set; }
    }

    public class SaveResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public bool IsForeignKeyViolation { get; set; }
    }

    public class LookupInfo
    {
        public string FkColumn { get; set; }
        public string ParentTable { get; set; }
        public string ParentPkColumn { get; set; }
        public string DisplayColumn { get; set; }
        public DataTable LookupTable { get; set; }
    }

    public class ProjectDataManager : IDisposable
    {
        private DataTable originalTable;
        private OleDbDataAdapter dataAdapter;
        private OleDbCommandBuilder commandBuilder;
        private string currentTableName;
        private readonly string connStr;
        private bool disposed = false;

        private readonly Dictionary<string, List<ForeignKeyInfo>> foreignKeys = new Dictionary<string, List<ForeignKeyInfo>>();
        private static readonly object _globalLock = new object();

        #region Свойства

        public DataTable OriginalTable => originalTable;
        public string CurrentTableName => currentTableName;
        public bool HasUnsavedChanges => originalTable?.GetChanges() != null;
        public bool IsFiltered { get; private set; }
        public bool IsSorted { get; private set; }
        public bool IsSearched { get; private set; }

        #endregion

        public ProjectDataManager(string connectionString)
        {
            connStr = connectionString;
        }

        #region Загрузка таблиц

        public List<string> GetTableNames(bool isAdmin)
        {
            var list = new List<string>();
            try
            {
                var schema = SafeDatabaseHelper.GetSchemaTable(connStr,
                    OleDbSchemaGuid.Tables,
                    new object[] { null, null, null, "TABLE" });

                if (schema != null)
                {
                    foreach (DataRow row in schema.Rows)
                    {
                        string name = row["TABLE_NAME"].ToString();
                        if (!name.StartsWith("MSys") && !name.StartsWith("USys"))
                        {
                            if (!isAdmin && name.Equals("Пользователи", StringComparison.OrdinalIgnoreCase))
                                continue;
                            list.Add(name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetTableNames error: {ex.Message}");
            }

            return list.OrderBy(x => x).ToList();
        }

        public void SelectTable(string tableName)
        {
            lock (_globalLock)
            {
                currentTableName = tableName;
                string sql = $"SELECT * FROM [{tableName}]";

                commandBuilder?.Dispose();
                dataAdapter?.Dispose();

                dataAdapter = new OleDbDataAdapter(sql, connStr);
                commandBuilder = new OleDbCommandBuilder(dataAdapter);

                var dt = new DataTable();
                dataAdapter.Fill(dt);
                originalTable = dt;

                ResetAll();
            }
        }

        #endregion

        #region Сортировка и поиск

        public DataView ApplySort(string columnName, bool ascending)
        {
            if (originalTable == null)
                throw new InvalidOperationException("Таблица не выбрана.");

            if (!originalTable.Columns.Contains(columnName))
                throw new ArgumentException($"Столбец '{columnName}' отсутствует в таблице.");

            DataView dv = new DataView(originalTable);
            string sortDirection = ascending ? "ASC" : "DESC";
            dv.Sort = $"[{columnName}] {sortDirection}";

            IsSorted = true;
            return dv;
        }

        public DataView PerformSearch(string searchText, string columnName)
        {
            if (originalTable == null)
                throw new InvalidOperationException("Таблица не выбрана.");

            var conditions = new List<string>();
            string escapedText = searchText.Replace("'", "''");
            bool isNumber = decimal.TryParse(searchText, out decimal numValue);
            bool isDate = DateTime.TryParse(searchText, out DateTime dateValue);

            IEnumerable<DataColumn> colsToSearch = columnName == "Все столбцы"
                ? originalTable.Columns.Cast<DataColumn>()
                : originalTable.Columns.Cast<DataColumn>().Where(c => c.ColumnName == columnName);

            foreach (DataColumn col in colsToSearch)
            {
                if (col.DataType == typeof(string))
                {
                    conditions.Add($"[{col.ColumnName}] LIKE '*{escapedText}*'");
                }
                else if (IsNumericType(col.DataType) && isNumber)
                {
                    string numStr = numValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    conditions.Add($"[{col.ColumnName}] = {numStr}");
                }
                else if (col.DataType == typeof(DateTime) && isDate)
                {
                    conditions.Add($"[{col.ColumnName}] = #{dateValue:MM/dd/yyyy}#");
                }
            }

            if (conditions.Count == 0)
                return null;

            string filter = string.Join(" OR ", conditions);
            DataView dv = new DataView(originalTable) { RowFilter = filter };

            IsSearched = true;
            return dv;
        }

        public void ResetSort() => IsSorted = false;
        public void ResetSearch() => IsSearched = false;
        public void MarkFiltered() => IsFiltered = true;

        public void ResetAll()
        {
            IsFiltered = false;
            IsSorted = false;
            IsSearched = false;
        }

        #endregion

        #region Внешние ключи

        public List<string> GetPrimaryKeyColumns()
        {
            var pkColumns = new List<string>();
            if (string.IsNullOrEmpty(currentTableName)) return pkColumns;

            try
            {
                var pkSchema = SafeDatabaseHelper.GetSchemaTable(connStr, OleDbSchemaGuid.Primary_Keys, new object[] { null, null, currentTableName });
                if (pkSchema != null)
                {
                    foreach (DataRow row in pkSchema.Rows)
                    {
                        string columnName = row["COLUMN_NAME"]?.ToString();
                        if (!string.IsNullOrEmpty(columnName))
                            pkColumns.Add(columnName);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetPrimaryKeyColumns error: {ex.Message}");
            }

            return pkColumns;
        }

        public object GetNextPrimaryKeyValue(string columnName)
        {
            if (originalTable == null || !originalTable.Columns.Contains(columnName))
                return null;

            DataColumn col = originalTable.Columns[columnName];
            decimal maxValue = 0;
            bool hasValues = false;

            foreach (DataRow row in originalTable.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;
                object value = row[columnName];
                if (value == null || value == DBNull.Value) continue;

                try
                {
                    decimal current = Convert.ToDecimal(value);
                    if (current > maxValue) maxValue = current;
                    hasValues = true;
                }
                catch { }
            }

            decimal nextValue = hasValues ? maxValue + 1 : 1;

            if (col.DataType == typeof(int)) return (int)nextValue;
            if (col.DataType == typeof(long)) return (long)nextValue;
            if (col.DataType == typeof(short)) return (short)nextValue;
            if (col.DataType == typeof(byte)) return (byte)nextValue;
            if (col.DataType == typeof(decimal)) return nextValue;
            if (col.DataType == typeof(double)) return (double)nextValue;
            return nextValue;
        }

        public void LoadForeignKeys()
        {
            foreignKeys.Clear();
            try
            {
                var fkSchema = SafeDatabaseHelper.GetSchemaTable(connStr, OleDbSchemaGuid.Foreign_Keys, new object[] { null, null, null });
                if (fkSchema == null) return;

                foreach (DataRow row in fkSchema.Rows)
                {
                    var fk = new ForeignKeyInfo
                    {
                        ParentTable = row["PK_TABLE_NAME"]?.ToString(),
                        ParentColumn = row["PK_COLUMN_NAME"]?.ToString(),
                        ChildTable = row["FK_TABLE_NAME"]?.ToString(),
                        ChildColumn = row["FK_COLUMN_NAME"]?.ToString()
                    };

                    if (string.IsNullOrEmpty(fk.ParentTable) || string.IsNullOrEmpty(fk.ChildTable))
                        continue;

                    if (!foreignKeys.ContainsKey(fk.ParentTable))
                        foreignKeys[fk.ParentTable] = new List<ForeignKeyInfo>();
                    foreignKeys[fk.ParentTable].Add(fk);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadForeignKeys error: {ex.Message}");
            }
        }

        public List<string> CheckForeignKeyViolations(DataRow row)
        {
            var violations = new List<string>();
            if (string.IsNullOrEmpty(currentTableName) || !foreignKeys.ContainsKey(currentTableName))
                return violations;

            DataTable table = row.Table ?? originalTable;
            if (table == null) return violations;

            foreach (var fk in foreignKeys[currentTableName])
            {
                if (fk.ParentTable != currentTableName || !table.Columns.Contains(fk.ParentColumn))
                    continue;

                object pkValue = null;
                try
                {
                    if (row.HasVersion(DataRowVersion.Original))
                        pkValue = row[fk.ParentColumn, DataRowVersion.Original];
                    else if (row.RowState != DataRowState.Deleted)
                        pkValue = row[fk.ParentColumn];
                }
                catch { continue; }

                if (pkValue == null || pkValue == DBNull.Value)
                    continue;

                string checkSql = $"SELECT COUNT(*) FROM [{fk.ChildTable}] WHERE [{fk.ChildColumn}] = ?";

                try
                {
                    using var conn = new OleDbConnection(connStr);
                    using var cmd = new OleDbCommand(checkSql, conn);
                    cmd.Parameters.AddWithValue("?", pkValue);
                    conn.Open();

                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    if (count > 0)
                    {
                        violations.Add($"• В таблице «{fk.ChildTable}» есть {count} записей, ссылающихся на эту (поле «{fk.ChildColumn}» = {pkValue}).");
                    }
                }
                catch (Exception ex)
                {
                    violations.Add($"• Ошибка проверки таблицы «{fk.ChildTable}»: {ex.Message}");
                }
            }

            return violations;
        }

        #endregion

        #region Lookup

        public List<LookupInfo> GetLookupsForCurrentTable()
        {
            var result = new List<LookupInfo>();
            if (string.IsNullOrEmpty(currentTableName)) return result;

            lock (_globalLock)
            {
                var fkSchema = SafeDatabaseHelper.GetSchemaTable(
                    connStr,
                    OleDbSchemaGuid.Foreign_Keys,
                    new object[] { null, null, null, null, null, currentTableName });

                if (fkSchema == null) return result;

                foreach (DataRow fkRow in fkSchema.Rows)
                {
                    string parentTable = fkRow["PK_TABLE_NAME"]?.ToString();
                    string parentPkCol = fkRow["PK_COLUMN_NAME"]?.ToString();
                    string childFkCol = fkRow["FK_COLUMN_NAME"]?.ToString();

                    if (string.IsNullOrEmpty(parentTable) || string.IsNullOrEmpty(parentPkCol) || string.IsNullOrEmpty(childFkCol))
                        continue;

                    string displayCol = GetDisplayColumn(parentTable, parentPkCol);
                    if (displayCol == null) continue;

                    DataTable lookupTable = LoadLookupTable(parentTable, parentPkCol, displayCol);

                    result.Add(new LookupInfo
                    {
                        FkColumn = childFkCol,
                        ParentTable = parentTable,
                        ParentPkColumn = parentPkCol,
                        DisplayColumn = displayCol,
                        LookupTable = lookupTable
                    });
                }
            }

            return result;
        }

        private string GetDisplayColumn(string tableName, string pkColumnName)
        {
            try
            {
                var colSchema = SafeDatabaseHelper.GetSchemaTable(connStr, OleDbSchemaGuid.Columns, new object[] { null, null, tableName, null });
                if (colSchema == null) return null;

                var columns = colSchema.Rows.Cast<DataRow>()
                    .OrderBy(r => Convert.ToInt32(r["ORDINAL_POSITION"]))
                    .Select(r => r["COLUMN_NAME"]?.ToString())
                    .Where(name => !string.IsNullOrEmpty(name))
                    .ToList();

                return columns.FirstOrDefault(c => !c.Equals(pkColumnName, StringComparison.OrdinalIgnoreCase));
            }
            catch { return null; }
        }

        private DataTable LoadLookupTable(string tableName, string pkCol, string displayCol)
        {
            string sql = $"SELECT [{pkCol}], [{displayCol}] FROM [{tableName}] ORDER BY [{displayCol}]";
            var dt = SafeDatabaseHelper.ExecuteQuery(connStr, sql);

            try
            {
                if (dt != null && dt.Columns.Count >= 2)
                {
                    dt.Columns[0].ColumnName = "ValueMember";
                    dt.Columns[1].ColumnName = "DisplayMember";
                    dt.Columns["ValueMember"].AllowDBNull = true;

                    DataRow emptyRow = dt.NewRow();
                    emptyRow["ValueMember"] = DBNull.Value;
                    emptyRow["DisplayMember"] = "";
                    dt.Rows.InsertAt(emptyRow, 0);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadLookupTable structural error [{tableName}]: {ex.Message}");
                if (dt == null || dt.Columns.Count == 0)
                {
                    dt = new DataTable();
                    dt.Columns.Add("ValueMember", typeof(int));
                    dt.Columns.Add("DisplayMember", typeof(string));
                    dt.Columns["ValueMember"].AllowDBNull = true;
                }
            }

            return dt;
        }

        #endregion

        #region Сохранение изменений

        public SaveResult TrySaveChanges()
        {
            if (dataAdapter == null || commandBuilder == null || originalTable == null)
                return new SaveResult { Success = false, ErrorMessage = "Нет данных для сохранения." };

            lock (_globalLock)
            {
                DataTable changes = originalTable.GetChanges();
                if (changes == null)
                    return new SaveResult { Success = true };

                foreach (DataRow row in changes.Rows)
                {
                    if (row.RowState == DataRowState.Deleted)
                    {
                        var violations = CheckForeignKeyViolations(row);
                        if (violations.Count > 0)
                        {
                            originalTable.RejectChanges();
                            return new SaveResult
                            {
                                Success = false,
                                IsForeignKeyViolation = true,
                                ErrorMessage = "Нельзя удалить запись — существуют связанные данные в других таблицах:\n\n"
                                            + string.Join("\n", violations)
                                            + "\n\nСначала удалите связанные записи из дочерних таблиц."
                            };
                        }
                    }
                    else if (row.RowState == DataRowState.Modified && HasPrimaryKeyChanged(row))
                    {
                        var violations = CheckForeignKeyViolations(row);
                        if (violations.Count > 0)
                        {
                            originalTable.RejectChanges();
                            return new SaveResult
                            {
                                Success = false,
                                IsForeignKeyViolation = true,
                                ErrorMessage = "Изменение первичного ключа невозможно — нарушаются связи с другими таблицами:\n\n"
                                            + string.Join("\n", violations)
                                            + "\n\nИзмените или удалите связанные записи сначала."
                            };
                        }
                    }
                }

                commandBuilder.RefreshSchema();

                try
                {
                    dataAdapter.Update(originalTable);
                    originalTable.AcceptChanges();

                    return new SaveResult { Success = true };
                }
                catch (OleDbException ex)
                {
                    string errorMsg = FormatOleDbError(ex);
                    try { originalTable.RejectChanges(); } catch { }
                    return new SaveResult { Success = false, ErrorMessage = errorMsg };
                }
                catch (Exception ex)
                {
                    try { originalTable.RejectChanges(); } catch { }
                    return new SaveResult { Success = false, ErrorMessage = ex.Message };
                }
            }
        }

        private bool HasPrimaryKeyChanged(DataRow row)
        {
            var pkCols = GetPrimaryKeyColumns();

            foreach (string pkCol in pkCols)
            {
                if (!row.Table.Columns.Contains(pkCol)) continue;
                if (!row.HasVersion(DataRowVersion.Original)) continue;

                object original = row[pkCol, DataRowVersion.Original];
                object current = row[pkCol];

                bool originalIsNull = (original == null || original == DBNull.Value);
                bool currentIsNull = (current == null || current == DBNull.Value);

                if (originalIsNull && currentIsNull) continue;
                if (originalIsNull != currentIsNull) return true;
                if (!object.Equals(original, current)) return true;
            }

            return false;
        }

        private string FormatOleDbError(OleDbException ex)
        {
            string errorMsg = ex.Message;

            if (errorMsg.Contains("relationship") || errorMsg.Contains("связ") || ex.ErrorCode == -2147217900)
            {
                return "Нарушена целостность связей между таблицами.\n\n" +
                       "Возможно, вы пытаетесь удалить запись, на которую ссылаются из другой таблицы, " +
                       "или изменить значение первичного ключа.\n\n" +
                       "Сначала удалите/измените связанные записи в дочерних таблицах.";
            }
            if (errorMsg.Contains("Null") || errorMsg.Contains("null"))
            {
                return "Нельзя оставить обязательное поле пустым.\n\n" +
                       "Проверьте, все ли обязательные поля заполнены.";
            }

            return errorMsg;
        }

        #endregion

        #region Вспомогательные методы и Dispose

        private bool IsNumericType(Type type)
        {
            return type == typeof(int) || type == typeof(long) ||
                   type == typeof(short) || type == typeof(byte) ||
                   type == typeof(decimal) || type == typeof(double) ||
                   type == typeof(float);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    commandBuilder?.Dispose();
                    dataAdapter?.Dispose();
                    originalTable?.Dispose();
                }
                disposed = true;
            }
        }

        #endregion
    }
}