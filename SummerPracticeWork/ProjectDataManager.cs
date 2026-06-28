using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace SummerPractice
{
    public class ForeignKeyInfo
    {
        public string ParentTable { get; set; }   // родительская таблица (на которую ссылаются)
        public string ParentColumn { get; set; }  // поле PK в родительской таблице
        public string ChildTable { get; set; }    // дочерняя таблица (содержит FK)
        public string ChildColumn { get; set; }   // поле FK в дочерней таблице
    }

    public class SaveResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public bool IsForeignKeyViolation { get; set; }
    }

    /// <summary>
    /// Описывает одну FK-связь для текущей таблицы как дочерней,
    /// дополненную данными справочника (родительской таблицы).
    /// </summary>
    public class LookupInfo
    {
        /// <summary>Столбец FK в текущей (дочерней) таблице.</summary>
        public string FkColumn { get; set; }

        /// <summary>Родительская (справочная) таблица.</summary>
        public string ParentTable { get; set; }

        /// <summary>PK-столбец родительской таблицы (хранится в FK).</summary>
        public string ParentPkColumn { get; set; }

        /// <summary>Столбец отображения из родительской таблицы (второй столбец / первый не-PK).</summary>
        public string DisplayColumn { get; set; }

        /// <summary>Все строки родительской таблицы: ValueMember → DisplayMember.</summary>
        public DataTable LookupTable { get; set; }
    }

    public class ProjectDataManager
    {
        private DataTable originalTable;
        private OleDbDataAdapter dataAdapter;
        private OleDbCommandBuilder commandBuilder;
        private string currentTableName;
        private readonly string connStr;

        // FK-карта: ParentTable -> список FK-описаний (используется для проверки удаления/изменения)
        private Dictionary<string, List<ForeignKeyInfo>> foreignKeys = new Dictionary<string, List<ForeignKeyInfo>>();

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
            using var conn = new OleDbConnection(connStr);
            conn.Open();

            string[] restrictions = new string[4] { null, null, null, "TABLE" };
            var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, restrictions);

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

            return list.OrderBy(x => x).ToList();
        }

        public void SelectTable(string tableName)
        {
            currentTableName = tableName;
            string sql = $"SELECT * FROM [{tableName}]";

            dataAdapter = new OleDbDataAdapter(sql, connStr);
            commandBuilder = new OleDbCommandBuilder(dataAdapter);

            var dt = new DataTable();
            dataAdapter.Fill(dt);
            originalTable = dt;

            IsFiltered = false;
            IsSorted = false;
            IsSearched = false;
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

            IEnumerable<DataColumn> colsToSearch;
            if (columnName == "Все столбцы")
            {
                colsToSearch = originalTable.Columns.Cast<DataColumn>();
            }
            else
            {
                colsToSearch = originalTable.Columns.Cast<DataColumn>()
                    .Where(c => c.ColumnName == columnName);
            }

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

        #region Внешние ключи (проверка целостности при удалении/изменении)

        public List<string> GetPrimaryKeyColumns()
        {
            var pkColumns = new List<string>();
            if (string.IsNullOrEmpty(currentTableName)) return pkColumns;

            using var conn = new OleDbConnection(connStr);
            conn.Open();

            var pkSchema = conn.GetOleDbSchemaTable(
                OleDbSchemaGuid.Primary_Keys,
                new object[] { null, null, currentTableName });

            if (pkSchema != null)
            {
                foreach (DataRow row in pkSchema.Rows)
                {
                    string columnName = row["COLUMN_NAME"]?.ToString();
                    if (!string.IsNullOrEmpty(columnName))
                        pkColumns.Add(columnName);
                }
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

            using var conn = new OleDbConnection(connStr);
            conn.Open();

            var fkSchema = conn.GetOleDbSchemaTable(
                OleDbSchemaGuid.Foreign_Keys,
                new object[] { null, null, null });

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

        public List<string> CheckForeignKeyViolations(DataRow row)
        {
            var violations = new List<string>();

            if (string.IsNullOrEmpty(currentTableName) || !foreignKeys.ContainsKey(currentTableName))
                return violations;

            foreach (var fk in foreignKeys[currentTableName])
            {
                if (fk.ParentTable != currentTableName)
                    continue;

                if (!row.Table.Columns.Contains(fk.ParentColumn))
                    continue;

                object pkValue;
                try
                {
                    pkValue = row[fk.ParentColumn, DataRowVersion.Original];
                }
                catch
                {
                    try { pkValue = row[fk.ParentColumn]; }
                    catch { continue; }
                }

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
                        violations.Add(
                            $"• В таблице «{fk.ChildTable}» есть {count} записей, " +
                            $"ссылающихся на эту (поле «{fk.ChildColumn}» = {pkValue}).");
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

        #region Lookup (ComboBox-столбцы для FK в текущей таблице)

        public List<LookupInfo> GetLookupsForCurrentTable()
        {
            var result = new List<LookupInfo>();

            if (string.IsNullOrEmpty(currentTableName))
                return result;

            using var conn = new OleDbConnection(connStr);
            conn.Open();

            // Получаем все FK схемы, где текущая таблица — дочерняя (FK_TABLE_NAME)
            var fkSchema = conn.GetOleDbSchemaTable(
                OleDbSchemaGuid.Foreign_Keys,
                new object[] { null, null, null, null, null, currentTableName });

            if (fkSchema == null || fkSchema.Rows.Count == 0)
                return result;

            foreach (DataRow fkRow in fkSchema.Rows)
            {
                string parentTable = fkRow["PK_TABLE_NAME"]?.ToString();
                string parentPkCol = fkRow["PK_COLUMN_NAME"]?.ToString();
                string childFkCol = fkRow["FK_COLUMN_NAME"]?.ToString();

                if (string.IsNullOrEmpty(parentTable) ||
                    string.IsNullOrEmpty(parentPkCol) ||
                    string.IsNullOrEmpty(childFkCol))
                    continue;

                // Определяем столбец отображения: первый не-PK столбец родительской таблицы
                string displayCol = GetDisplayColumn(conn, parentTable, parentPkCol);
                if (displayCol == null)
                    continue; // нет подходящего столбца — пропускаем

                // Загружаем данные справочника
                DataTable lookupTable = LoadLookupTable(conn, parentTable, parentPkCol, displayCol);

                result.Add(new LookupInfo
                {
                    FkColumn = childFkCol,
                    ParentTable = parentTable,
                    ParentPkColumn = parentPkCol,
                    DisplayColumn = displayCol,
                    LookupTable = lookupTable
                });
            }

            return result;
        }

        private string GetDisplayColumn(OleDbConnection conn, string tableName, string pkColumnName)
        {
            try
            {
                // Читаем схему столбцов родительской таблицы
                var colSchema = conn.GetOleDbSchemaTable(
                    OleDbSchemaGuid.Columns,
                    new object[] { null, null, tableName, null });

                if (colSchema == null) return null;

                // Собираем все столбцы в порядке ORDINAL_POSITION
                var columns = colSchema.Rows.Cast<DataRow>()
                    .OrderBy(r => Convert.ToInt32(r["ORDINAL_POSITION"]))
                    .Select(r => r["COLUMN_NAME"]?.ToString())
                    .Where(name => !string.IsNullOrEmpty(name))
                    .ToList();

                // Первый не-PK столбец — кандидат на отображение
                string displayCol = columns.FirstOrDefault(
                    c => !c.Equals(pkColumnName, StringComparison.OrdinalIgnoreCase));

                return displayCol;
            }
            catch
            {
                return null;
            }
        }

        private DataTable LoadLookupTable(OleDbConnection conn, string tableName, string pkCol, string displayCol)
        {
            // Сначала читаем реальные данные через адаптер, чтобы получить правильный тип PK
            var dt = new DataTable();

            try
            {
                string sql = $"SELECT [{pkCol}], [{displayCol}] FROM [{tableName}] ORDER BY [{displayCol}]";
                using var cmd = new OleDbCommand(sql, conn);
                using var da = new OleDbDataAdapter(cmd);
                da.Fill(dt); // Fill сохраняет реальные типы (Int32, String и т.д.)

                // Переименовываем столбцы
                dt.Columns[0].ColumnName = "ValueMember";
                dt.Columns[1].ColumnName = "DisplayMember";

                // Разрешаем NULL в ValueMember (для пустой строки)
                dt.Columns["ValueMember"].AllowDBNull = true;

                // Добавляем пустую строку первой (для необязательных FK)
                DataRow emptyRow = dt.NewRow();
                emptyRow["ValueMember"] = DBNull.Value;
                emptyRow["DisplayMember"] = "";
                dt.Rows.InsertAt(emptyRow, 0);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadLookupTable error [{tableName}]: {ex.Message}");

                // Fallback: возвращаем хотя бы пустую таблицу с правильными колонками
                if (dt.Columns.Count == 0)
                {
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

            try
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
                    else if (row.RowState == DataRowState.Modified)
                    {
                        // Проверяем связи ТОЛЬКО если изменился первичный ключ.
                        // Изменение FK-поля (выбор другого значения из справочника) — это нормально,
                        // блокировать его не нужно.
                        if (HasPrimaryKeyChanged(row))
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
                }

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

        // Проверяет, изменился ли первичный ключ строки.
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
            else if (errorMsg.Contains("Null") || errorMsg.Contains("null"))
            {
                return "Нельзя оставить обязательное поле пустым.\n\n" +
                       "Проверьте, все ли обязательные поля заполнены.";
            }

            return errorMsg;
        }


        #endregion

        #region Вспомогательные методы

        private bool IsNumericType(Type type)
        {
            return type == typeof(int) || type == typeof(long) ||
                   type == typeof(short) || type == typeof(byte) ||
                   type == typeof(decimal) || type == typeof(double) ||
                   type == typeof(float);
        }

        #endregion
    }
}