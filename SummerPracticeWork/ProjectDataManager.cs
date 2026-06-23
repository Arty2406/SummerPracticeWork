using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace SummerPractice
{
    public class ForeignKeyInfo
    {
        public string ParentTable { get; set; }   // Родительская таблица (на которую ссылаются)
        public string ParentColumn { get; set; }  // Поле PK в родительской таблице
        public string ChildTable { get; set; }    // Дочерняя таблица (содержит FK)
        public string ChildColumn { get; set; }   // Поле FK в дочерней таблице
    }

    public class SaveResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public bool IsForeignKeyViolation { get; set; }
    }

    public class ProjectDataManager
    {
        private DataTable originalTable;
        private OleDbDataAdapter dataAdapter;
        private OleDbCommandBuilder commandBuilder;
        private string currentTableName;
        private readonly string connStr;
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

        /// <summary>
        /// Получает список пользовательских таблиц из БД.
        /// </summary>
        /// <param name="isAdmin">Если false — таблица "Пользователи" исключается.</param>
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

            var conn = new OleDbConnection(connStr);
            dataAdapter = new OleDbDataAdapter(sql, conn);
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

        /// <summary>
        /// Выполняет поиск по всем видимым столбцам. Возвращает DataView с отфильтрованными строками.
        /// </summary>
        /// <param name="searchText">Текст для поиска.</param>
        /// <param name="visibleColumnNames">Список имён видимых столбцов (поиск только по ним).</param>
        public DataView PerformSearch(string searchText, IEnumerable<string> visibleColumnNames)
        {
            if (originalTable == null)
                throw new InvalidOperationException("Таблица не выбрана.");

            var conditions = new List<string>();
            string escapedText = searchText.Replace("'", "''");
            bool isNumber = decimal.TryParse(searchText, out decimal numValue);
            bool isDate = DateTime.TryParse(searchText, out DateTime dateValue);

            var visibleSet = new HashSet<string>(visibleColumnNames);

            foreach (DataColumn col in originalTable.Columns)
            {
                if (!visibleSet.Contains(col.ColumnName))
                    continue;

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

        public void ResetSort()
        {
            IsSorted = false;
        }

        public void ResetSearch()
        {
            IsSearched = false;
        }

        public void MarkFiltered()
        {
            IsFiltered = true;
        }

        public void ResetAll()
        {
            IsFiltered = false;
            IsSorted = false;
            IsSearched = false;
        }

        #endregion

        #region Внешние ключи

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
                    cmd.Parameters.AddWithValue("@p1", pkValue);
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

        #region Сохранение изменений

        public SaveResult TrySaveChanges()
        {
            if (dataAdapter == null || commandBuilder == null || originalTable == null)
                return new SaveResult { Success = false, ErrorMessage = "Нет данных для сохранения." };

            try
            {
                DataTable changes = originalTable.GetChanges();
                if (changes != null)
                {
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
            return type == typeof(int) || type == typeof(long) || type == typeof(short) ||
                   type == typeof(byte) || type == typeof(decimal) || type == typeof(double) ||
                   type == typeof(float);
        }

        #endregion
    }
}