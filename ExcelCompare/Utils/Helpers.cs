using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ExcelDataReader;
using McMaster.Extensions.CommandLineUtils;

namespace ExcelCompare.Utils
{
    public static class Helpers
    {
        private const string UniqueIdentifierName = "UNIQUE_IDENTIFIER";

        public static string GetValidExcelFile(string filePath, string fileOrder, bool isDuplicate = false)
        {
            var newFile = filePath;
            var secondaryText1 = isDuplicate ? "is the same as the first Excel file" : "was either missing or invalid";
            var secondaryText2 = isDuplicate ? "UNIQUE and proper" : "proper";

            while (!File.Exists(newFile) || File.Exists(filePath) && Path.GetExtension(newFile) != ".xlsx")
            {
                newFile = Prompt.GetString($"The {fileOrder.ToUpper()} {secondaryText1}. Please {Environment.NewLine}provide a {secondaryText2} Excel file, or type eXit to end:",
                            promptColor: ConsoleColor.Black,
                            promptBgColor: ConsoleColor.Yellow);

                if (IsExit(newFile))
                {
                    newFile = null;
                    break;
                }
            }

            return newFile;
        }

        private static bool IsExit(string text)
        {
            if (text.Equals("x", StringComparison.InvariantCultureIgnoreCase) || 
                text.Equals("exit", StringComparison.InvariantCultureIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static bool IsNotValidColumnIndex(string userColumnIndex)
        {
            return userColumnIndex == null || string.IsNullOrWhiteSpace(userColumnIndex) ||
                userColumnIndex.Any(char.IsDigit);
        }

        private static bool IsNotLocatedColumnIndex(int indexValue, DataTable dataTable)
        {
            return indexValue > dataTable.Columns.Count;
        }

        private static bool IsNotValidRowIndex(string rowIndex)
        {
            return rowIndex == null || string.IsNullOrWhiteSpace(rowIndex) ||
                        !rowIndex.All(char.IsDigit);
        }

        private static bool IsNotLocatedRowIndex(string rowIndex, DataTable dataTable)
        {
            return Convert.ToInt32(rowIndex) > dataTable.Rows.Count;
        }

        public static DataSet GetDataSet(string filePath, string columnRowKey, string rowColumnKey)
        {
            //var keyColumnIndex = GetExcelColumnLettersAsIndex("A");
            //var headerRowIndex = 1;

            using FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            reader.AsDataSet();
            var result = reader.AsDataSet();

            var tablesToAdd = new List<DataTable>();

            foreach (DataTable table in result.Tables)
            {
                int keyColumnIndex;
                int headerRowIndex;

                if (IsNotValidColumnIndex(columnRowKey) || IsNotLocatedColumnIndex(GetExcelColumnLettersAsIndex(columnRowKey), table))
                {
                    var userProvidedColumnIndex = string.Empty;

                    var promptColumnText = $"Provide key COLUMN letter(s) for: '{Path.GetFileNameWithoutExtension(filePath)} | {table.TableName}', or type eXit to end:";
                    var wrongPromptColumnText = string.Empty;

                    while (true)
                    {
                        userProvidedColumnIndex = Prompt.GetString($"{wrongPromptColumnText}{promptColumnText}",
                                promptColor: ConsoleColor.Black,
                                promptBgColor: ConsoleColor.Yellow);

                        wrongPromptColumnText = "Invalid key COLUMN letters. ";

                        if (IsNotValidColumnIndex(userProvidedColumnIndex))
                        {
                            continue;
                        }

                        if (IsExit(userProvidedColumnIndex))
                        {
                            userProvidedColumnIndex = null;
                            return null;
                        }

                        var indexValue = GetExcelColumnLettersAsIndex(userProvidedColumnIndex);

                        if (IsNotLocatedColumnIndex(indexValue, table))
                        {
                            continue;
                        }
                        else
                        {
                            keyColumnIndex = indexValue;
                            break;
                        }
                    }
                }
                else
                {
                    keyColumnIndex = GetExcelColumnLettersAsIndex(columnRowKey);
                }

                if (IsNotValidRowIndex(rowColumnKey) || IsNotLocatedRowIndex(rowColumnKey, table))
                {
                    var userProvidedRowIndex = string.Empty;

                    var promptRowText = $"Provide ROW key index for: '{Path.GetFileNameWithoutExtension(filePath)} | {table.TableName}', or type eXit to end:";
                    var wrongPromptRowText = string.Empty;

                    while (true)
                    {
                        userProvidedRowIndex = Prompt.GetString($"{wrongPromptRowText}{promptRowText}",
                                promptColor: ConsoleColor.Black,
                                promptBgColor: ConsoleColor.Yellow);

                        wrongPromptRowText = "Invalid ROW. ";

                        if (IsNotValidRowIndex(userProvidedRowIndex))
                        {
                            continue;
                        }

                        if (IsExit(userProvidedRowIndex))
                        {
                            userProvidedRowIndex = null;
                            return null;
                        }

                        if (IsNotLocatedRowIndex(userProvidedRowIndex, table))
                        {
                            continue;
                        }
                        else
                        {
                            headerRowIndex = Convert.ToInt32(userProvidedRowIndex) - 1;
                            break;
                        }
                    }
                }
                else
                {
                    headerRowIndex = Convert.ToInt32(rowColumnKey) - 1;
                }

                if (keyColumnIndex > table.Columns.Count)
                {
                    Console.WriteLine("Invalid key COLUMN selection, please try again.");
                    return null;
                }

                if (headerRowIndex > table.Rows.Count - 1)
                {
                    Console.WriteLine("Invalid key ROW selection, please try again.");
                    return null;
                }

                var headerRowCopy = table.Rows[headerRowIndex].ItemArray.Clone() as object[];

                foreach (DataColumn dataColumn in table.Columns)
                {
                    dataColumn.ColumnName = $"{{{GetExcelColumnLetters(dataColumn.Ordinal + 1)}}} - {headerRowCopy[dataColumn.Ordinal]}";
                }

                var uniqueIdentifierColumn = table.Columns.Add(UniqueIdentifierName, typeof(int));
                uniqueIdentifierColumn.SetOrdinal(0);

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    table.Rows[i][uniqueIdentifierColumn] = i + 1;
                }

                foreach (DataRow dataRow in table.Rows)
                {
                    for (int i = 1; i < dataRow.ItemArray.Length; i++)
                    {
                        var currentValue = dataRow[i];

                        if (i == keyColumnIndex) continue;
                        if (currentValue != null && currentValue != DBNull.Value && currentValue.ToString().ToLower() == "null" ||
                            (currentValue != null && double.TryParse(currentValue.ToString(), out var doubleValue)))
                        {
                            break;
                        }

                        dataRow.Delete();
                        break;
                    }
                }

                table.AcceptChanges();

                foreach (DataRow dataRow in table.Rows)
                {
                    for (int i = 1; i < dataRow.ItemArray.Length; i++)
                    {
                        var currentValue = dataRow[i];

                        if (i == keyColumnIndex) continue;

                        var valueToAssign = 0.0;

                        if (currentValue != DBNull.Value && currentValue != null && currentValue.ToString() is string stringValue &&
                            double.TryParse(stringValue, out var doubleValue))
                        {
                            valueToAssign = doubleValue;
                        }

                        dataRow[i] = valueToAssign;
                    }
                }

                table.AcceptChanges();

                DataTable clonedTable = table.Clone();

                for (int i = 2; i < clonedTable.Columns.Count; i++)
                {
                    clonedTable.Columns[i].DataType = typeof(double);
                }

                foreach (DataRow row in table.Rows)
                {
                    clonedTable.ImportRow(row);
                }

                tablesToAdd.Add(clonedTable);
            }

            result.Tables.Clear();
            result.Tables.AddRange(tablesToAdd.ToArray());
            return result;
        }

        private static string GetExcelColumnLetters(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        public static int GetExcelColumnLettersAsIndex(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += columnName[i] - 'A' + 1;
            }

            return sum;
        }
    }
}
