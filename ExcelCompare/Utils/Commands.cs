using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ConsoleAppFramework;
using CommandAttribute = ConsoleAppFramework.CommandAttribute;
using OptionAttribute = ConsoleAppFramework.OptionAttribute;
using ExcelCompare.Extensions;
using McMaster.Extensions.CommandLineUtils;
using System.Diagnostics;

namespace ExcelCompare.Utils
{
    public class Commands : ConsoleAppBase
    {
        [Command("compare", "Cleanses and compares two Excel Files and returns their differences")]
        public void Compare(
                [Option("f1", "First Excel file path")] string file1Path = "",
                [Option("f1ck", "First Excel file column key (leave default to use column 'A')")] string file1RowKey = "A",
                [Option("f1rk", "First Excel file row key (leave default to use row 1)")] string file1ColumnKey = "3",
                [Option("f2", "Second Excel file path")] string file2Path = "",
                [Option("f2ck", "Second Excel file column key (leave default to use column 'A')")] string file2RowKey = "A",
                [Option("f2rk", "Second Excel file row key (leave default to use row 1)")] string file2ColumnKey = "3",
                [Option("o", "Specify output CSV results location (leave as default to output to current directory)")] string outputLocation = null
            )
        {
            try
            {
                outputLocation = outputLocation ??
                $@"{Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)}\COMPARE-{DateTime.Now:yyyyMMddHHmmssfff}.csv";

                Console.WriteLine("Validating first Excel file...");
                var finalFile1Path = Helpers.GetValidExcelFile(file1Path, "first");
                if (finalFile1Path is null) return;
                Console.WriteLine("Validating second Excel file...");
                var finalFile2Path = Helpers.GetValidExcelFile(file2Path, "second");
                if (finalFile2Path is null) return;

                while (finalFile1Path == finalFile2Path && finalFile2Path != null)
                {
                    finalFile2Path = Helpers.GetValidExcelFile(string.Empty, "second", true);
                    if (finalFile2Path is null) return;
                }

                Console.WriteLine("Extracting data from first Excel file...");
                var dataSet1 = Helpers.GetDataSet(finalFile1Path, file1RowKey, file1ColumnKey);
                if (dataSet1 is null) return;
                Console.WriteLine("Extracting data from second Excel file...");
                var dataSet2 = Helpers.GetDataSet(finalFile2Path, file2RowKey, file2ColumnKey);
                if (dataSet2 is null) return;

                var finalDataTable = new DataTable();
                var sheetColumn = new DataColumn { ColumnName = "sht", DataType = typeof(string) };
                var letterColumn = new DataColumn { ColumnName = "ltr", DataType = typeof(string) };
                var numberColumn = new DataColumn { ColumnName = "num", DataType = typeof(string) };
                var columnCaptionColumn = new DataColumn { ColumnName = "cCap", DataType = typeof(string) };
                var rowCaptionColumn = new DataColumn { ColumnName = "rCap", DataType = typeof(string) };
                var value1Column = new DataColumn { ColumnName = "v1", DataType = typeof(double) };
                var value2Column = new DataColumn { ColumnName = "v2", DataType = typeof(double) };
                var differenceColumn = new DataColumn { ColumnName = "d", DataType = typeof(double) };
                var absDifferenceColumn = new DataColumn { ColumnName = "abs", DataType = typeof(double) };
                finalDataTable.Columns.AddRange(new[]
                {
                sheetColumn,
                letterColumn,
                numberColumn,
                columnCaptionColumn,
                rowCaptionColumn,
                value1Column,
                value2Column,
                differenceColumn,
                absDifferenceColumn
            });

                var counter = 0;

                foreach (DataTable table1 in dataSet1.Tables)
                {
                    Console.WriteLine($"Analyzing and comparing data from matching sheets ({table1.TableName})...");

                    if (dataSet2.Tables.Contains(table1.TableName) &&
                        dataSet2.Tables[table1.TableName] is DataTable table2)
                    {
                        var doubleColumns = new List<DataColumn>();

                        foreach (DataColumn column in table1.Columns)
                        {
                            if (column.DataType == typeof(double))
                            {
                                doubleColumns.Add(column);
                            }
                        }

                        foreach (DataRow row in table1.Rows)
                        {

                            foreach (DataColumn column in doubleColumns)
                            {
                                var thisRowIndex = table1.Rows.IndexOf(row);
                                if (row[column] is double double1Value &&
                                    table2.Columns.Contains(column.ColumnName) &&
                                    table2.Rows.Count >= thisRowIndex &&
                                    table2.Rows[thisRowIndex][column.ColumnName] is double double2Value &&
                                    double1Value != double2Value)
                                {
                                    var lastIndexOf = column.ColumnName.LastIndexOf("} - ");
                                    var columnSubString = column.ColumnName.Substring(1, lastIndexOf - 1);

                                    var finalDataTableRow = finalDataTable.NewRow();

                                    finalDataTableRow[sheetColumn] = table1.TableName;
                                    finalDataTableRow[letterColumn] = columnSubString;
                                    finalDataTableRow[numberColumn] = table1.Rows[thisRowIndex][0];
                                    finalDataTableRow[columnCaptionColumn] = column.ColumnName.Replace($"{{{columnSubString}}} - ", string.Empty);
                                    finalDataTableRow[rowCaptionColumn] = table1.Rows[thisRowIndex][1];
                                    finalDataTableRow[value1Column] = double1Value;
                                    finalDataTableRow[value2Column] = double2Value;
                                    finalDataTableRow[differenceColumn] = double1Value - double2Value;
                                    finalDataTableRow[absDifferenceColumn] = Math.Abs(double1Value - double2Value);
                                    
                                    finalDataTable.Rows.Add(finalDataTableRow);
                                    counter++;
                                    Console.WriteLine($"Completed analyzing difference #{counter}");
                                }
                            }
                        }
                    }
                }

                if(counter == 0)
                {
                    Console.WriteLine($"No differences found, the program will now exit...");
                }
                else
                {
                    Console.WriteLine($"Sorting data...");
                    var dataView = finalDataTable.DefaultView;
                    dataView.Sort = $"{sheetColumn.ColumnName}, {absDifferenceColumn.ColumnName} desc";
                    var toPrintDataTable = dataView.ToTable();
                    Console.WriteLine("Creatng CSV file...");

                    toPrintDataTable.WriteToCsvFile(outputLocation);
                    Console.WriteLine($"File successfully written to {outputLocation}...");
                    if (Prompt.GetYesNo($"Success!! Found {counter} difference(s). Launch file now?",
                        false, ConsoleColor.Black, ConsoleColor.Green))
                    {
                        var process = new Process();
                        process.StartInfo = new ProcessStartInfo(outputLocation)
                        {
                            UseShellExecute = true
                        };
                        process.Start();
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Unhandles Exception! ({exception.Message})");
            }
        }
    }
}
