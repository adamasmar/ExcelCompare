using System;
using System.Data;
using System.IO;
using System.Linq;
using ConsoleAppFramework;
using Newtonsoft.Json;
using CommandAttribute = ConsoleAppFramework.CommandAttribute;
using OptionAttribute = ConsoleAppFramework.OptionAttribute;

namespace ExcelCompare.Utils
{

    // sample change
    public class Commands : ConsoleAppBase
    {
        [Command("compare", "Cleanses and compares two Excel Files and returns their differences")]
        public void Compare(
            [Option("f1", "Path of first Excel file")] string file1Path = "",
            [Option("f2", "Path of second Excel file")] string file2Path = "",
            [Option("d", "Use default column and row indexes per file, per sheet (leave as default to use when scripting)")] bool defaultIndex = true,
            [Option("o", "Specify output CSV results location (leave as default to output to current directory)")] string outputLocation = null)
        {
            outputLocation = outputLocation ?? 
                $@"{Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)}\COMPARE-{DateTime.Now:yyyyMMddHHmmssfff}.csv";

            var finalFile1Path = Helpers.GetValidExcelFile(file1Path, "first");
            if (finalFile1Path is null) return;
            var finalFile2Path = Helpers.GetValidExcelFile(file2Path, "second");
            if (finalFile2Path is null) return;

            while (finalFile1Path == finalFile2Path && finalFile2Path != null)
            {
                finalFile2Path = Helpers.GetValidExcelFile(string.Empty, "second", true);
                if (finalFile2Path is null) return;
            }

            var dataSet1 = Helpers.GetDataSet(finalFile1Path, defaultIndex);
            if (dataSet1 is null) return;
            var dataSet2 = Helpers.GetDataSet(finalFile2Path, defaultIndex);
            if (dataSet2 is null) return;

            CustomDataRowComparer myDRComparer = new CustomDataRowComparer();
            var result2 = dataSet1.Tables[0].AsEnumerable().Except(dataSet2.Tables[0].AsEnumerable(), myDRComparer).CopyToDataTable();



            var json = JsonConvert.SerializeObject(dataSet1, Formatting.Indented);

            Console.WriteLine("Done to here.");
            Console.ReadLine();
        }
    }
}
