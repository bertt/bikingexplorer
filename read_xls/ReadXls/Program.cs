using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadXls
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("read xls");
            var xls = @"D:\gisdata\utrecht\Utrecht.Fietstellingen.2018-09.xlsx";

            var xlApp = new Excel.Application();
            var workbook = xlApp.Workbooks.Open(xls);

            var number = ExcelColumnNameToNumber("CY");

            var telpuntdata  = GetTelpuntData(workbook, number);
            WriteToCsv(telpuntdata, "180314");

            GC.Collect();
            GC.WaitForPendingFinalizers();
            // Marshal.ReleaseComObject(range);
            // Marshal.ReleaseComObject(worksheet);
            workbook.Close();
            Marshal.ReleaseComObject(workbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private static void WriteToCsv(List<string> data, string day)
        {
            using (var w = new StreamWriter($"{day}_data.csv"))
            {
                foreach (var line in data)
                {
                    w.WriteLine(line);
                    w.Flush();
                }
            }
        }

        private static List<string> GetTelpuntData(Excel.Workbook workbook, int number)
        {
            var result = new List<String>();
            for (var i = 2; i <= 16; i++)
            {
                Excel._Worksheet worksheet = workbook.Sheets[i];
                var name = worksheet.Name;
                Console.WriteLine(name);
                Excel.Range range = worksheet.UsedRange;

                var latlon = GetLatLon(workbook, i - 1);

                // richting1
                var tp = "tp" + (i - 1).ToString();
                var richting1 = GetCvs(range, number, 73, 96);
                var richting2 = GetCvs(range, number, 183, 207);

                var tpresult = $"{tp}, {latlon}, r1({richting1}), r2({richting2})";
                result.Add(tpresult);
            }
            return result;
        }

        private static string GetLatLon(Excel.Workbook workbook, int telpunt)
        {
            Excel._Worksheet worksheet = workbook.Sheets[19];
            Excel.Range range = worksheet.UsedRange;

            var column = ExcelColumnNameToNumber("Q");
            var row = telpunt+1;
            var val = (string)range.Cells[row, column].Value;
            return val;
        }

        private static string GetCvs(Excel.Range range, int columnnumber, int from, int to)
        {
            var numbers = new List<int>();
            for (var j = from; j <= to; j++)
            {
                var val = (int)range.Cells[j, columnnumber].Value;
                numbers.Add(val);
            }
            return string.Join(", ", numbers);
        }

        public static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");
            char[] characters = columnName.ToUpperInvariant().ToCharArray();
            int sum = 0;
            for (int i = 0; i < characters.Length; i++)
            {
                sum *= 26;
                sum += (characters[i] - 'A' + 1);
            }
            return sum;  // in this example, sum would be "1" representing the column # where Customer Name resides 
        }

    }
}
