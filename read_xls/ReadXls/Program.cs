using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
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
                Excel.Range range = worksheet.UsedRange;

                var latlon = GetLatLon(workbook, i - 1);
                Console.WriteLine($"{name}: {latlon}");

                float forward;
                float backward;

                if(i == 14)
                {
                    forward = 90;
                    backward = 270;
                }
                else
                {
                    var directions = GetSharedStreetsDirections(latlon).Result;
                    forward = directions.forward;
                    backward = directions.backward;
                }


                // richting1
                var tp = "tp" + (i - 1).ToString();
                var richting1 = GetCvs(range, number, 73, 96);
                var richting2 = GetCvs(range, number, 183, 207);

                var tpresult = $"{tp}, {latlon}, r1, {forward}, ({richting1}), r2, {backward},({richting2})";
                result.Add(tpresult);
            }
            return result;
        }


        public static async Task<(float forward, float backward)> GetSharedStreetsDirections(string latlon)
        {
            var key = "bdd23fa1-7ac5-4158-b354-22ec946bb575";
            // sample input: 52.079037, 5.081251
            var ll= latlon.Split(',');
            var url = $"https://api.sharedstreets.io/v0.1.0/match/point/{ll[1]},{ll[0]}?auth={key}&searchRadius=50&maxCandidates=5";

            var client = new HttpClient();

            var response = await client.GetAsync(url);

            var json = response.Content.ReadAsStringAsync().Result;

            var rootobject = JsonConvert.DeserializeObject<Rootobject>(json);

            var forward = (rootobject.features[0].properties.bearing);
            var backward = rootobject.features[1].properties.bearing;
            var t = (forward, backward);
            return t;
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
