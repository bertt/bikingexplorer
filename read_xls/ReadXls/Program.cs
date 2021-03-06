﻿using Newtonsoft.Json;
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

            var from_day = new DateTime(2018, 3, 15);
            var to_day = new DateTime(2018, 3, 15);

            var telpuntdata = GetTelpuntData(workbook, from_day, to_day);
            WriteToCsv(telpuntdata, "tp.csv");

            GC.Collect();
            GC.WaitForPendingFinalizers();
            // Marshal.ReleaseComObject(range);
            // Marshal.ReleaseComObject(worksheet);
            workbook.Close();
            Marshal.ReleaseComObject(workbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        private static void WriteToCsv(List<Telpunt> telpunten, string file)
        {
            using (var w = new StreamWriter(file))
            {
                w.WriteLine($"id,latitude,longitude,time,direction,measurement");
                foreach (var telpunt in telpunten)
                {
                    for (int h = 0; h < 24; h++)
                    {
                        var day = telpunt.DateTime;
                        day=day.AddHours(h);
                        var day_String = day.ToString("yyyy-MM-ddTHH:mm:ss");
                        var richting1 = $"{telpunt.Id}, {telpunt.LatLon}, {day_String}, {telpunt.Richting1Dir}, {telpunt.Richting1Measurements[h]}";
                        var richting2 = $"{telpunt.Id}, {telpunt.LatLon}, {day_String}, {telpunt.Richting2Dir}, {telpunt.Richting2Measurements[h]}";
                        w.WriteLine(richting1);
                        w.WriteLine(richting2);
                    }
                    w.Flush();
                }
            }
        }

        private static List<Telpunt> GetTelpuntData(Excel.Workbook workbook, DateTime from_day, DateTime to_day)
        {
            var days = (to_day - from_day).Days;
            var result = new List<Telpunt>();

            for (var i = 0; i <= days; i++)
            {
                var day = from_day.AddDays(i);
                var tp_day = GetTelpuntDataDay(workbook, day);
                result.AddRange(tp_day);
            }
            return result;
        }

        private static List<Telpunt> GetTelpuntDataDay(Excel.Workbook workbook, DateTime day)
        {
            var first_day_column_name = "AE";
            var first_day_column = ExcelColumnNameToNumber(first_day_column_name);
            var first_day = new DateTime(2018, 1, 1);

            var days = (day - first_day).Days;
            var number = first_day_column + days;

            var result = new List<Telpunt>();
            for (var i = 2; i <= 16; i++)
            {
                Excel._Worksheet worksheet = workbook.Sheets[i];
                Excel.Range range = worksheet.UsedRange;

                var telpunt = GetTelpunt(workbook, i - 1);
                Console.WriteLine($"{day.ToString("yyyy-MM-dd")}: {telpunt.Name}");

                float forward;
                float backward;

                if (i == 14)
                {
                    forward = 270;
                    backward = 90;
                }
                else
                {
                    var directions = GetSharedStreetsDirections(telpunt.Id, telpunt.LatLon).Result;
                    forward = directions.forward;
                    backward = directions.backward;
                }
                telpunt.Forward = forward;
                telpunt.Backward = backward;
                telpunt.DateTime = day;

                telpunt.CheckDirections();

                telpunt.Richting1Measurements = GetMeasurements(range, number, 73, 96);
                telpunt.Richting2Measurements = GetMeasurements(range, number, 183, 207);

                result.Add(telpunt);
            }
            return result;
        }

        public static async Task<(float forward, float backward)> GetSharedStreetsDirections(int telpuntid, string latlon)
        {
            var key = "bdd23fa1-7ac5-4158-b354-22ec946bb575";
            // sample input: 52.079037, 5.081251
            var ll = latlon.Split(',');
            var url = $"https://api.sharedstreets.io/v0.1.0/match/point/{ll[1]},{ll[0]}?auth={key}&searchRadius=50&maxCandidates=5";

            var client = new HttpClient();

            var response = await client.GetAsync(url);

            var json = response.Content.ReadAsStringAsync().Result;

            var rootobject = JsonConvert.DeserializeObject<Rootobject>(json);

            var forward = (rootobject.features[0].properties.bearing);
            var backward = rootobject.features[1].properties.bearing;

            // something different in berenkuil....
            if (telpuntid == 9)
            {
                backward = rootobject.features[3].properties.bearing;
            }

            var t = (forward, backward);
            return t;
        }


        private static Telpunt GetTelpunt(Excel.Workbook workbook, int telpuntid)
        {
            Excel._Worksheet worksheet = workbook.Sheets[19];
            Excel.Range range = worksheet.UsedRange;

            var columnA = ExcelColumnNameToNumber("A");
            var columnB = ExcelColumnNameToNumber("B");
            var columnQ = ExcelColumnNameToNumber("Q");
            var columnF = ExcelColumnNameToNumber("F");
            var columnG = ExcelColumnNameToNumber("G");

            var row = telpuntid + 1;
            var val = (string)range.Cells[row, columnQ].Value;

            var richting1 = (string)range.Cells[row, columnF].Value;
            var richting2 = (string)range.Cells[row, columnG].Value;
            var name = (string)range.Cells[row, columnB].Value;
            var id = (int)range.Cells[row, columnA].Value;

            var telpunt = new Telpunt { Id = id, Name = name, LatLon = val, Richting1 = GetDirection(richting1), Richting2 = GetDirection(richting2) };
            return telpunt;
        }

        private static string GetDirection(string richting)
        {
            return richting.Split(' ')[1];
        }

        private static List<int> GetMeasurements(Excel.Range range, int columnnumber, int from, int to)
        {
            var numbers = new List<int>();
            for (var j = from; j <= to; j++)
            {
                var val = (int)range.Cells[j, columnnumber].Value;
                numbers.Add(val);
            }
            return numbers;
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
