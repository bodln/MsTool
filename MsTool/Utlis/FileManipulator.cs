using ClosedXML.Excel;
using ExcelDataReader;
using MsTool.Models;
using System.Data;
using System.Text.RegularExpressions;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;

namespace MsTool.Utlis
{
    public static class FileManipulator
    {
        public static Dictionary<string, XlsRecord> LoadXls(string path, Dictionary<string, CsvRecord> csvRecs)
        {
            var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);
            var dict = new Dictionary<string, XlsRecord>();

            // Identify populated columns in rows 10 and 11
            var populatedColsRow10 = Enumerable.Range(1, ws.LastColumnUsed().ColumnNumber())
                .Where(col => !string.IsNullOrWhiteSpace(ws.Cell(10, col).GetString()))
                .ToArray();

            var populatedColsRow11 = Enumerable.Range(1, ws.LastColumnUsed().ColumnNumber())
                .Where(col => !string.IsNullOrWhiteSpace(ws.Cell(11, col).GetString()))
                .ToArray();

            int ifraCol = -1, // Column number of the bill number
                valueCol = -1,  // -||- of the main relevant value (Flag 1)
                substitCol2 = -1, // -||- of the values relevant to Flag 2
                substitCol3 = -1; // -||- of the values relevant to Flag 3

            ifraCol = populatedColsRow11[3];
            valueCol = populatedColsRow10[6];

            int dateCol1 = populatedColsRow10[3], // DATPRI
                dateCol2 = populatedColsRow10[4]; // DATDOK

            for (int row = 10; ; row += 2) // Data begins on row 10
            {
                var rawMarker = ws.Cell(row, populatedColsRow10[1]).GetString();
                if (string.IsNullOrWhiteSpace(rawMarker))
                    rawMarker = ws.Cell(++row, populatedColsRow10[1]).GetString();
                if (string.IsNullOrWhiteSpace(ws.Cell(row, populatedColsRow10[0]).GetString()))
                    break;

                var marker = rawMarker.Trim().ToUpper(); // ex. UN0

                string origKey = ws.Cell(row + 1, ifraCol).GetString(); // Original bill number
                string cleanKey = Regex.Replace(origKey, @"[\/\-\s]", "").ToUpperInvariant();
                double val = ParseCell(ws.Cell(row, valueCol).GetString());
                int flag = 1;

                if (csvRecs.ContainsKey(cleanKey) && csvRecs[cleanKey].Flag == 2)
                {
                    var populatedColsRowX = Enumerable.Range(valueCol, ws.LastColumnUsed().ColumnNumber())
                        .Where(col => !string.IsNullOrWhiteSpace(ws.Cell(row, col).GetString()))
                        .ToArray();
                    substitCol2 = populatedColsRowX[0]; // Column number of OSNOV OPSTA values
                    val = ParseCell(ws.Cell(row, substitCol2).GetString());
                    flag = 2;
                }
                else if (csvRecs.ContainsKey(cleanKey) && csvRecs[cleanKey].Flag == 3)
                {
                    var populatedColsRowX = Enumerable.Range(valueCol, ws.LastColumnUsed().ColumnNumber())
                        .Where(col => !string.IsNullOrWhiteSpace(ws.Cell(row, col).GetString()))
                        .ToArray(); // Assumes column where value for substitCol 3 would've been is empty
                    substitCol3 = populatedColsRowX[0]; // Column number of OSNOV POS. values
                    val = ParseCell(ws.Cell(row, substitCol3).GetString());
                    flag = 3;
                }


                var d1 = ws.Cell(row, dateCol1).GetString().Trim();
                var d2 = ws.Cell(row, dateCol2).GetString().Trim();

                // If a same bill number appears it prioritizes either the one with the UN0 marker, if such exists, or the first one appeared 
                if (!dict.ContainsKey(cleanKey))
                {
                    dict[cleanKey] = new XlsRecord(origKey, val, d2, marker, flag);
                }
                else if (marker == "UN0")
                {
                    dict[cleanKey] = new XlsRecord(origKey, val, d2, marker, flag);
                }
            }

            return dict;
        }

        public static Dictionary<string, CsvRecord> LoadCsv(string path)
        {
            var dict = new Dictionary<string, CsvRecord>();

            // detect delimiter by comparing semicolons vs commas on the header line
            var firstLine = File.ReadLines(path).First();
            var delim = firstLine.Count(c => c == ';') > firstLine.Count(c => c == ',') ? ';' : ',';

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = delim.ToString(),
                HasHeaderRecord = true
            };

            using var reader = new StreamReader(path);
            using var csv = new CsvReader(reader, config);

            csv.Read();
            csv.ReadHeader();

            while (csv.Read())
            {
                string origKey = csv.GetField("Broj dokumenta");
                string cleanKey = Regex.Replace(origKey, @"[\/\-\s]", "").ToUpperInvariant();

                double v1 = ParseCell(csv.GetField("PDV 20%")); // PDV 20%
                double v2 = ParseCell(csv.GetField("PDV 10%")); // PDV 10%
                double v3 = ParseCell(csv.GetField("Osnovica 20%")); // Osnovica 20%
                double v4 = ParseCell(csv.GetField("Osnovica 10%")); // Osnovica 10%
                double sum = v1 + v2 + v3 + v4;

                string date1 = csv.GetField("Datum PDV obaveze/evidentiranja"); // Datum PDV obaveze/evidentiranja
                string date2 = csv.GetField("Datum obrade"); // Datum obrade
                string position = csv.GetField(0);
                string pib = csv.GetField("PIB prodavca");

                if (v3 != 0 && v1 == 0)
                {
                    sum = v3;
                    dict[cleanKey] = new CsvRecord(origKey, sum, date1, date2, position, pib, 2); // See Flag correspondence above
                }
                else if (v4 != 0 && v2 == 0)
                {
                    sum = v4;
                    dict[cleanKey] = new CsvRecord(origKey, sum, date1, date2, position, pib, 3);
                }
                else
                {
                    dict[cleanKey] = new CsvRecord(origKey, sum, date1, date2, position, pib, 1);
                }
            }

            return dict;
        }

        public static string ConvertXlsToXlsx(string xlsFilePath)
        {
            var stream = File.Open(xlsFilePath, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateBinaryReader(stream);

            var config = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = false
                }
            };

            var dataSet = reader.AsDataSet(config);
            reader.Close();
            stream.Close();

            var wb = new XLWorkbook();

            foreach (DataTable table in dataSet.Tables)
            {
                var ws = wb.Worksheets.Add(table.TableName);
                for (int row = 0; row < table.Rows.Count; row++)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        var value = table.Rows[row][col]?.ToString() ?? "";
                        ws.Cell(row + 1, col + 1).Value = value;
                    }
                }
            }

            var newPath = Path.Combine(Path.GetTempPath(), // Saves in users temp files
                Path.GetFileNameWithoutExtension(xlsFilePath) + ".converted.xlsx");
            wb.SaveAs(newPath);
            return newPath;
        }

        private static double ParseCell(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            var trimmed = s.Trim();
            var parts = trimmed.Split(new[] { '.', ',' }, 2);
            var match = Regex.Match(parts[0], @"-?\d+");
            return match.Success && double.TryParse(match.Value, out var w) ? w : 0;
        }

        private static string Normalize(string input) =>
            Regex.Replace(input ?? "", @"[^\u0020-\u007E]", "")
                 .ToUpperInvariant();
    }
}
