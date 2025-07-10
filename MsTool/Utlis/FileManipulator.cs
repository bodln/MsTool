using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelDataReader;
using MsTool.Models;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace MsTool.Utlis
{
    public static class FileManipulator
    {
        public static Dictionary<string, XlsAnalyticsRecord> LoadXlsAnalytics(string path)
        {
            var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);
            var dict = new Dictionary<string, XlsAnalyticsRecord>();

            int lastRow = ws.LastRowUsed().RowNumber();
            int lastCol = ws.LastColumnUsed().ColumnNumber();

            for (int row = 18; row < lastRow; row++)
            {
                var lastCellValue = ws.Cell(row, lastCol).GetString().Trim();
                if (string.Equals(lastCellValue, "P.S.", StringComparison.OrdinalIgnoreCase))
                    continue;

                var colD = ws.Cell(row, "D").GetString().Trim();
                if (string.Equals(colD, "SRAVNJENJE", StringComparison.OrdinalIgnoreCase))
                    continue;

                string origKey = colD;
                string date = ws.Cell(row, "B").GetString().Trim();
                string account = ws.Cell(row, "K").GetString().Trim();

                double valueMain = ParseCell(ws.Cell(row, "F").GetString());

                double valueRef = ParseCell(ws.Cell(row, "H").GetString());

                string cleanKey = Regex.Replace(origKey, @"[^A-Za-z0-9]", "");

                dict[cleanKey] = new XlsAnalyticsRecord(
                    OriginalKey: origKey,
                    ValueMain: valueMain,
                    ValueRef: valueRef,
                    Date: date,
                    Account: account,
                    Flag: false
                );
            }

            return dict;
        }

        public static Dictionary<string, XlsRecord> LoadXls(string path, Dictionary<string, CsvRecord> csvRecs)
        {
            var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);
            var dict = new Dictionary<string, XlsRecord>();

            int startingRow = GetStartingRow(path);

            // Identify populated columns in rows 10 and 11
            var populatedColsRowFirst = Enumerable.Range(1, ws.LastColumnUsed().ColumnNumber())
                .Where(col => !string.IsNullOrWhiteSpace(ws.Cell(startingRow, col).GetString()))
                .ToArray();

            int firstUn0Row = FindFirstUn0Row(ws, startingRow, populatedColsRowFirst[1]);

            var populatedColsRowSecond = Enumerable.Range(1, ws.LastColumnUsed().ColumnNumber())
                .Where(col => !string.IsNullOrWhiteSpace(ws.Cell(firstUn0Row + 1, col).GetString()))
                .ToArray();

            int ifraCol = -1, // Column number of the bill number
                valueCol = -1,  // -||- of the main relevant value (Flag 1)
                substitCol2 = -1, // -||- of the values relevant to Flag 2
                substitCol3 = -1; // -||- of the values relevant to Flag 3

            ifraCol = populatedColsRowSecond[3];
            valueCol = populatedColsRowFirst[6];

            int dateCol1 = populatedColsRowFirst[3], // DATPRI
                dateCol2 = populatedColsRowFirst[4]; // DATDOK

            for (int row = startingRow; ; row += 2) // Data begins on row 10
            {
                var rawMarker = ws.Cell(row, populatedColsRowFirst[1]).GetString();
                if (string.IsNullOrWhiteSpace(rawMarker))
                    rawMarker = ws.Cell(++row, populatedColsRowFirst[1]).GetString();
                if (string.IsNullOrWhiteSpace(ws.Cell(row, populatedColsRowFirst[0]).GetString()))
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

                string status = csv.GetField("Status");

                if (status == null)
                {
                    status = "Nema kolona";
                }

                if (v3 != 0 && v1 == 0)
                {
                    sum = v3;
                    dict[cleanKey] = new CsvRecord(origKey, sum, date1, date2, position, pib, 2, status); // See Flag correspondence above
                }
                else if (v4 != 0 && v2 == 0)
                {
                    sum = v4;
                    dict[cleanKey] = new CsvRecord(origKey, sum, date1, date2, position, pib, 3, status);
                }
                else
                {
                    dict[cleanKey] = new CsvRecord(origKey, sum, date1, date2, position, pib, 1, status);
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

            var normalized = s
                .Replace(" ", "")
                .Replace(",", ".");

            return double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out var result)
                ? result
                : 0;
        }

        private static string Normalize(string input) =>
            Regex.Replace(input ?? "", @"[^\u0020-\u007E]", "")
                 .ToUpperInvariant();

        public static int GetStartingRow(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found at {filePath}");
                return -1;
            }

            try
            {
                using (var wb = new XLWorkbook(filePath))
                {
                    var ws = wb.Worksheet(1);

                    foreach (var row in ws.RowsUsed())
                    {
                        var cellA = row.Cell(1);

                        if (cellA.TryGetValue(out double numericValue) && numericValue == 1.0)
                        {
                            return row.RowNumber();
                        }
                        else if (cellA.GetString().Equals("1", StringComparison.OrdinalIgnoreCase))
                        {
                            return row.RowNumber();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while reading the Excel file to find starting row: {ex.Message}");
                return -1;
            }

            return -1;
        }
        public static int FindFirstUn0Row(IXLWorksheet ws, int startingRow, int targetCol)
        {
            for (int row = startingRow; row <= ws.LastRowUsed().RowNumber(); row++)
            {
                string cellValue = ws.Cell(row, targetCol).GetString()?.Trim().ToUpperInvariant();
                if (cellValue == "UN0")
                {
                    return row;
                }
            }

            throw new Exception("Nema UN0 markera");
        }

    }
}
