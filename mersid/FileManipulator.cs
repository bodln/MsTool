using ClosedXML.Excel;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using mersid.Models;

namespace mersid
{
    public static class FileManipulator
    {
        public static Dictionary<string, XlsRecord> LoadXls(string path, Dictionary<string, CsvRecord> csvRecs)
        {
            var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);
            var dict = new Dictionary<string, XlsRecord>();

            int ifraCol = -1, valueCol = -1, substitCol2 = -1, substitCol3 = -1;
            for (int col = 1; col <= ws.LastColumnUsed().ColumnNumber(); col++)
            {
                string r9 = Normalize(ws.Cell(9, col).GetString());
                string r8 = Normalize(ws.Cell(8, col).GetString());
                if (ifraCol == -1 && (r9.Contains("IFRA") || r9.Contains("\u008aIFRA")))
                    ifraCol = col;
                if (valueCol == -1 && r8.Contains("VALUTA"))
                    valueCol = col;
                if (substitCol2 == -1 && r8.Contains("OSNOV OPSTA"))
                    substitCol2 = col;
                if (substitCol3 == -1 && r8.Contains("OSNOV POS."))
                    substitCol3 = col;
            }
            if (ifraCol < 0 || valueCol < 0)
                throw new Exception("Nije moguće pronaći kolone");

            for (int row = 10; ; row += 2)
            {
                var rawMarker = ws.Cell(row, "B").GetString();
                if (string.IsNullOrWhiteSpace(rawMarker))
                    rawMarker = ws.Cell(++row, "B").GetString();
                if (string.IsNullOrWhiteSpace(ws.Cell(row, "A").GetString()))
                    break;

                var marker = rawMarker.Trim().ToUpper();
                //if (checkBox1.Checked && marker != "UN0")
                //    continue;

                string origKey = ws.Cell(row + 1, ifraCol).GetString();
                string cleanKey = Regex.Replace(origKey, @"[\/\-\s]", "").ToUpperInvariant();
                double val = ParseCell(ws.Cell(row, valueCol).GetString());
                int flag = 1;

                if (csvRecs.ContainsKey(cleanKey) && csvRecs[cleanKey].Flag == 2)
                {
                    val = ParseCell(ws.Cell(row, substitCol2).GetString());
                    flag = 2;
                }
                else if (csvRecs.ContainsKey(cleanKey) && csvRecs[cleanKey].Flag == 3)
                {
                    val = ParseCell(ws.Cell(row, substitCol3).GetString());
                    flag = 3;
                }

                dict[cleanKey] = new XlsRecord(origKey, val, marker, flag);
            }

            return dict;
        }

        public static Dictionary<string, CsvRecord> LoadCsv(string path)
        {
            var dict = new Dictionary<string, CsvRecord>();
            var lines = File.ReadAllLines(path);
            foreach (var line in lines.Skip(1))
            {
                var parts = line.Split(',');
                if (parts.Length < 12) continue;

                string origKey = parts[1];
                string cleanKey = Regex.Replace(origKey, @"[\/\-\s]", "").ToUpperInvariant();

                double v1 = ParseCell(parts[9]);
                double v2 = ParseCell(parts[11]);
                double v3 = ParseCell(parts[8]);
                double v4 = ParseCell(parts[10]);
                double sum = v1 + v2 + v3 + v4;

                string date1 = parts[6];
                string date2 = parts[7];
                string position = parts[0];
                string pib = parts[5];

                if (v3 != 0 && v1 == 0)
                {
                    sum = v3;
                    dict[cleanKey] = new CsvRecord(origKey, sum, date1, date2, position, pib, 2);
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

            var newPath = Path.Combine(Path.GetTempPath(),
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
