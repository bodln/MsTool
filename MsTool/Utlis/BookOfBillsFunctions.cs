using MsTool.Models;
using NPOI.OpenXmlFormats.Vml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsTool.Utlis
{
    public static class BookOfBillsFunctions
    {
        public async static void Proceed(string xlsPath, string csvPath, bool allMistakes, bool assumptions)
        {
            if (string.IsNullOrEmpty(xlsPath) || string.IsNullOrEmpty(csvPath))
            {
                MessageBox.Show("Molim vas, prvo izaberite oba fajla.");
                return;
            }

            try
            {
                var csvRecs = FileManipulator.LoadCsv(csvPath);
                var xlsRecs = FileManipulator.LoadXls(xlsPath, csvRecs);

                //foreach (var kvp in csvRecs)
                //{
                //    Console.WriteLine($"Key: {kvp.Key}, Value: {kvp.Value}");
                //}

                //foreach (var kvp in xlsRecs)
                //{
                //    Console.WriteLine($"Key: {kvp.Key}, Value: {kvp.Value}");
                //}

                List<DiffRecord>  diffs = new List<DiffRecord>();

                foreach (var key in csvRecs.Keys)
                {
                    var csv = csvRecs[key];
                    xlsRecs.TryGetValue(key, out var xls);

                    double xVal = xls?.Value ?? 0;
                    double cSum = csv.SumValue;
                    string xOrig = xls?.OriginalKey ?? "";

                    bool equal = Math.Abs(xVal - cSum) <= 5.0;
                    bool doubleTake = false; // Flag used to signal second round of searches was done, the one by value and date pair

                    if (!equal)
                    {
                        var matchingKey = xlsRecs
                           .Where(kvp =>
                           {
                               bool valueMatch = Math.Abs(kvp.Value.Value - cSum) <= 5.0;

                               if (!valueMatch)
                                   return false;

                               if (!DateTime.TryParseExact(kvp.Value.Date, "dd-MM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var xlsDate))
                                   return false;

                               var cleanCsvDate = csv.Date1.TrimEnd('.');

                               if (!DateTime.TryParseExact(cleanCsvDate, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var csvDate))
                                   return false;

                               return xlsDate.Date == csvDate.Date;
                           })
                           .OrderByDescending(kvp => kvp.Value.Marker == "UN0")
                           .Select(kvp => kvp.Key)
                           .FirstOrDefault();

                        if (matchingKey != null)
                        {
                            doubleTake = true;
                            xls = xlsRecs[matchingKey];
                            xOrig = xls.OriginalKey;
                            xVal = xls.Value;
                        }
                    }

                    if (xls == null || !equal)
                    {
                        diffs.Add(new DiffRecord
                        {
                            Position = csv.Position,
                            XlsMarker = xls?.Marker ?? "Nema",
                            XlsOriginalKey = xOrig ?? "Nema",
                            XlsValue = xVal,
                            CsvSumValue = cSum,
                            CsvOriginalKey = csv.OriginalKey,
                            Pib = csv.Pib,
                            CsvDate1 = csv.Date1,
                            CsvDate2 = csv.Date2,
                            CompanyName = "",
                            DoubleTake = doubleTake,
                            Status = csv.Status
                        });
                    }
                }

                foreach (var diff in diffs)
                {
                    if (string.IsNullOrWhiteSpace(diff.Pib))
                    {
                        diff.CompanyName = "";
                        continue;
                    }

                    var name = PibStore.Instance.Lookup(diff.Pib);

                    if (name == null)
                    {
                        name = await NbsPibLookup.LookupNameAsync(diff.Pib) ?? "";

                        PibStore.Instance.AddOrUpdate(diff.Pib, name ?? "");
                    }

                    diff.CompanyName = name;
                }

                SaveDialog.ShowSaveDialog(diffs, allMistakes, assumptions);

                //IReadOnlyDictionary<string, string> db = PibStore.Instance.GetAll();
                //db.ToList().ForEach(kvp => Console.WriteLine($"{kvp.Key} -> {kvp.Value}"));

                //if (File.Exists(xlsPath))
                //    File.Delete(xlsPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Greška BookOfBillsFunctions.cs: " + ex.Message);
            }
        }
    }
}
