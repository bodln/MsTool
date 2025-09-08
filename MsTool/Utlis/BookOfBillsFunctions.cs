using MsTool.Models;
using NPOI.OpenXmlFormats.Vml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows; // Ako koristite WinForms, menjajte u System.Windows.Forms

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

                List<DiffRecord> diffs = new List<DiffRecord>();

                foreach (var key in csvRecs.Keys)
                {
                    if (key == null)
                    {
                        Console.WriteLine("⚠ Nađen je NULL ključ u CSV podacima – preskačem taj zapis.");
                        continue;
                    }

                    var csv = csvRecs[key];
                    if (csv == null)
                    {
                        Console.WriteLine($"⚠ Null vrednost za ključ '{key}' u CSV podacima – preskačem.");
                        continue;
                    }

                    if (!xlsRecs.TryGetValue(key, out var xls))
                    {
                        xls = null;
                    }

                    double xVal = xls?.Value ?? 0;
                    double cSum = csv.SumValue;
                    string xOrig = xls?.OriginalKey ?? "";

                    bool equal = Math.Abs(xVal - cSum) <= 5.0;
                    bool doubleTake = false;

                    if (!equal)
                    {
                        var matchingKey = xlsRecs
                           .Where(kvp =>
                           {
                               if (kvp.Value == null)
                               {
                                   Console.WriteLine("⚠ Null vrednost u XLS podacima – preskačem poređenje.");
                                   return false;
                               }

                               bool valueMatch = Math.Abs(kvp.Value.Value - cSum) <= 5.0;
                               if (!valueMatch)
                                   return false;

                               if (!DateTime.TryParseExact(kvp.Value.Date, "dd-MM-yy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var xlsDate))
                                   return false;

                               var cleanCsvDate = csv.Date1?.TrimEnd('.') ?? "";
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
                            CsvOriginalKey = csv.OriginalKey ?? "Nema",
                            Pib = csv.Pib ?? "",
                            CsvDate1 = csv.Date1 ?? "",
                            CsvDate2 = csv.Date2 ?? "",
                            CompanyName = "",
                            DoubleTake = doubleTake,
                            Status = csv.Status ?? ""
                        });
                    }
                }

                foreach (var diff in diffs)
                {
                    if (string.IsNullOrWhiteSpace(diff.Pib))
                    {
                        diff.CompanyName = "";
                        Console.WriteLine($"⚠ Nedostaje PIB za poziciju {diff.Position}");
                        continue;
                    }

                    string name = null;
                    try
                    {
                        name = PibStore.Instance.Lookup(diff.Pib);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠ Greška pri čitanju iz PibStore za PIB {diff.Pib}: {ex.Message}");
                    }

                    if (string.IsNullOrEmpty(name))
                    {
                        try
                        {
                            name = await NbsPibLookup.LookupNameAsync(diff.Pib) ?? "";
                            PibStore.Instance.AddOrUpdate(diff.Pib, name ?? "");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"⚠ Greška pri online pretrazi PIB-a {diff.Pib}: {ex.Message}");
                        }
                    }

                    diff.CompanyName = name ?? "";
                }

                SaveDialog.ShowSaveDialog(diffs, allMistakes, assumptions);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Greška BookOfBillsFunctions.cs: " + ex.Message);
            }
        }
    }
}
