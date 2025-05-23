using ClosedXML.Excel;
using MsTool.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsTool.Utlis
{
    public static class AnalyticsSaveDialog
    {
        public static void ShowSaveDialog(List<DiffAnalyticsRecord> diffs, bool showAssumptions)
        {
            int count = 0;

            if (!showAssumptions)
            {
                count = diffs.Count();
            }
            else
            {
                count = diffs.Where(d => d.DoubleTake == false).Count();
            }

            var dlg = new Form
            {
                Width = 360,
                Height = 200,
                Text = "Rezultat poređenja",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent
            };

            var lbl = new Label
            {
                Text = $"Broj razlika: {count}",
                AutoSize = true,
                Top = 20,
                Left = 20
            };
            dlg.Controls.Add(lbl);

            var btnDesk = new Button
            {
                Text = "Sačuvaj na Desktop",
                Width = 300,
                Top = 60,
                Left = 20
            };
            btnDesk.Click += (s, e) =>
            {
                var outp = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "ARazlike.xlsx");
                outp = GetUniquePath(outp);
                SaveDiff(outp, diffs, showAssumptions);
                dlg.Close();
            };
            dlg.Controls.Add(btnDesk);

            var btnFolder = new Button
            {
                Text = "Sačuvaj u folder",
                Width = 300,
                Top = 100,
                Left = 20
            };
            btnFolder.Click += (s, e) =>
            {
                var fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    var outp = Path.Combine(fbd.SelectedPath, "ARazlike.xlsx");
                    outp = GetUniquePath(outp);
                    SaveDiff(outp, diffs, showAssumptions);
                    dlg.Close();
                }
            };
            dlg.Controls.Add(btnFolder);

            var btnClose = new Button
            {
                Text = "Zatvori",
                Width = 300,
                Top = 140,
                Left = 20
            };
            btnClose.Click += (s, e) => dlg.Close();
            dlg.Controls.Add(btnClose);

            dlg.ShowDialog();
        }

        private static async void SaveDiff(string path, List<DiffAnalyticsRecord> diffs, bool showAssumptions)
        {
            var sortedDiffs = diffs.OrderBy(d => d.DateMain).ToList();

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Razlike");

            double sumMain = 0;
            double sumRef = 0;

            if (showAssumptions)
            {
                ws.Cell("A1").Value = "Prp.";
                ws.Cell("B1").Value = "Datum";
                ws.Cell("C1").Value = "Racun";
                ws.Cell("D1").Value = "Duguje";
                ws.Cell("E1").Value = "Potrazuje";
            }
            else
            {
                ws.Cell("A1").Value = "Datum";
                ws.Cell("B1").Value = "Racun";
                ws.Cell("C1").Value = "Duguje";
                ws.Cell("D1").Value = "Potrazuje";
            }

            int excelRow = 2;

            foreach (var diff in sortedDiffs)
            {
                if (!showAssumptions && diff.DoubleTake)
                {
                    continue;
                }

                sumMain += diff.ValueMain;
                sumRef += diff.ValueRef;

                if (showAssumptions)
                {
                    ws.Cell(excelRow, 1).Value = diff.DoubleTake ? "-->" : "";
                    ws.Cell(excelRow, 2).Value = diff.DateMain;
                    ws.Cell(excelRow, 3).Value = diff.OriginalMainKey;
                    ws.Cell(excelRow, 4).Value = diff.ValueMain;
                    ws.Cell(excelRow, 5).Value = diff.ValueRef;
                }
                else
                {
                    ws.Cell(excelRow, 1).Value = diff.DateMain;
                    ws.Cell(excelRow, 2).Value = diff.OriginalMainKey;
                    ws.Cell(excelRow, 3).Value = diff.ValueMain;
                    ws.Cell(excelRow, 4).Value = diff.ValueRef;
                }

                excelRow++;
            }

            if (showAssumptions)
            {
                ws.Cell(++excelRow, 3).Value = "Svega:";
                ws.Cell(excelRow, 4).Value = sumMain;
                ws.Cell(excelRow, 5).Value = sumRef;
            }
            else
            {
                ws.Cell(++excelRow, 2).Value = "Svega:";
                ws.Cell(excelRow, 3).Value = sumMain;
                ws.Cell(excelRow, 4).Value = sumRef;
            }

            ws.RangeUsed().SetAutoFilter();
            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
            MessageBox.Show("Uspešno sačuvano:\n" + path);
        }

        // Making sure there is no identical name conflicts
        private static string GetUniquePath(string basePath)
        {
            var dir = Path.GetDirectoryName(basePath);
            var name = Path.GetFileNameWithoutExtension(basePath);
            var ext = Path.GetExtension(basePath);
            int idx = 1;
            var p = basePath;
            while (File.Exists(p))
                p = Path.Combine(dir, $"{name}({idx++}){ext}");
            return p;
        }
    }
}
