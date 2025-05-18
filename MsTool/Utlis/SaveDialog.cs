using ClosedXML.Excel;
using MsTool.Models;

namespace MsTool.Utlis
{
    public static class SaveDialog
    {
        public static void ShowSaveDialog(List<DiffRecord> diffs, bool includeAll, bool showAssumptions)
        {
            int count = includeAll
                        ? diffs.Count
                        : diffs.Count(d => d.XlsMarker == "Nema");

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
                    "Razlike.xlsx");
                outp = GetUniquePath(outp);
                SaveDiff(outp, diffs, includeAll, showAssumptions);
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
                    var outp = Path.Combine(fbd.SelectedPath, "Razlike.xlsx");
                    outp = GetUniquePath(outp);
                    SaveDiff(outp, diffs, includeAll, showAssumptions);
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

        private static async void SaveDiff(string path, List<DiffRecord> diffs, bool includeAll, bool showAssumptions)
        {
            var sortedDiffs = diffs.OrderBy(d => d.Pib).ToList();

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Razlike");

            if (showAssumptions)
            {
                ws.Cell("A1").Value = "Pretpostavka";
                ws.Cell("B1").Value = "Pozicija";
                ws.Cell("C1").Value = "Oznaka";
                ws.Cell("D1").Value = "Racun kod mene";
                ws.Cell("E1").Value = "Moja vrednost";
                ws.Cell("F1").Value = "Poreska suma";
                ws.Cell("G1").Value = "Racun poreska";
                ws.Cell("H1").Value = "Datum evidentiranja";
                ws.Cell("I1").Value = "Datum obrade";
                ws.Cell("J1").Value = "PIB";
                ws.Cell("K1").Value = "Naziv firme";
            }
            else
            {
                ws.Cell("A1").Value = "Pozicija";
                ws.Cell("B1").Value = "Oznaka";
                ws.Cell("C1").Value = "Moja vrednost";
                ws.Cell("D1").Value = "Poreska suma";
                ws.Cell("E1").Value = "Racun poreska";
                ws.Cell("F1").Value = "Datum evidentiranja";
                ws.Cell("G1").Value = "Datum obrade";
                ws.Cell("H1").Value = "PIB";
                ws.Cell("I1").Value = "Naziv firme";
            }

            int excelRow = 2;

            foreach (var diff in sortedDiffs)
            {
                if (!includeAll && diff.XlsMarker != "Nema")
                    continue;

                if (!showAssumptions && diff.DoubleTake)
                {
                    continue;
                }

                if (showAssumptions)
                {
                    ws.Cell(excelRow, 1).Value = diff.DoubleTake ? "-->" : "";
                    ws.Cell(excelRow, 2).Value = diff.Position;
                    ws.Cell(excelRow, 3).Value = diff.XlsMarker;
                    ws.Cell(excelRow, 4).Value = diff.XlsOriginalKey;
                    ws.Cell(excelRow, 5).Value = diff.XlsValue;
                    ws.Cell(excelRow, 6).Value = diff.CsvSumValue;
                    ws.Cell(excelRow, 7).Value = diff.CsvOriginalKey;
                    ws.Cell(excelRow, 8).Value = diff.CsvDate1;
                    ws.Cell(excelRow, 9).Value = diff.CsvDate2;
                    ws.Cell(excelRow, 10).Value = diff.Pib;
                    ws.Cell(excelRow, 11).Value = diff.CompanyName;
                }
                else
                {
                    ws.Cell(excelRow, 1).Value = diff.Position;
                    ws.Cell(excelRow, 2).Value = diff.XlsMarker;
                    ws.Cell(excelRow, 3).Value = diff.XlsValue;
                    ws.Cell(excelRow, 4).Value = diff.CsvSumValue;
                    ws.Cell(excelRow, 5).Value = diff.CsvOriginalKey;
                    ws.Cell(excelRow, 6).Value = diff.CsvDate1;
                    ws.Cell(excelRow, 7).Value = diff.CsvDate2;
                    ws.Cell(excelRow, 8).Value = diff.Pib;
                    ws.Cell(excelRow, 9).Value = diff.CompanyName;
                }

                excelRow++;
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
