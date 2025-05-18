using ClosedXML.Excel;
using mersid.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mersid
{
    public static class SaveDialog
    {
        public static void ShowSaveDialog(List<DiffRecord> diffs)
        {
            int count = diffs.Count;

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
                SaveDiff(outp, diffs);
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
                    SaveDiff(outp, diffs);
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

        private static async void SaveDiff(string path, List<DiffRecord> diffs)
        {
            var sortedDiffs = diffs.OrderBy(d => d.Pib).ToList();

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Razlike");
            ws.Cell("A1").Value = "Pozicija";
            ws.Cell("B1").Value = "Oznaka";
            ws.Cell("C1").Value = "Moja vrednost";
            ws.Cell("D1").Value = "Poreska suma";
            ws.Cell("E1").Value = "Racun poreska";
            ws.Cell("F1").Value = "Datum evidentiranja";
            ws.Cell("G1").Value = "Datum obrade";
            ws.Cell("H1").Value = "PIB";
            ws.Cell("I1").Value = "Naziv firme";

            for (int i = 0; i < sortedDiffs.Count; i++)
            {
                int r = i + 2;
                var diff = sortedDiffs[i];
                ws.Cell(r, 1).Value = diff.Position;
                ws.Cell(r, 2).Value = diff.Marker;
                ws.Cell(r, 3).Value = diff.XlsValue;
                ws.Cell(r, 4).Value = diff.CsvSumValue;
                ws.Cell(r, 5).Value = diff.CsvOriginalKey;
                ws.Cell(r, 6).Value = diff.CsvDate1;
                ws.Cell(r, 7).Value = diff.CsvDate2;
                ws.Cell(r, 8).Value = diff.Pib;
                ws.Cell(r, 9).Value = await NbsPibLookup.LookupNameAsync(diff.Pib);//ws.Cell(r, 7).Value = await GetPIB(diffs[i].Pib)
            }

            ws.RangeUsed().SetAutoFilter();
            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
            MessageBox.Show("Uspešno sačuvano:\n" + path);
        }

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
