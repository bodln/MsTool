using ClosedXML.Excel;
using ExcelDataReader;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Data;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;


namespace mersid
{
    public partial class Form1 : Form
    {
        private string xlsPath;
        private string csvPath;
        private List<DiffRecord> diffs;

        // Selenium driver fields
        private readonly ChromeDriverService _driverService;
        private readonly ChromeOptions _chromeOptions;
        private readonly IWebDriver _driver;

        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        public Form1()
        {
            InitializeComponent();
            AllocConsole();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // ─── SETUP SELENIUM ONCE ───────────────────────────────────────────
            _driverService = ChromeDriverService.CreateDefaultService();
            _driverService.HideCommandPromptWindow = true;
            _driverService.SuppressInitialDiagnosticInformation = true;

            _chromeOptions = new ChromeOptions();
            _chromeOptions.AddArgument("--headless");
            _chromeOptions.AddArgument("--disable-gpu");
            _chromeOptions.AddArgument("--disable-extensions");
            _chromeOptions.AddArgument("--disable-popup-blocking");
            _chromeOptions.AddArgument("--log-level=3");
            _chromeOptions.PageLoadStrategy = PageLoadStrategy.Eager;

            _driver = new ChromeDriver(_driverService, _chromeOptions);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _driver?.Quit();
            _driver?.Dispose();
            base.OnFormClosing(e);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog { Filter = "Excel 97-2003 Workbook|*.xls" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                xlsPath = ConvertXlsToXlsx(ofd.FileName);
                label1.Text = Path.GetFileName(xlsPath);
            }
        }

        private string ConvertXlsToXlsx(string xlsFilePath)
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

        private void button2_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog { Filter = "CSV File|*.csv" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                csvPath = ofd.FileName;
                label2.Text = Path.GetFileName(csvPath);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(xlsPath) || string.IsNullOrEmpty(csvPath))
            {
                MessageBox.Show("Molim vas, prvo izaberite oba fajla.");
                return;
            }

            try
            {
                var csvRecs = LoadCsv(csvPath);
                var xlsRecs = LoadXls(xlsPath, csvRecs);

                foreach (var kvp in csvRecs)
                {
                    Console.WriteLine($"Key: {kvp.Key}, Value: {kvp.Value}");
                }

                foreach (var kvp in xlsRecs)
                {
                    Console.WriteLine($"Key: {kvp.Key}, Value: {kvp.Value}");
                }

                diffs = new List<DiffRecord>();

                foreach (var key in csvRecs.Keys)
                {
                    var csv = csvRecs[key];
                    xlsRecs.TryGetValue(key, out var xls);

                    double xVal = xls?.Value ?? 0;
                    double cSum = csv.SumValue;
                    string xOrig = xls?.OriginalKey ?? "";

                    //if (xls != null && checkBox1.Checked && xls.Marker != "UN0" && !ValuesEqual(xVal, cSum))
                    //    continue;

                    if (xls == null || !ValuesEqual(xVal, cSum)) // <----- this will need an UN0 check if you want only them
                    {
                        diffs.Add(new DiffRecord
                        {
                            Position = csv.Position,
                            Marker = xls?.Marker ?? "Nema",
                            OriginalKey = xOrig,
                            XlsValue = xVal,
                            CsvSumValue = cSum,
                            CsvOriginalKey = csv.OriginalKey,
                            Pib = csv.Pib,
                            CsvDate1 = csv.Date1,
                            CsvDate2 = csv.Date2
                        });
                    }
                }

                ShowSaveDialog(diffs.Count);

                if (File.Exists(xlsPath))
                    File.Delete(xlsPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Greška: " + ex.Message);
            }
        }

        private Dictionary<string, XlsRecord> LoadXls(string path, Dictionary<string, CsvRecord> csvRecs)
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
                }else if (csvRecs.ContainsKey(cleanKey) && csvRecs[cleanKey].Flag == 3)
                {
                    val = ParseCell(ws.Cell(row, substitCol3).GetString());
                    flag = 3;
                }

                dict[cleanKey] = new XlsRecord(origKey, val, marker, flag);
            }

            return dict;
        }

        private Dictionary<string, CsvRecord> LoadCsv(string path)
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

        private double ParseCell(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            var trimmed = s.Trim();
            var parts = trimmed.Split(new[] { '.', ',' }, 2);
            var match = Regex.Match(parts[0], @"-?\d+");
            return match.Success && double.TryParse(match.Value, out var w) ? w : 0;
        }

        private bool ValuesEqual(double a, double b) =>
            Math.Abs(a - b) <= 1.0;

        private void ShowSaveDialog(int count)
        {
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
                SaveDiff(outp);
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
                    SaveDiff(outp);
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

        private async void SaveDiff(string path)
        {
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

            for (int i = 0; i < diffs.Count; i++)
            {
                int r = i + 2;
                ws.Cell(r, 1).Value = diffs[i].Position;
                ws.Cell(r, 2).Value = diffs[i].Marker;
                ws.Cell(r, 3).Value = diffs[i].XlsValue;
                ws.Cell(r, 4).Value = diffs[i].CsvSumValue;
                ws.Cell(r, 5).Value = diffs[i].CsvOriginalKey;
                ws.Cell(r, 6).Value = diffs[i].CsvDate1;
                ws.Cell(r, 7).Value = diffs[i].CsvDate2;
                ws.Cell(r, 8).Value = diffs[i].Pib;
                //ws.Cell(r, 7).Value = await GetPIB(diffs[i].Pib);
                ws.Cell(r, 9).Value = await NbsPibLookup.LookupNameAsync(diffs[i].Pib);
            }

            ws.RangeUsed().SetAutoFilter();
            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
            MessageBox.Show("Uspešno sačuvano:\n" + path);
        }

        private string GetUniquePath(string basePath)
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

        private string Normalize(string input) =>
            Regex.Replace(input ?? "", @"[^\u0020-\u007E]", "")
                 .ToUpperInvariant();

        // Selenium <----------------------------------------------------------------------------------------------------------
        private async Task<string> GetPIB(string pib)
        {
            if (string.IsNullOrEmpty(pib))
                return "";  // no PIB → blank

            string nazivFirme = "";

            await Task.Run(() =>
            {
                lock (_driver)
                {
                    _driver.Navigate().GoToUrl("https://www.nbs.rs/rir_pn/rir.html.jsp");
                    var input = _driver.FindElement(By.Name("pib"));
                    input.Clear();
                    input.SendKeys(pib);
                    _driver.FindElement(By.CssSelector("input[type='submit'][value='Pretraži']")).Click();

                    var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
                    try
                    {
                        wait.Until(d =>
                            d.FindElements(By.Name("nazivULinku")).Count > 0
                        );
                        var hidden = _driver.FindElement(By.Name("nazivULinku"));
                        nazivFirme = hidden.GetAttribute("value")?.Trim('"') ?? "";
                    }
                    catch (WebDriverTimeoutException)
                    {

                    }
                }
            });

            return nazivFirme;
        }

        //private async void button4_Click(object sender, EventArgs e)
        //{
        //    string name = await NbsPibLookup.LookupNameAsync("101346140");
        //    button4.Text = string.IsNullOrEmpty(name)
        //        ? "Firma nije pronađena."
        //        : name;
        //}

        private record XlsRecord(string OriginalKey, double Value, string Marker, int Flag);
        private record CsvRecord(string OriginalKey, double SumValue, string Date1, string Date2, string Position, string Pib, int Flag);
        private record DiffRecord
        {
            public string Marker { get; set; }
            public string OriginalKey { get; init; }
            public double XlsValue { get; init; }
            public double CsvSumValue { get; init; }
            public string CsvOriginalKey { get; init; }
            public string CsvDate1 { get; init; }
            public string CsvDate2 { get; init; }
            public string Position { get; set; }
            public string Pib { get; set; }
        }
        public static class NbsPibLookup
        {
            // Single handler, but we disable auto-redirect (we parse the POST directly).
            private static readonly HttpClientHandler _handler = new HttpClientHandler
            {
                AllowAutoRedirect = false,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            private static readonly HttpClient _http = new HttpClient(_handler)
            {
                Timeout = TimeSpan.FromSeconds(10)
            };

            static NbsPibLookup()
            {
                // A realistic User-Agent so the server treats us like a browser:
                _http.DefaultRequestHeaders.UserAgent.ParseAdd(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
                    "AppleWebKit/537.36 (KHTML, like Gecko) " +
                    "Chrome/112.0.0.0 Safari/537.36"
                );
                // Accept headers for HTML
                _http.DefaultRequestHeaders.Accept.ParseAdd("text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
            }

            /// <summary>
            /// Looks up a company name by PIB in the NBS register.
            /// Will retry up to 3 times on transient network errors,
            /// and waits a short delay between calls to avoid hammering the server.
            /// </summary>
            public static async Task<string> LookupNameAsync(string pib)
            {
                if (string.IsNullOrWhiteSpace(pib))
                    return "";

                const string url = "https://www.nbs.rs/rir_pn/pn_rir.html.jsp?type=rir_results&lang=SER_CIR&konverzija=yes";
                var form = new Dictionary<string, string>
                {
                    ["pib"] = pib,
                    ["Submit"] = "Pretraži"
                };

                for (int attempt = 1; attempt <= 3; attempt++)
                {
                    try
                    {
                        Debug.WriteLine($"[DEBUG] Attempt #{attempt}: POST PIB={pib}");
                        using var content = new FormUrlEncodedContent(form);
                        var resp = await _http.PostAsync(url, content);
                        Debug.WriteLine($"[DEBUG] Status: {(int)resp.StatusCode} {resp.StatusCode}");

                        var html = await resp.Content.ReadAsStringAsync();
                        Debug.WriteLine($"[DEBUG] HTML snippet:\n{html.Substring(0, Math.Min(200, html.Length))}...\n---");

                        var doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(html);
                        var node = doc.DocumentNode.SelectSingleNode("//input[@name='nazivULinku']");
                        if (node != null)
                        {
                            var raw = node.GetAttributeValue("value", "");
                            Debug.WriteLine($"[DEBUG] Found nazivULinku = {raw}");
                            return raw.Trim('\"');
                        }
                        Debug.WriteLine("[DEBUG] No nazivULinku input found in response.");
                        return "";
                    }
                    catch (HttpRequestException ex) when (ex.InnerException is System.IO.IOException)
                    {
                        Debug.WriteLine($"[WARN] Network error on attempt {attempt}: {ex.Message}");
                        // exponential backoff: 500ms, 1000ms, 2000ms
                        await Task.Delay(500 * (1 << (attempt - 1)));
                    }
                    catch (TaskCanceledException ex) when (!ex.CancellationToken.IsCancellationRequested)
                    {
                        Debug.WriteLine($"[WARN] Timeout on attempt {attempt}: {ex.Message}");
                        await Task.Delay(500 * (1 << (attempt - 1)));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[ERROR] Unexpected error on attempt {attempt}: {ex}");
                        break;
                    }
                }

                Debug.WriteLine($"[ERROR] All attempts failed for PIB={pib}");
                return "";
            }
        }
    }
}
