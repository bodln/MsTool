//using OpenQA.Selenium;
//using OpenQA.Selenium.Chrome;
//using OpenQA.Selenium.Support.UI;
using MsTool.Models;
using MsTool.Utlis;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;


namespace MsTool
{
    public partial class Form1 : Form
    {
        private string xlsPath;
        private string csvPath;
        private List<DiffRecord> diffs;

        //private readonly ChromeDriverService _driverService;
        //private readonly ChromeOptions _chromeOptions;
        //private readonly IWebDriver _driver;

        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        public Form1()
        {
            InitializeComponent();
            //AllocConsole();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            //// Setup selenium once and use it for all searches
            //_driverService = ChromeDriverService.CreateDefaultService();
            //_driverService.HideCommandPromptWindow = true;
            //_driverService.SuppressInitialDiagnosticInformation = true;

            //_chromeOptions = new ChromeOptions();
            //_chromeOptions.AddArgument("--headless");
            //_chromeOptions.AddArgument("--disable-gpu");
            //_chromeOptions.AddArgument("--disable-extensions");
            //_chromeOptions.AddArgument("--disable-popup-blocking");
            //_chromeOptions.AddArgument("--log-level=3");
            //// Does not wait for images and such to load 
            //_chromeOptions.PageLoadStrategy = PageLoadStrategy.Eager;

            //_driver = new ChromeDriver(_driverService, _chromeOptions);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            //_driver?.Quit();
            //_driver?.Dispose();

            try
            {
                if (!string.IsNullOrWhiteSpace(xlsPath) && File.Exists(xlsPath))
                {
                    File.Delete(xlsPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Greška pri brisanju privremenog fajla: {ex.Message}");
            }

            base.OnFormClosing(e);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog { Filter = "Excel 97-2003 Workbook|*.xls" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                xlsPath = FileManipulator.ConvertXlsToXlsx(ofd.FileName);
                label1.Text = Path.GetFileName(xlsPath);
            }
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

        private async void button3_Click(object sender, EventArgs e)
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

                diffs = new List<DiffRecord>();

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
                            DoubleTake = doubleTake
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

                SaveDialog.ShowSaveDialog(diffs, checkBox1.Checked, AssumptionsCB.Checked);

                //IReadOnlyDictionary<string, string> db = PibStore.Instance.GetAll();
                //db.ToList().ForEach(kvp => Console.WriteLine($"{kvp.Key} -> {kvp.Value}"));

                //if (File.Exists(xlsPath))
                //    File.Delete(xlsPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Greška Form1.cs: " + ex.Message);
            }
        }

        // Selenium <----------------------------------------------------------------------------------------------------------
        //private async Task<string> GetPIB(string pib)
        //{
        //    if (string.IsNullOrEmpty(pib))
        //        return "";

        //    string nazivFirme = "";

        //    await Task.Run(() =>
        //    {
        //        lock (_driver)
        //        {
        //            _driver.Navigate().GoToUrl("https://www.nbs.rs/rir_pn/rir.html.jsp");
        //            var input = _driver.FindElement(By.Name("pib"));
        //            input.Clear();
        //            input.SendKeys(pib);
        //            _driver.FindElement(By.CssSelector("input[type='submit'][value='Pretraži']")).Click();

        //            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
        //            try
        //            {
        //                wait.Until(d =>
        //                    d.FindElements(By.Name("nazivULinku")).Count > 0
        //                );
        //                var hidden = _driver.FindElement(By.Name("nazivULinku"));
        //                nazivFirme = hidden.GetAttribute("value")?.Trim('"') ?? "";
        //            }
        //            catch (WebDriverTimeoutException)
        //            {

        //            }
        //        }
        //    });

        //    return nazivFirme;
        //}

        //private async void button4_Click(object sender, EventArgs e)
        //{
        //    string name = await NbsPibLookup.LookupNameAsync("101346140");
        //    button4.Text = string.IsNullOrEmpty(name)
        //        ? "Firma nije pronađena."
        //        : name;
        //}
    }
}
