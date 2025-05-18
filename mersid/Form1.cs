using ClosedXML.Excel;
using ExcelDataReader;
//using OpenQA.Selenium;
//using OpenQA.Selenium.Chrome;
//using OpenQA.Selenium.Support.UI;
using System.Data;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using mersid.Models;
using mersid;


namespace mersid
{
    public partial class Form1 : Form
    {
        private string xlsPath;
        private string csvPath;
        private List<DiffRecord> diffs;

        // Selenium driver fields
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

            // ─── SETUP SELENIUM ONCE ───────────────────────────────────────────
            //_driverService = ChromeDriverService.CreateDefaultService();
            //_driverService.HideCommandPromptWindow = true;
            //_driverService.SuppressInitialDiagnosticInformation = true;

            //_chromeOptions = new ChromeOptions();
            //_chromeOptions.AddArgument("--headless");
            //_chromeOptions.AddArgument("--disable-gpu");
            //_chromeOptions.AddArgument("--disable-extensions");
            //_chromeOptions.AddArgument("--disable-popup-blocking");
            //_chromeOptions.AddArgument("--log-level=3");
            //_chromeOptions.PageLoadStrategy = PageLoadStrategy.Eager;

            //_driver = new ChromeDriver(_driverService, _chromeOptions);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            //_driver?.Quit();
            //_driver?.Dispose();
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

        private void button3_Click(object sender, EventArgs e)
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

                    if (xls == null || !ValuesEqual(xVal, cSum))
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

                SaveDialog.ShowSaveDialog(diffs);

                if (File.Exists(xlsPath))
                    File.Delete(xlsPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Greška: " + ex.Message);
            }
        }

        private bool ValuesEqual(double a, double b) =>
            Math.Abs(a - b) <= 4.0;

        // Selenium <----------------------------------------------------------------------------------------------------------
        //private async Task<string> GetPIB(string pib)
        //{
        //    if (string.IsNullOrEmpty(pib))
        //        return "";  // no PIB → blank

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
