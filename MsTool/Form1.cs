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
        private string xlsMainPath; // Putanje fajlova za analiticke kartice
        private string xlsRefPath;
        private List<DiffRecord> diffs;

        bool analytics = false;

        //private readonly ChromeDriverService _driverService;
        //private readonly ChromeOptions _chromeOptions;
        //private readonly IWebDriver _driver;

        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        public Form1()
        {
            InitializeComponent();
            AllocConsole();
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
                if (!analytics)
                {
                    xlsPath = FileManipulator.ConvertXlsToXlsx(ofd.FileName);
                    label1.Text = Path.GetFileName(xlsPath);
                }
                else
                {
                    xlsRefPath = FileManipulator.ConvertXlsToXlsx(ofd.FileName);
                    label1.Text = Path.GetFileName(xlsRefPath);

                    Console.WriteLine("Analytics mode");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!analytics)
            {
                var ofd = new OpenFileDialog { Filter = "CSV File|*.csv" };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    csvPath = ofd.FileName;
                    label2.Text = Path.GetFileName(csvPath);
                }
            }
            else
            {
                var ofd = new OpenFileDialog { Filter = "Excel 97-2003 Workbook|*.xls" };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    xlsMainPath = FileManipulator.ConvertXlsToXlsx(ofd.FileName);
                    label2.Text = Path.GetFileName(xlsMainPath);
                }

                Console.WriteLine("Analytics mode");
            }
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            if (!analytics)
            {
                BookOfBillsFunctions.Proceed(xlsPath, csvPath, checkBox1.Checked, AssumptionsCB.Checked);
            }
            else
            {
                AnalyticsFunctions.Proceed(xlsMainPath, xlsRefPath, AssumptionsCB.Checked);

                Console.WriteLine("Analytics functions here");
            }
        }

        private void AnalyticsCB_CheckedChanged(object sender, EventArgs e)
        {
            analytics = !analytics;
            if (analytics)
            {
                button1.Text = "Referentni fajl";
                button2.Text = "Moj fajl";

                checkBox1.Enabled = false;
            }
            else
            {
                button1.Text = "Moj fajl";
                button2.Text = "Fajl poreske uprave";

                checkBox1.Enabled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PibStore.Instance.Delete("100000821");
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
