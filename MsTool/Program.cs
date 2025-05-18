using MsTool.Utlis;

namespace MsTool
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            // Touch Sqlite db fo initialization
            var _ = PibStore.Instance;
            Application.Run(new Form1());
        }
    }
}