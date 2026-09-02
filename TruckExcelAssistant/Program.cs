namespace TruckExcelAssistant;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Contains("--excel-smoke-test", StringComparer.OrdinalIgnoreCase))
        {
            Services.ExcelExportSmokeTest.Run();
            return;
        }

        ApplicationConfiguration.Initialize();
        Application.SetDefaultFont(new Font("Segoe UI", 9F));
        try
        {
            var database = new Services.DatabaseService();
            database.Initialize();
            Application.Run(new MainForm(database));
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Truck Excel Assistant tidak dapat membuka database lokal.\n\n{ex.Message}",
                "Database tidak tersedia",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
