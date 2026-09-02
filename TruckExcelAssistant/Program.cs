namespace TruckExcelAssistant;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.SetDefaultFont(new Font("Segoe UI", 9F));
        Application.Run(new MainForm());
    }
}
