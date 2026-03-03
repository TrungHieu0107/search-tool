using OfficeOpenXml;

namespace ExcelSearchTool;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        ApplicationConfiguration.Initialize();
        Application.Run(new UI.MainForm());
    }
}
