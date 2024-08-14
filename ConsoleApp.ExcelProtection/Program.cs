using ConsoleApp.ExcelProtection.Services;
using OfficeOpenXml;
using Serilog;
using System;

namespace ConsoleApp.ExcelProtection
{
    class Program
    {
        static void Main(string[] args)
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .WriteTo.File("logs/ExcelProtection.txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();

            Console.WriteLine("Starting Excel Protection Application...");

            //string masterFilePath = @"D:\SuThetK\Documentation\master.xlsx";
            Console.Write("Please enter the path of the master Excel file: ");
            string masterFilePath = Console.ReadLine();           

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelProtectionService excelProtectionService = new ExcelProtectionService();
            excelProtectionService.ProcessMasterFile(masterFilePath);

            Console.WriteLine("Password protection applied to relevant files.");
            Console.WriteLine("\nPress the Enter key to exit the application...\n");
            Console.ReadLine();

            excelProtectionService.Dispose();
        }
    }
}
