using OfficeOpenXml;
using Serilog;
using System;
using System.IO;

namespace ConsoleApp.ExcelProtection.Services
{
    public class ExcelProtectionService : IDisposable
    {
        public void ProcessMasterFile(string masterFilePath)
        {
            if (File.Exists(masterFilePath))
            {
                using (var package = new ExcelPackage(new FileInfo(masterFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++) // Assuming first row is header
                    {
                        int staffId = int.Parse(worksheet.Cells[row, 1].Value?.ToString().Trim());
                        string filePath = worksheet.Cells[row, 3].Value?.ToString().Trim();
                        string password = worksheet.Cells[row, 4].Value?.ToString().Trim();

                        if (CheckStaffIdInFilePath(staffId, filePath))
                        {
                            SetExcelPassword(filePath, password);
                        }
                    }
                }
            }
        }

        private bool CheckStaffIdInFilePath(int staffId, string filePath)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            string staffIdStr = staffId.ToString();
            return fileNameWithoutExtension.EndsWith(staffIdStr);
        }

        private void SetExcelPassword(string filePath, string password)
        {
            if (File.Exists(filePath))
            {
                FileInfo fileInfo = new FileInfo(filePath);
                try
                {
                    using (var package = new ExcelPackage(fileInfo))
                    {
                        if (package.Encryption.IsEncrypted)
                        {
                            LogMessage($"File is already protected: {filePath}");
                            return; // Exit as the file is already protected
                        }

                        package.Encryption.IsEncrypted = true;
                        package.Encryption.Algorithm = OfficeOpenXml.EncryptionAlgorithm.AES256; // AES256 is more secure
                        package.Encryption.Password = password;

                        package.Save();
                    }
                    LogMessage($"Password set for file: {filePath}");
                }
                catch (InvalidDataException ex)
                {
                    LogMessage($"The file format is not supported or the file is already encrypted: {filePath}. Error: {ex.Message}");
                }
                catch (Exception ex)
                {
                    LogMessage($"Failed to set password and encrypt file: {filePath}. Error: {ex.Message}");
                }
            }
            else
            {
                LogMessage($"File not found: {filePath}");
            }
        }

        private void LogMessage(string message)
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string logFilePath = Path.Combine(path, $"ProtectionService_{DateTime.Now:yyyy-MM-dd}.txt");

            try
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine($"{DateTime.Now}: {message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Logging Error: {ex.Message}");
            }
        }

        public void Dispose()
        {
            // Dispose resources if needed
        }
    }
}
