using ExcelDataToTxtWebApplication.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Text;

namespace ExcelDataToTxtWebApplication.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View(new ExcelUploadModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Index(ExcelUploadModel model)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var config = new
            {
                DateCellAddress = "G2",
                MainPathCell = "I6",
                BackupPathCell = "I7",
                StartRow = 3,
                EndRow = 40,
                Columns = new[] { 3, 4, 5, 6, 7 }
            };

            try
            {
                if (model.ExcelFile?.Length > 0)
                {
                    // Копируем поток в MemoryStream
                    using var memoryStream = new MemoryStream();
                    model.ExcelFile.CopyTo(memoryStream);
                    memoryStream.Position = 0;

                    // Используем MemoryStream для ExcelPackage
                    using var package = new ExcelPackage(memoryStream, "38214");
                    var worksheet = package.Workbook.Worksheets[0];

                    // Parse date
                    var rawDate = worksheet.Cells[config.DateCellAddress].Text;
                    if (!DateTime.TryParseExact(rawDate, "dd.MM.yy", null,
                        System.Globalization.DateTimeStyles.None, out DateTime fileDate))
                    {
                        throw new FormatException("Неверный формат даты в ячейке G2");
                    }

                    // Get paths
                    var mainDir = @"C:\ASSPOOTI";
                    var backupDir = @"E:\BALLANS";
                    Directory.CreateDirectory(mainDir);
                    Directory.CreateDirectory(backupDir);

                    // Generate content
                    var sb = new StringBuilder();
                    for (int row = config.StartRow; row <= config.EndRow; row++)
                    {
                        var values = config.Columns
                            .Select(col => worksheet.Cells[row, col].Text?.Trim() ?? "")
                            .ToList();
                        sb.AppendLine(string.Join(";", values));
                    }

                    // Save files
                    var fileName = $"KTMS0L00_{fileDate:yyyyMMdd}.txt";
                    var mainPath = Path.Combine(mainDir, fileName);
                    var backupPath = Path.Combine(backupDir, fileName);

                    System.IO.File.WriteAllText(mainPath, sb.ToString(), Encoding.UTF8);
                    System.IO.File.WriteAllText(backupPath, sb.ToString(), Encoding.UTF8);

                    model.GeneratedFiles.Add(fileName);
                    model.Message = "Файлы успешно сгенерированы!";
                }
                else
                {
                    model.Message = "Пожалуйста, выберите файл";
                }
            }
            catch (Exception ex)
            {
                model.Message = $"Ошибка: {ex.Message}";
            }

            return View(model);
        }
    }
}