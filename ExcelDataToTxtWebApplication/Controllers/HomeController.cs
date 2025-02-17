using ExcelToTxtWebApp.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Text;

namespace ExcelToTxtWebApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _env;

        public HomeController(IWebHostEnvironment env)
        {
            _env = env;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public IActionResult Index()
        {
            return View(new ExcelUploadModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Index(ExcelUploadModel model)
        {
            try
            {
                if (model?.ExcelFile == null || model.ExcelFile.Length == 0)
                {
                    model.Message = "Пожалуйста, выберите файл";
                    return View(model);
                }

                var config = new
                {
                    DateCellAddress = "G2",
                    MainPathCell = "I6",
                    BackupPathCell = "I7",
                    StartRow = 3,
                    EndRow = 40,
                    Columns = new[] { 3, 4, 5, 6, 7 },
                    SavePaths = new
                    {
                        Main = Path.Combine(_env.WebRootPath, "generated"),
                        Backup = Path.Combine(_env.WebRootPath, "backup")
                    }
                };

                using var memoryStream = new MemoryStream();
                await model.ExcelFile.CopyToAsync(memoryStream);
                memoryStream.Position = 0;

                using var package = new ExcelPackage(memoryStream, "38214");
                var worksheet = package.Workbook.Worksheets[0];

                // Парсинг данных
                var rawDate = worksheet.Cells[config.DateCellAddress].Text;
                if (!DateTime.TryParseExact(rawDate, "dd.MM.yy", null,
                    System.Globalization.DateTimeStyles.None, out DateTime fileDate))
                {
                    throw new FormatException("Неверный формат даты в ячейке G2");
                }

                // Генерация содержимого
                var sb = new StringBuilder();
                for (int row = config.StartRow; row <= config.EndRow; row++)
                {
                    var values = config.Columns
                        .Select(col => worksheet.Cells[row, col].Text?.Trim() ?? "")
                        .ToList();
                    sb.AppendLine(string.Join(";", values));
                }

                // Сохранение файлов
                var fileName = $"KTMS0L00_{fileDate:yyyyMMdd}.txt";
                Directory.CreateDirectory(config.SavePaths.Main);
                Directory.CreateDirectory(config.SavePaths.Backup);

                var mainPath = Path.Combine(config.SavePaths.Main, fileName);
                var backupPath = Path.Combine(config.SavePaths.Backup, fileName);

                await System.IO.File.WriteAllTextAsync(mainPath, sb.ToString(), Encoding.UTF8);
                await System.IO.File.WriteAllTextAsync(backupPath, sb.ToString(), Encoding.UTF8);

                model.GeneratedFiles.Add($"/generated/{fileName}");
                model.GeneratedFiles.Add($"/backup/{fileName}");
                model.Message = "Файлы успешно сгенерированы!";
            }
            catch (Exception ex)
            {
                model.Message = $"Ошибка: {ex.Message}";
            }

            return View(model);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View();
        }
    }
}