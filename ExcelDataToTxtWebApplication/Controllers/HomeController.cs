using ExcelToTxtWebApp.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Security;
using System.Text;

namespace ExcelToTxtWebApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _env;
        private readonly IConfiguration _config;

        public HomeController(IWebHostEnvironment env, IConfiguration config)
        {
            _env = env;
            _config = config;
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
                    throw new Exception("Пожалуйста, выберите файл");

                var config = new
                {
                    DateCellAddress = "G2",
                    StartRow = 3,
                    EndRow = 40,
                    Columns = new[] { 3, 4, 5, 6, 7 }
                };

                using var stream = new MemoryStream();
                await model.ExcelFile.CopyToAsync(stream);

                using var package = new ExcelPackage(stream, "38214");
                var worksheet = package.Workbook.Worksheets[0];

               // получения путей из Excel
                var serverDirs = new List<string>
                {
                    worksheet.Cells["I6"].Text.Trim(),
                    worksheet.Cells["I7"].Text.Trim()
                };
                // Парсинг даты
                var rawDate = worksheet.Cells[config.DateCellAddress].Text;
                if (!DateTime.TryParseExact(rawDate, "dd.MM.yy", null,
                    System.Globalization.DateTimeStyles.None, out DateTime fileDate))
                    throw new FormatException("Неверный формат даты в ячейке G2");

                // Генерация содержимого
                var sb = new StringBuilder();
                for (int row = config.StartRow; row <= config.EndRow; row++)
                {
                    var values = config.Columns
                        .Select(col => worksheet.Cells[row, col].Text?.Trim() ?? "")
                        .ToList();
                    sb.AppendLine(string.Join(";", values));
                }

                var fileName = $"KTMS0L00_{fileDate:yyyyMMdd}.txt";
                model.GeneratedFiles.Add(new GeneratedFile
                {
                    FileName = fileName,
                    Content = sb.ToString()
                });

                // Сохранение на сервере
                foreach (var dir in serverDirs)
                {
                    ValidateServerPath(dir); // Валидация пути

                    Directory.CreateDirectory(dir);
                    var fullPath = Path.Combine(dir, fileName);
                    await System.IO.File.WriteAllTextAsync(fullPath, sb.ToString());
                }

                model.Message = "Файлы готовы к сохранению!";
                return View(model);
            }
            catch (Exception ex)
            {
                model.Message = $"Ошибка: {ex.Message}";
                return View(model);
            }
        }

        private void ValidateServerPath(string path)
        {
            var allowedPaths = _config.GetSection("StorageSettings:AllowedBasePaths")
                .Get<List<string>>();

            var isValid = allowedPaths.Any(allowed =>
                Path.GetFullPath(path).StartsWith(Path.GetFullPath(allowed)));

            if (!isValid)
                throw new SecurityException($"Недопустимый путь для сохранения: {path}");
        }

        [HttpPost]
        public IActionResult GetFileContents([FromBody] FileRequest request)
        {
            // В реальном приложении здесь должна быть проверка авторизации
            var files = new List<GeneratedFile>
            {
                new() { FileName = request.FileName, Content = request.Content }
            };
            return Json(files);
        }

        public IActionResult Privacy() => View();

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error() => View();
    }



    public class FileRequest
    {
        public string FileName { get; set; }
        public string Content { get; set; }
    }
}