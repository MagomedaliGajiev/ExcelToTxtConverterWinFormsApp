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

                model.Message = "Файлы готовы к сохранению!";
                return View(model);
            }
            catch (Exception ex)
            {
                model.Message = $"Ошибка: {ex.Message}";
                return View(model);
            }
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