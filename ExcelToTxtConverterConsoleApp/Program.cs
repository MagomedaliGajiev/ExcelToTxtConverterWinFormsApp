using OfficeOpenXml;
using System.Text;

public class ExcelToTxtConverter
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Конфигурация
        Console.Write("Укажите путь к файлу:");
        var sourceExcelPath = Console.ReadLine(); // Укажите путь к исходному файлу
        var dateCellAddress = "G2";
        var mainPathCell = "I6";
        var backupPathCell = "I7";
        var startRow = 3;
        var endRow = 40;
        var columns = new int[]{ 3, 4, 5, 6, 7 }; // Колонки C-G

        try
        {
            // Чтение данных из Excel
            using (var package = new ExcelPackage(new FileInfo(sourceExcelPath), "38214"))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Первый лист

                // Получение значений из ячеек
                var rawDate = worksheet.Cells[dateCellAddress].Text;
                var mainDir = worksheet.Cells[mainPathCell].Text;
                var backupDir = worksheet.Cells[backupPathCell].Text;

                // Парсинг даты
                if (!DateTime.TryParseExact(rawDate, "dd.MM.yy", null, System.Globalization.DateTimeStyles.None, out DateTime fileDate))
                    throw new FormatException("Неверный формат даты в ячейке G2");

                // Формирование имени файла
                var fileName = $"KTMS0L00_{fileDate:yyyyMMdd}.txt";

                // Генерация данных
                var sb = new StringBuilder();
                for (int row = startRow; row <= endRow; row++)
                {
                    var values = new List<string>();
                    foreach (int col in columns)
                    {
                        values.Add(worksheet.Cells[row, col].Text?.Trim() ?? "");
                    }
                    sb.AppendLine(string.Join(";", values));
                }

                // Сохранение файлов
                SaveToFile(mainDir, fileName, sb.ToString());
                SaveToFile(backupDir, fileName, sb.ToString());

                Console.WriteLine($"Файлы успешно созданы: {Path.Combine(mainDir, fileName)} и {Path.Combine(backupDir, fileName)}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка: {ex.Message}");
        }
    }

    private static void SaveToFile(string directory, string fileName, string content)
    {
        // Создание директории, если не существует
        Directory.CreateDirectory(directory);

        // Полный путь к файлу
        var fullPath = Path.Combine(directory, fileName);

        // Запись содержимого
        File.WriteAllText(fullPath, content, Encoding.UTF8);
    }
}