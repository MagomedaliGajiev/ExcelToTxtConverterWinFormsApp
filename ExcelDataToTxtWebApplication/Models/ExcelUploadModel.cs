using System.ComponentModel.DataAnnotations;

namespace ExcelToTxtWebApp.Models
{
    public class ExcelUploadModel
    {
        [Required(ErrorMessage = "Выберите файл Excel")]
        [DataType(DataType.Upload)]
        [FileExtensions(Extensions = "xlsx", ErrorMessage = "Только Excel файлы")]
        public IFormFile ExcelFile { get; set; }

        public string Message { get; set; }
        public List<GeneratedFile> GeneratedFiles { get; set; } = new();
        [Display(Name = "Серверные пути сохранения")]
        public List<string> ServerPaths { get; set; } = new();
    }

    public class GeneratedFile
    {
        public string FileName { get; set; }
        public string Content { get; set; }
    }
}