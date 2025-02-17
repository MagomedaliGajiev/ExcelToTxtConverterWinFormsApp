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
        public List<string> GeneratedFiles { get; set; } = new List<string>();
    }
}