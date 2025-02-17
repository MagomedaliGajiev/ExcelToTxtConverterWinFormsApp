using System.ComponentModel.DataAnnotations;

namespace ExcelDataToTxtWebApplication.Models
{
    public class ExcelUploadModel
    {
        [Required(ErrorMessage = "Выберите файл Excel")]
        public IFormFile ExcelFile { get; set; }
        public string Message { get; set; }
        public List<string> GeneratedFiles { get; set; } = new List<string>();
    }
}