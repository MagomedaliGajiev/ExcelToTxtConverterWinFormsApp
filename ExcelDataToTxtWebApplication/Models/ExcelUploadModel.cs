namespace ExcelDataToTxtWebApplication.Models
{
    public class ExcelUploadModel
    {
        public HttpPostedFileBase ExcelFile { get; set; }
        public string Message { get; set; }
        public List<string> GeneratedFiles { get; set; } = new List<string>();
    }
}
