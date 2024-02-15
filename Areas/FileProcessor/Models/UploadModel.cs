namespace ExcelProcessor.Areas.FileProcessor.Models
{
    public class UploadModel
    {
        public IFormFile ExcelFile { get; set; }
        public string TextContent { get; set; }
    }
} 