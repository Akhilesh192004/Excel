namespace Excel.ViewModels
{
    public class ExcelUploadViewModel
    {
        public IFormFile? ExcelFile { get; set; }
        public int ClientId { get; set; } = 1;
        public int CollRequestId { get; set; } = 2;
        public int DocCategory { get; set; } = 1;
        public int DocCategory1 { get; set; } = 1;
        public int DocCategory2 { get; set; } = 2;
        public int UserID { get; set; }
        public long? UploadId { get; set; }
        public string SuccessMessage { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
    }
}
