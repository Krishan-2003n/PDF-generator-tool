//namespace YourNamespace.Models.Request
//{
//    public class EmailRequest
//    {
//        public string Email { get; set; }              // recipient email
//        public byte[] PdfBytes { get; set; }           // PDF as byte[]
//        public string FileName { get; set; } = "Report.pdf"; // optional, default to Report.pdf
//    }
//}
using System.Text.Json.Serialization;

public class EmailRequest
{
    public string Email { get; set; }
    public string PdfBase64 { get; set; }
    public string FileName { get; set; }
}

