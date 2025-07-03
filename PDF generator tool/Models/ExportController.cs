//using Microsoft.AspNetCore.Mvc;
//using System.Threading.Tasks;
//using YourNamespace.Models.Request;
//using YourNamespace.Services;

//[Route("api/[controller]")]
//[ApiController]
//public class ExportController : ControllerBase
//{
//    private readonly EmailService _emailService;

//    public ExportController(EmailService emailService)
//    {
//        _emailService = emailService;
//    }

//    [HttpPost("send-email")]
//    public async Task<IActionResult> SendEmail([FromBody] EmailRequest model)
//    {
//        if (string.IsNullOrWhiteSpace(model.Email) || model.PdfBytes == null || model.PdfBytes.Length == 0)
//        {
//            return BadRequest(new { message = "Invalid request. Email or PDF is missing." });
//        }

//        var result = await _emailService.SendEmailAsync(
//            model.Email,
//            "Your PDF Report",
//            "Please find your exported PDF report attached.",
//            model.PdfBytes,
//            model.FileName ?? "Report.pdf"
//        );

//        if (result)
//            return Ok(new { success = true, message = "Email sent successfully!" });

//        return StatusCode(500, new { success = false, message = "Failed to send email." });
//    }
//}
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using YourNamespace.Services;
using System;

[Route("api/[controller]")]
[ApiController]
public class ExportController : ControllerBase
{
    private readonly EmailService _emailService;

    public ExportController(EmailService emailService)
    {
        _emailService = emailService;
    }

    [HttpPost("send-email")]
    public async Task<IActionResult> SendEmail([FromBody] EmailRequest model)
    {
        if (string.IsNullOrWhiteSpace(model.Email) || string.IsNullOrWhiteSpace(model.PdfBase64))
        {
            return BadRequest(new { message = "Invalid request. Email or PDF data is missing." });
        }

        try
        {
            byte[] pdfBytes = Convert.FromBase64String(model.PdfBase64);

            var result = await _emailService.SendEmailAsync(
                model.Email,
                "Your PDF Report",
                "Please find your exported PDF report attached.",
                pdfBytes,
                model.FileName ?? "Report.pdf"
            );

            if (result)
                return Ok(new { success = true, message = "Email sent successfully!" });

            return StatusCode(500, new { success = false, message = "Failed to send email." });
        }
        catch
        {
            return BadRequest(new { message = "Invalid base64 PDF data." });
        }
    }
}
