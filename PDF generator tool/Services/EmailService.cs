using System.Net;
using System.Net.Mail;
using System.IO;
using System.Threading.Tasks;

namespace YourNamespace.Services
{
    public class EmailService
    {
        private readonly string _smtpUser = "darjikrishan12@gmail.com"; // Replace with your Gmail
        private readonly string _smtpPass = "psdsdixfiaczuima";     // Replace with your 16-char App Password (no spaces)

        public async Task<bool> SendEmailAsync(string toEmail, string subject, string body, byte[] attachmentBytes, string fileName)
        {
            try
            {
                var message = new MailMessage
                {
                    From = new MailAddress(_smtpUser),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = false
                };

                message.To.Add(toEmail);

                if (attachmentBytes != null && attachmentBytes.Length > 0)
                {
                    var stream = new MemoryStream(attachmentBytes);
                    var attachment = new Attachment(stream, fileName, "application/pdf");
                    message.Attachments.Add(attachment);
                }

                using (var client = new SmtpClient("smtp.gmail.com", 587))
                {
                    client.Credentials = new NetworkCredential(_smtpUser, _smtpPass);
                    client.EnableSsl = true;

                    await client.SendMailAsync(message);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
