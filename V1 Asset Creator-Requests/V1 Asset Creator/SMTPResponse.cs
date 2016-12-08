using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Web;
using OpenPop.Mime;
using System.Diagnostics;
using V1_Asset_Creator;

namespace VersionOne_Asset_Creator
{

    class SMTPResponse
    {
        static bool mailSent = false;

        public SMTPResponse()
        {

        }

        public void SendResponse(Message OriginalMail, string RequestInfo, string ProjectScope)
        {
            V1Logging Logs = new V1Logging();
            string MailBody = "Thanks for submitting your Request, the VersionOne Request ID is " + RequestInfo;
            V1XMLHelper V1XMLHelp = new V1XMLHelper();
            EmailSettings EmailSetting = new EmailSettings();
            string ToEmailAddress = V1XMLHelp.GetEmailAddressFromProjectScope(ProjectScope);
            //MailMessage mail = new MailMessage("versiononetesting@yahoo.com", "terry.densmore@versionone.com", "VersionOne Request Submission", MailBody);


            MailMessage mail = new MailMessage(ToEmailAddress, OriginalMail.Headers.From.ToString(), "VersionOne Request Submission", MailBody);

            EmailSetting = EmailSetting.GetEmailSettings(ToEmailAddress);


            //SmtpClient client = new SmtpClient("smtp.mail.yahoo.com");
            SmtpClient client = new SmtpClient(EmailSetting.SmtpServer);

            //client.Port = 587;
            client.Port = EmailSetting.SmtpPortNumber;

            //client.Credentials = new System.Net.NetworkCredential("versiononetesting@yahoo.com", "altec123");
            client.Credentials = new System.Net.NetworkCredential(ToEmailAddress, EmailSetting.Password);

            //client.EnableSsl = true;
            client.EnableSsl = EmailSetting.UseSSLSmtp;

            Debug.WriteLine("Sending Message To: " + OriginalMail.Headers.From.ToString() + "Request Info: " + MailBody);
            Logs.LogEvent("Operation - Sending Response Email to SMTP.");
            client.Send(mail);
        }
    }
}
