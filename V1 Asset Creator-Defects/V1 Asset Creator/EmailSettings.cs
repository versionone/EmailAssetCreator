using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace VersionOne_Asset_Creator
{
    class EmailSettings
    {
        public string Pop3Server { get; set; }
        public int Pop3PortNumber { get; set; }
        public string EmailAddress { get; set; }
        public string Password { get; set; }
        public Boolean UseSSLPop3 { get; set; }
        public string SmtpServer { get; set; }
        public int SmtpPortNumber { get; set; }
        public Boolean UseSSLSmtp { get; set; }

        public EmailSettings()
        {

        }

        public EmailSettings GetEmailSettings(string EmailAddress)
        {
            EmailSettings Settings = new EmailSettings();
            XElement root = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + "V1EmailConfig.xml");

            var EmailInfo = from cli in root.Elements("EmailAccount") where cli.Element("EmailAddress").Value == EmailAddress select cli;

            if (EmailInfo.Any())
            {
                Settings.Password = EmailInfo.Elements("EmailPassword").First().Value;
                Settings.Pop3Server = EmailInfo.Elements("EmailPOP3").First().Value;
                Settings.Pop3PortNumber = Convert.ToInt32(EmailInfo.Elements("POP3PortNumber").First().Value);
                Settings.UseSSLPop3 = Convert.ToBoolean(EmailInfo.Elements("EmailUseSSL").First().Value);
                Settings.SmtpServer = EmailInfo.Elements("EmailSmtpServer").First().Value;
                Settings.SmtpPortNumber = Convert.ToInt32(EmailInfo.Elements("EmailSmtpPort").First().Value);
                Settings.UseSSLSmtp = Convert.ToBoolean(EmailInfo.Elements("EmailSmtpUseSSL").First().Value);
            }


            return Settings;
        }

    }
}
