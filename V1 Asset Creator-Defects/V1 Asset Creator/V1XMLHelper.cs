using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace VersionOne_Asset_Creator
{
    class V1XMLHelper
    {

        public V1XMLHelper()
        {

        }

        public string GetProjectScope(string EmailAddress)
        {
            string ProjectScope = null;
            XElement root = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + "V1EmailConfig.xml");
            //IEnumerable<XElement> projectoid = from el in root.Elements("ProjectOID") where (string)el.Element("EmailAddress") == EmailAddress select el.Element("ProjectOID");

            //XElement[] ElementInfo = projectoid.ToArray();

            var projectinfo = from cli in root.Elements("EmailAccount") where cli.Element("EmailAddress").Value == EmailAddress select cli;

            /* This works!
            foreach(var d in projectinfo)
            {
                ProjectScope = d.Element("ProjectScope").Value;
            }*/

            //var blah = XmlSerilizer.Deserialize<List<EmailAccount>>(file)

            // So does this
            if (projectinfo.Any())
                ProjectScope = projectinfo.Elements("ProjectScope").First().Value;


            return ProjectScope;

        }

        public string GetEmailAddressFromProjectScope(string ProjectScope)
        {
            string EmailAddress = null;
            XElement root = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + "V1EmailConfig.xml");


            var EmailInfo = from cli in root.Elements("EmailAccount") where cli.Element("ProjectScope").Value == ProjectScope select cli;


            if (EmailInfo.Any())
                EmailAddress = EmailInfo.Elements("EmailAddress").First().Value;


            return EmailAddress;

        }
    }
}
