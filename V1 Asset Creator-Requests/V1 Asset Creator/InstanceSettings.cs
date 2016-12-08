using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace VersionOne_Asset_Creator
{
    class InstanceSettings
    {
        public string   InstanceURL { get; set; }
        public string   AccessToken { get; set; }
        public string ProxyURL { get; set; }
        public string ProxyUserName { get; set; }
        public string ProxyPassword { get; set; }
        public int PollingInterval { get; set; }

        public InstanceSettings()
        {

        }

        public InstanceSettings GetInstanceSettings()
        {
            InstanceSettings Settings = new InstanceSettings();
            XElement root = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + "V1InstanceConfig.xml");

            var ServerInfo = from cli in root.Elements("V1Server") select cli;

            if (ServerInfo.Any())
            {
                Settings.InstanceURL = ServerInfo.Elements("InstanceURL").First().Value;
                Settings.AccessToken = ServerInfo.Elements("AccessToken").First().Value;
                Settings.ProxyURL = ServerInfo.Elements("ProxyURL").First().Value;
                Settings.ProxyUserName = ServerInfo.Elements("ProxyUserName").First().Value;
                Settings.ProxyPassword = ServerInfo.Elements("ProxyPassword").First().Value;
                Settings.PollingInterval = int.Parse(ServerInfo.Elements("IntervalInMinutes").First().Value);
            }


            return Settings;
        }
    }
}
