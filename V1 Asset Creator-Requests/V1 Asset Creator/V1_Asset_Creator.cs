using OpenPop.Mime;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using VersionOne.SDK.APIClient;
using System.Timers;
using VersionOne_Asset_Creator;


namespace V1_Asset_Creator
{
    public partial class V1_Asset_Creator : ServiceBase
    {
        //private static System.Timers.Timer ServiceTimer;
        private static EventLog eventLog1 = new EventLog();

        public V1_Asset_Creator()
        {
            InitializeComponent();
            // Turn off autologging 
            this.AutoLog = false;
            // create an event source, specifying the name of a log that 
            // does not currently exist to create a new, custom log 
            if (!System.Diagnostics.EventLog.SourceExists("V1 Request Asset Creator"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "V1 Request Asset Creator", "V1 Request Asset Creator Log");
            }
            // configure the event log instance to use this source name
            eventLog1.Source = "V1 Request Asset Creator";
        }

        public void DuringDebug()
        {
            OnStart(null);
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                //System.Diagnostics.Debugger.Launch();

                InstanceSettings InstanceInfo = new InstanceSettings();
                InstanceInfo = InstanceInfo.GetInstanceSettings();

                System.Timers.Timer ServiceTimer = new Timer((InstanceInfo.PollingInterval*60000)/*1800000 300000*/);
                ServiceTimer.AutoReset = true;
                ServiceTimer.Elapsed += CheckMail;
                ServiceTimer.Enabled = true; //ServiceTimer.Start();//ServiceTimer.Enabled = true;
                ServiceTimer.Start();

                
                GC.KeepAlive(ServiceTimer);

                // write an entry to the log
                eventLog1.WriteEntry("Operation - V1 Request Asset Creator Service Started");
            }
            catch (Exception e)
            {
               eventLog1.WriteEntry("ERROR - OnStart - " + e.InnerException.Message);
               
            }

           
        }

        protected override void OnStop()
        {
            // write an entry to the log
            eventLog1.WriteEntry("Operation - V1 Request Asset Creator Service Stopped");
        }

        private static void CheckMail(Object source, ElapsedEventArgs e)
        {
            List<Message> Messages = new List<Message>();
            V1Logging Logs = new V1Logging();
            bool V1TestResult = false;
            V1_Connector V1Connection = new V1_Connector();
            V1Connector V1Instance = V1Connection.ConnectToV1();
            MailChecker GetMail = new MailChecker();

            //System.IO.File.Create(AppDomain.CurrentDomain.BaseDirectory + "MailChecker.txt");

            //Messages = GetMail.FetchAllMessages("pop.mail.yahoo.com", 995, true, "versiononetesting@yahoo.com", "altec123");
            Messages = GetMail.FetchAllMessagesWithMultipleAccounts();
            Console.Write(Messages.ToString());

            // write an entry to the log
            eventLog1.WriteEntry("Operation - Email Checked, " + Messages.Count.ToString() + " Found");

            if (Messages.Count != 0)
                V1TestResult = V1Connection.V1TestConnection(V1Instance);

            if (Messages.Count != 0 && V1TestResult == true)
            {
                //NewConnection = V1Connection.ConnectToV1();
                V1Connection.CreateRequest(V1Instance, Messages);
            }

            Logs.LogEvent("Operation - V1 Request Asset Creator Check Complete.");

        }

    }
}
