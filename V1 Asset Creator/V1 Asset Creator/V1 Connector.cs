using OpenPop.Mime;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using VersionOne.SDK;
using VersionOne.SDK.APIClient;
using VersionOne_Asset_Creator;

namespace V1_Asset_Creator
{
    public class V1_Connector
    {
        public V1_Connector()
        {

        }
        /*
        public bool ConnectToV1()
        {
            bool Connected = false;

            CreateDefect();

            V1Connector connector = V1Connector
                .WithInstanceUrl("https://www8.v1host.com/Terrys_Instance052215")
                .WithUserAgentHeader("V1-Asset-Creator", "1.0")
                .WithAccessToken("1.eqpQfm0root/o7s+lOXLWYK3wCQ=")
                .Build();

            IServices services = new Services(connector);


            Oid projectId = services.GetOid("Scope:1111");
            IAssetType defectType = services.Meta.GetAssetType("Defect");
            Asset NewDefect = services.New(defectType, projectId);
            IAttributeDefinition nameAttribute = defectType.GetAttributeDefinition("Name");
            NewDefect.SetAttributeValue(nameAttribute, "My New Defect");
            services.Save(NewDefect);

            
            
            return (Connected);

        }*/

        public V1Connector ConnectToV1()
        {
            InstanceSettings InstanceInfo = new InstanceSettings();

            InstanceInfo = InstanceInfo.GetInstanceSettings();

            V1Connector connector = null;

            if(String.IsNullOrEmpty(InstanceInfo.ProxyURL)==true)
            {
                connector = V1Connector
                    .WithInstanceUrl(InstanceInfo.InstanceURL)
                    .WithUserAgentHeader("V1-Asset-Creator", "1.0")
                    .WithAccessToken(InstanceInfo.AccessToken)
                    .Build();
            }
            else
            {
                Uri proxyURL = new Uri(InstanceInfo.ProxyURL); 
                ProxyProvider proxySettings = new ProxyProvider(proxyURL,
                                                                InstanceInfo.ProxyUserName,
                                                                InstanceInfo.ProxyPassword);
                
                connector = V1Connector
                    .WithInstanceUrl(InstanceInfo.InstanceURL)
                    .WithUserAgentHeader("V1-Asset-Creator", "1.0")
                    .WithAccessToken(InstanceInfo.AccessToken)
                    .WithProxy(proxySettings)
                    .Build();

            }
            /*
            V1Connector connector = V1Connector
                .WithInstanceUrl("https://www8.v1host.com/Terrys_Instance052215")
                .WithUserAgentHeader("V1-Asset-Creator", "1.0")
                .WithAccessToken("1.eqpQfm0root/o7s+lOXLWYK3wCQ=")
                .Build();
            */
            return (connector);

        }

        public void CreateDefect(V1Connector CurrentV1Connection, List<Message> CurrentEmails)
        {
            V1XMLHelper V1XML = new V1XMLHelper();
            V1Logging Logs = new V1Logging();
            string ProjectScope = "Something for Testing";
            string EmailBody;
            MessagePart MessageParts;
            SMTPResponse V1Response = new SMTPResponse();
            MailChecker MailCheck = new MailChecker();

            IServices services = new Services(CurrentV1Connection);


            try
            {
                foreach (Message MailItem in CurrentEmails)
                {
                    for (int ToCtr = 0; ToCtr < MailItem.Headers.To.Count; ToCtr++)
                    {
                        ProjectScope = V1XML.GetProjectScope(MailItem.Headers.To[ToCtr].Address);
                        if (ProjectScope != null)
                        {
                            Oid projectId = services.GetOid(ProjectScope);
                            IAssetType defectType = services.Meta.GetAssetType("Defect");
                            Asset NewDefect = services.New(defectType, projectId);
                            IAttributeDefinition nameAttribute = defectType.GetAttributeDefinition("Name");
                            NewDefect.SetAttributeValue(nameAttribute, MailItem.Headers.Subject.ToString());

                            MessageParts = MailItem.FindFirstHtmlVersion();
                            if (MessageParts == null)
                                MessageParts = MailItem.FindFirstPlainTextVersion();
                            EmailBody = MessageParts.GetBodyAsText();

                            Logs.LogEvent("Operation - Creating Defect for " + MailItem.Headers.To[ToCtr].Address);

                            IAttributeDefinition descriptionAttribute = defectType.GetAttributeDefinition("Description");
                            NewDefect.SetAttributeValue(descriptionAttribute, EmailBody);
                            IAttributeDefinition FoundByAttribute = defectType.GetAttributeDefinition("FoundBy");
                            NewDefect.SetAttributeValue(FoundByAttribute, MailItem.Headers.From.ToString());
                            services.Save(NewDefect);

                            IAttributeDefinition DefectIDAttribute = defectType.GetAttributeDefinition("Number");
                            Query IDQuery = new Query(NewDefect.Oid);
                            IDQuery.Selection.Add(DefectIDAttribute);
                            QueryResult ResultID = services.Retrieve(IDQuery);
                            Asset defect = ResultID.Assets[0];

                            //NewDefect.GetAttribute(DefectIDAttribute).Value
                            Logs.LogEvent("Operation - Sending Response to Defect Sender.");
                            //Commented out the Response back to the Sender per John Waedekin 
                            //V1Response.SendResponse(MailItem, defect.GetAttribute(DefectIDAttribute).Value + " " + NewDefect.GetAttribute(nameAttribute).Value, ProjectScope);
                            MailCheck.DeleteMessagesOnServer(MailItem, ProjectScope);
                        }

                    }

                }
            }
            catch (Exception ex)
            {
              Logs.LogEvent("ERROR - Creating Defect - " + ex.InnerException.Message);
            }


        }

        public bool V1TestConnection(V1Connector CurrentV1Connection)
        {
            V1Logging Logs = new V1Logging();
            bool TestResult = false;
            IServices services = new Services(CurrentV1Connection);

            try
            {
                Logs.LogEvent("Operation - Attempting Test Connection to VersionOne");
                
                Logs.LogEvent("Test - Testing URI - Pulling Meta");
                IAssetType defectType = services.Meta.GetAssetType("Scope");
                Logs.LogEvent("Test - Testing URI - SUCCESS");


                Logs.LogEvent("Test - Testing Token - Pulling Data");
                Oid projectId = services.GetOid("Scope:0");
                IAssetType ScopeType = services.Meta.GetAssetType("Scope");
                Asset NewScope = services.New(defectType, projectId);

                Logs.LogEvent("Test - Testing Token - SUCCESS");

                TestResult = true;

                Logs.LogEvent("Operation - VersionOne Connection Test Complete - " + TestResult.ToString());
            }
            catch(Exception ex)
            {
                Logs.LogEvent("ERROR - Test Failure - " + ex.ToString());
                Logs.LogEvent("Operation - VersionOne Connection Test Complete - " + TestResult.ToString());
            }

            return (TestResult);
        }


    }



}
