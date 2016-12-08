using System;
using OpenPop;
using OpenPop.Pop3;
using System.Collections.Generic;
using OpenPop.Mime;
using System.Xml;
using System.Diagnostics;
using VersionOne_Asset_Creator;
using V1_Asset_Creator;

public class MailChecker
{

    public MailChecker()
    {

    }

    public List<Message> FetchAllMessages(string hostname, int port, bool useSsl, string username, string password)
    {
        // The client disconnects from the server when being disposed
        using (Pop3Client client = new Pop3Client())
        {

            // Connect to the server
            client.Connect(hostname, port, useSsl);

            // Authenticate ourselves towards the server
            client.Authenticate(username, password);

            // Get the number of messages in the inbox
            int messageCount = client.GetMessageCount();

            // We want to download all messages
            List<Message> allMessages = new List<Message>(messageCount);

            // Messages are numbered in the interval: [1, messageCount]
            // Ergo: message numbers are 1-based.
            // Most servers give the latest message the highest number
            for (int i = messageCount; i > 0; i--)
            {
                allMessages.Add(client.GetMessage(i));
            }

            // Now return the fetched messages
            return allMessages;
        }
    }

    public List<Message> FetchAllMessagesWithMultipleAccounts()
    {
        V1Logging Logs = new V1Logging();
        // The client disconnects from the server when being disposed
        using (Pop3Client client = new Pop3Client())
        {

            // Get the number of messages in the inbox
            int messageCount;
            // We want to download all messages
            List<Message> allMessages = new List<Message>();

            XmlDocument EmailConfig = new XmlDocument();
            EmailConfig.PreserveWhitespace = true;
            EmailConfig.Load(AppDomain.CurrentDomain.BaseDirectory + "V1EmailConfig.xml");

            XmlElement EmailRoot = EmailConfig.DocumentElement;
            XmlNodeList EmailNodes = EmailRoot.SelectNodes("/EmailInfo/EmailAccount");



            try
            {
                foreach (XmlNode EmailNode in EmailNodes)
                {
                    Logs.LogEvent("Operation - Checking Email Server " + EmailNode["EmailAddress"].InnerText);
                    // Connect to the server
                    client.Connect(EmailNode["EmailPOP3"].InnerText, Convert.ToInt32(EmailNode["POP3PortNumber"].InnerText), Convert.ToBoolean(EmailNode["EmailUseSSL"].InnerText));

                    // Authenticate ourselves towards the server
                    client.Authenticate(EmailNode["EmailAddress"].InnerText, EmailNode["EmailPassword"].InnerText);

                    if (client.Connected == true)
                    {
                        // Get the number of messages in the inbox
                        messageCount = client.GetMessageCount();
                        // We want to download all messages
                        //allMessages = new List<Message>(messageCount);

                        // Messages are numbered in the interval: [1, messageCount]
                        // Ergo: message numbers are 1-based.
                        // Most servers give the latest message the highest number
                        for (int i = messageCount; i > 0; i--)
                        {
                            allMessages.Add(client.GetMessage(i));
                        }

                        client.Disconnect();
                    }
                    else
                    {
                        Logs.LogEvent("ERROR - Fetching Messages - Can't connect to mail server");
                    }
                }
            }catch(Exception e)
            {
                Logs.LogEvent("ERROR - Fetching Messages - " + e);
            }
            



            
            // Now return the fetched messages
            Debug.WriteLine("Total Number of Messages Found: " + allMessages.Count.ToString());
            return allMessages;

        }
    }

    public void DeleteMessagesOnServer(Message DefectEmail, string ProjectScope)
    {
        EmailSettings EmailSetting = new EmailSettings();
        V1Logging Logs = new V1Logging();
        V1XMLHelper V1XMLHelp = new V1XMLHelper();
        string ToEmailAddress = V1XMLHelp.GetEmailAddressFromProjectScope(ProjectScope);


        try
        {
            EmailSetting = EmailSetting.GetEmailSettings(ToEmailAddress);

            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(EmailSetting.Pop3Server, EmailSetting.Pop3PortNumber, EmailSetting.UseSSLPop3);

                // Authenticate ourselves towards the server
                client.Authenticate(ToEmailAddress, EmailSetting.Password);
                if(client.Connected == true)
                {
                    // Get the number of messages on the POP3 server
                    int messageCount = client.GetMessageCount();

                    // Run trough each of these messages and download the headers
                    for (int messageItem = messageCount; messageItem > 0; messageItem--)
                    {
                        // If the Message ID of the current message is the same as the parameter given, delete that message
                        if (client.GetMessageHeaders(messageItem).MessageId == DefectEmail.Headers.MessageId)
                        {
                            // Delete
                            client.DeleteMessage(messageItem);
                        }
                    }
                    client.Disconnect();
                }else
                {
                    Logs.LogEvent("ERROR - Deleting Message - Can't connect to server to delete message.");
                }


            }
        }
        catch(Exception ex)
        {
            Logs.LogEvent("ERROR - Deleting Message - " + ex.ToString());
        }

    }
}
