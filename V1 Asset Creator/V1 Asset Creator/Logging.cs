using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace V1_Asset_Creator
{
    class V1Logging
    {
        private static EventLog eventLog1 = new EventLog();

        public V1Logging ()
        {

        }

        public void LogEvent(string LogMessage)
        {
            // create an event source, specifying the name of a log that 
            // does not currently exist to create a new, custom log 
            if (!System.Diagnostics.EventLog.SourceExists("V1 Asset Creator"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "V1 Asset Creator", "V1 Asset Creator Log");
            }
            // configure the event log instance to use this source name
            eventLog1.Source = "V1 Asset Creator";

            eventLog1.WriteEntry(LogMessage);
        }
    }
}
