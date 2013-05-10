using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace Repository
{
    public static class Utility
    {
        public static void WriteLog(string message,EventLogEntryType type)
        {
            System.Diagnostics.EventLog appLog =
    new System.Diagnostics.EventLog();
            appLog.Source = "Look Up Tree Update from SSIS";
            appLog.WriteEntry(message,type);
        }
    }
}
