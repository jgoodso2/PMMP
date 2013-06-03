using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Security.Principal;

namespace Repository
{
    public static class Utility
    {

        public static void WriteLog(string message, EventLogEntryType type)
        {
            //WindowsIdentity winId = (WindowsIdentity)System.Security.Principal.WindowsIdentity.GetCurrent();
            //WindowsImpersonationContext ctx = null;
            try
            {
                // Start impersonating
                //ctx = winId.Impersonate();
                // Now impersonating
                // Access resources using the identity of the authenticated user
                System.Diagnostics.EventLog appLog =
 new System.Diagnostics.EventLog();
                appLog.Source = "PMM Presentation";
                appLog.WriteEntry(message, type);
            }

            // Prevent exceptions from propagating
            catch
            {
            }
            //finally
            //{
            //    // Revert impersonation
            //    if (ctx != null)
            //        ctx.Undo();
            //}

        }

    }
}