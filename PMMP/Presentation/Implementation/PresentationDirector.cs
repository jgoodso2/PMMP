using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace PMMP
{
    class PresentationDirector : IDirector
    {
        public Stream Construct(IBuilder builder, byte[] fileName,string projectGuid)
        {
            Repository.Utility.WriteLog("Construct started", System.Diagnostics.EventLogEntryType.Information);
            //builder.BuildDataFromDataSource(projectGuid);
           Stream oStream = builder.CreateDocument(fileName,projectGuid);
           Repository.Utility.WriteLog("Construct completed successfully", System.Diagnostics.EventLogEntryType.Information);
           return oStream;
        }
    }
}
