using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace PMMP
{
    public class PresentationDocumentFactory
    {
        public static Stream CreateDocument(string template, byte[] fileName,string projectGuid)
        {
            Repository.Utility.WriteLog("CreateDocument started", System.Diagnostics.EventLogEntryType.Information);
            switch (template)
            {
                case "Presentation":
                    PresentationDirector director = new PresentationDirector();
                    Repository.Utility.WriteLog("CreateDocument completed successfuly", System.Diagnostics.EventLogEntryType.Information);
                    return director.Construct(new PresentationBuilder(), fileName, projectGuid);
            }
            Repository.Utility.WriteLog("CreateDocument completed successfuly", System.Diagnostics.EventLogEntryType.Information);
            return null;
        }
    }
}