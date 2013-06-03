using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Repository;

namespace PMMP
{
    public class PMPDocument :IPMPDocument
    {
        public Stream CreateDocument(string template,byte[] fileName,string projectUID)
        {
            Utility.WriteLog("Create Document Started", System.Diagnostics.EventLogEntryType.Information);
            Stream stream  = PresentationDocumentFactory.CreateDocument(template, fileName, projectUID);
            Utility.WriteLog("Create Document ompleted Successfully", System.Diagnostics.EventLogEntryType.Information);
            return stream;
        }
    }
}
