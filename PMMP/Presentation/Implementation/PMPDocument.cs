using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace PMMP
{
    public class PMPDocument :IPMPDocument
    {
        public Stream CreateDocument(string template,byte[] fileName,string projectUID)
        {
            return PresentationDocumentFactory.CreateDocument(template, fileName, projectUID);
        }
    }

   
}
