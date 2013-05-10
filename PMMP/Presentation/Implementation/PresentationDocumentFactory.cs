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
        public static Stream CreateDocument(string template, byte[] fileName)
        {
            switch (template)
            {
                case "Presentation":
                    PresentationDirector director = new PresentationDirector();
                    return director.Construct(new PresentationBuilder(),fileName);
            }
            return null;
        }
    }
}