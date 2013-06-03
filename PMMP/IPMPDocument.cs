using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace PMMP
{
    /// <summary>
    /// 
    /// </summary>
    interface IPMPDocument
    {
        Stream CreateDocument(string template, byte[] fileName,string projectUID);
    }
}
