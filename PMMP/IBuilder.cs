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
    interface IBuilder
    {
        object BuildDataFromDataSource(string projectGuid);
        MemoryStream CreateDocument(byte[] template,string projectUID);
    }
}
