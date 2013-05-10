using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace PMMP
{
    interface IBuilder
    {
        object BuildDataFromDataSource();
        MemoryStream CreateDocument(byte[] template);
    }
}
