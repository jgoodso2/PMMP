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
            builder.BuildDataFromDataSource(projectGuid);
           return builder.CreateDocument(fileName,projectGuid);
        }
    }
}
