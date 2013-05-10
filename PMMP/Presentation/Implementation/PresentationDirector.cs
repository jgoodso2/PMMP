using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace PMMP
{
    class PresentationDirector : IDirector
    {
        public Stream Construct(IBuilder builder, byte[] fileName)
        {
            builder.BuildDataFromDataSource();
           return builder.CreateDocument(fileName);
        }
    }
}
