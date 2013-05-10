using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Repository
{
    public class PSIDataSetFactory
    {
        public static IPSIDataSet GetPSISDataSet(string type)
        {
            switch (type)
            {
                case "Lookup": return new LookupPSIDataSet();
            }
            return null;
        }
    }
}
