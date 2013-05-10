using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace Repository
{
    public interface IPSIDataSet
    {
        DataSet GetDataSet();
        DataSet GetChanges();
        DataSet GetDelta(DataSet source, DataSet changes);
        void Update(DataSet ds);
    }
}
