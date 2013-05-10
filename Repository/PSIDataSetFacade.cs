using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace Repository
{
    public class PSIDataSetFacade
    {
        public void Update(IPSIDataSet dataSet)
        {
            //Read lookup tables from the project server
            DataSet source = dataSet.GetDataSet();
            //Read DataSet from the database updated by SSIS Package
            DataSet destination = dataSet.GetChanges();
            // Get the delta of the changes merged into the dataset returned
            DataSet delta = dataSet.GetDelta(source, destination);
            // Update project server with the changes
            dataSet.Update(delta);
        }
    }
}
