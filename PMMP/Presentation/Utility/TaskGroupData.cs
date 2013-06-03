using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Repository;

namespace PMMP
{
    /// <summary>
    /// 
    /// </summary>
    public class TaskGroupData
    {
        public IList<TaskItemGroup> TaskItemGroups { get; set; }
        public Dictionary<string, IList<TaskItem>> ChartsData { get; set; }
        public IList<TaskItemGroup> LateTaskGroups { get; set; }
        public FiscalMonth FiscalPeriod { get; set; }
    }
}
