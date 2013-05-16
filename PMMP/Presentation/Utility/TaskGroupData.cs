using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PMMP
{
    public class TaskGroupData
    {
        public IList<TaskItemGroup> TaskItemGroups { get; set; }
        public Dictionary<string, IList<TaskItem>> ChartsData { get; set; }
        public IList<TaskItemGroup> LateTaskGroups { get; set; }
    }
}
