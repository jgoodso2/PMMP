using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PMMP
{
    public class TaskItem
    {
        public int ID { get; set; }
        public string UniqueID { get; set; }
        public string DrivingPath { get; set; }
        public string Task { get; set; }
        public string Duration { get; set; }
        public string Predecessor { get; set; }
        public DateTime? Start { get; set; }
        public DateTime? Finish { get; set; }
        public DateTime? Deadline { get; set; }
        public DateTime? ModifiedOn { get; set; }
        public string[] ShowOn { get; set; }
        public int WorkCompletePercentage { get; set; }
    }
}
