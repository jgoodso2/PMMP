using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Repository
{
    public class LookupDTO
    {
        public Guid ID {get;set;}
        public Guid ParentID { get; set; }
        public string Text { get; set; }
        public int SortIndex { get; set; }
        public int RowLevel { get; set; }
        public string DotNotation { get; set; }
        public string COID { get; set; }
        public DateTime ProcessingDate { get; set; }
        public DateTime LastLoad { get; set; }
        public LookupDTO ParentNode {get;set;}
    }
}
