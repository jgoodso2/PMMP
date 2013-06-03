using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Repository
{
    /// <summary>
    /// 
    /// </summary>
    class CustomFieldDTO
    {
        public string Text { get; set; }
        public string Description { get; set; }

        internal void AppendText(string text)
        {
            Text += "," + text;
            Text = string.Join(",", Text.Split(",".ToCharArray(),StringSplitOptions.RemoveEmptyEntries));
        }

        internal void AppendDescription(string description)
        {
            Description += "," + description;
            Description = string.Join(",", Description.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries));
        }
    }
}
