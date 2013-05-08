using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Drawing;

namespace PMMP
{
    public class TableUtilities
    {
        public static void PopulateTable(Table table, IList<TaskItem> items)
        {
            foreach (TaskItem item in items)
            {
                TableRow tr = new TableRow();
                tr.Height = 304800L;
                tr.Append(CreateTextCell(item.ID.ToString()));
                //tr.Append(CreateTextCell(item.UniqueID.ToString()));
                tr.Append(CreateTextCell(item.Task));
                tr.Append(CreateTextCell(item.Duration));
                tr.Append(CreateTextCell(item.Predecessor));
                tr.Append(CreateTextCell(item.Start.HasValue ? item.Start.Value.ToShortDateString() : String.Empty));
                tr.Append(CreateTextCell(item.Finish.HasValue ? item.Start.Value.ToShortDateString() : String.Empty));
                tr.Append(CreateTextCell(item.ModifiedOn.HasValue ? item.ModifiedOn.Value.ToShortDateString() : String.Empty));
                table.Append(tr);
            }
        }

        static TableCell CreateTextCell(string text)
        {
            TableCell tc = new TableCell(
            new TextBody(
            new BodyProperties(),
            new Paragraph(
            new Run(
            new RunProperties() { FontSize = 1200 },
            new Text(text)))),
            new TableCellProperties());

            return tc;
        }
    }
}
