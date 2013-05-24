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
                tr.Append(CreateTextCell(item.UniqueID.ToString()));
                tr.Append(CreateTextCell(item.Task));
                if (Convert.ToInt32(item.Duration) != 0)
                {
                    tr.Append(CreateTextCell((Convert.ToInt32(item.Duration) / 4800).ToString()));
                }
                else
                {
                    tr.Append(CreateTextCell(item.Duration.ToString()));
                }
                tr.Append(CreateTextCell(item.Predecessor));
                tr.Append(CreateTextCell(item.Start.HasValue ? item.Start.Value.ToShortDateString() : String.Empty));
                tr.Append(CreateTextCell(item.Finish.HasValue ? item.Start.Value.ToShortDateString() : String.Empty));
                //tr.Append(CreateTextCell(item.ModifiedOn.HasValue ? item.ModifiedOn.Value.ToShortDateString() : String.Empty));
                table.Append(tr);
            }
        }

        static TableCell CreateTextCell(string text, params System.Drawing.Color[] color)
        {
            TableCellProperties tableCellProperty = new TableCellProperties();
            if (color.Length > 0)
            {

                SolidFill solidFill1 = new SolidFill();
                RgbColorModelHex rgbColorModelHex1 = new RgbColorModelHex() { Val = color[0].ToHexString() };
                solidFill1.Append(rgbColorModelHex1);
                tableCellProperty.Append(solidFill1);
            }
            TableCell tc = new TableCell(
            new TextBody(
            new BodyProperties(),
            new Paragraph(
            new Run(
            new RunProperties() { FontSize = 1200 },
            new Text(text)))),
            tableCellProperty);

            return tc;

        }

        internal static void PopulateLateTasksTable(Table table, IList<TaskItem> iList, Repository.FiscalMonth fiscalMonth)
        {

            foreach (TaskItem item in iList)
            {
                //shp.Table.Cell(2, 2).Shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                TableRow tr = new TableRow();
                tr.Height = 304800L;
                tr.Append(CreateTextCell(item.ID.ToString()));
                //tr.Append(CreateTextCell(item.UniqueID.ToString()));
                tr.Append(CreateTextCell(item.CA));
                tr.Append(CreateTextCell(item.Task));
                if (Convert.ToInt32(item.TotalSlack) != 0)
                {
                    tr.Append(CreateTextCell((Convert.ToInt32(item.TotalSlack) / 4800).ToString()));
                }
                else
                {
                    tr.Append(CreateTextCell(item.TotalSlack.ToString()));
                }
                tr.Append(CreateTextCell(item.Start.HasValue ? item.Start.Value.ToVeryShortDateString() : String.Empty));
                tr.Append(CreateTextCell(item.Finish.HasValue ? item.Finish.Value.ToVeryShortDateString() : String.Empty));
                TableCell baseLineStart;
                TableCell baseLineFinish;
                if (!item.BaseLineStart.HasValue)
                {
                    baseLineStart = CreateTextCell(String.Empty);
                }
                else
                {
                    if (item.Start.HasValue && item.Start.Value <= fiscalMonth.From && item.Start.Value >= fiscalMonth.To && item.Start.Value.Month > fiscalMonth.To.Month)
                    {
                        baseLineStart = CreateTextCell(item.BaseLineStart.Value.ToVeryShortDateString(), System.Drawing.Color.Red);
                    }

                    else if (item.Start.HasValue && item.Start > item.BaseLineStart)
                    {
                        baseLineStart = CreateTextCell(item.BaseLineStart.Value.ToVeryShortDateString(), System.Drawing.Color.Yellow);
                    }

                    else if (item.Start.HasValue && item.Start <= item.BaseLineStart)
                    {
                        baseLineStart = CreateTextCell(item.BaseLineStart.Value.ToVeryShortDateString(), System.Drawing.Color.Green);
                    }
                    else if (item.Start.HasValue && item.Start.Value.Year == DateTime.Now.Year && item.Start.Value.Month < DateTime.Now.Month
                        && item.BaseLineStart.Value.Month == DateTime.Now.Month)
                    {
                        baseLineStart = CreateTextCell(item.BaseLineStart.Value.ToVeryShortDateString(), System.Drawing.Color.Blue);
                    }
                    else
                    {
                        baseLineStart = CreateTextCell(item.BaseLineStart.Value.ToVeryShortDateString());
                    }
                }

                if (!item.BaseLineFinish.HasValue)
                {
                    baseLineFinish = CreateTextCell(String.Empty);
                }
                else
                {
                    if (item.Finish.HasValue && item.Finish.Value <= fiscalMonth.From && item.Finish.Value >= fiscalMonth.To && item.Finish.Value.Month > fiscalMonth.To.Month)
                    {
                        baseLineFinish = CreateTextCell(item.BaseLineFinish.Value.ToVeryShortDateString(), System.Drawing.Color.Red);
                    }

                    else if (item.Finish.HasValue && item.Finish > item.BaseLineFinish)
                    {
                        baseLineFinish = CreateTextCell(item.BaseLineFinish.Value.ToVeryShortDateString(), System.Drawing.Color.Yellow);
                    }

                    else if (item.Finish.HasValue && item.Finish <= item.BaseLineFinish)
                    {
                        baseLineFinish = CreateTextCell(item.BaseLineFinish.Value.ToVeryShortDateString(), System.Drawing.Color.Green);
                    }
                    else if (item.Finish.HasValue && item.Finish.Value.Year == DateTime.Now.Year && item.Finish.Value.Month < DateTime.Now.Month
                        && item.BaseLineFinish.Value.Month == DateTime.Now.Month)
                    {
                        baseLineFinish = CreateTextCell(item.BaseLineFinish.Value.ToVeryShortDateString(), System.Drawing.Color.Blue);
                    }
                    else
                    {
                        baseLineFinish = CreateTextCell(item.BaseLineFinish.Value.ToVeryShortDateString());
                    }
                }

                tr.Append(baseLineStart);
                tr.Append(baseLineFinish);
                tr.Append(Convert.ToInt32(item.Hours) != 0 ? CreateTextCell((Convert.ToInt32(item.Hours) / 60000).ToString()) : CreateTextCell(item.Hours));
                tr.Append(CreateTextCell(item.PMT));
                tr.Append(CreateTextCell(item.ReasonRecovery));
                tr.Append(CreateTextCell(""));
                if (Convert.ToInt32(item.Duration) != 0)
                {
                    tr.Append(CreateTextCell((Convert.ToInt32(item.Duration) / 4800).ToString()));
                }
                else
                {
                    tr.Append(CreateTextCell(item.Duration));
                }
                tr.Append(CreateTextCell(item.EstStart.HasValue ? item.EstStart.Value.ToVeryShortDateString() : String.Empty));
                tr.Append(CreateTextCell(item.EstFinish.HasValue ? item.EstFinish.Value.ToVeryShortDateString() : String.Empty));
                table.Append(tr);
            }
        }
    }
}
