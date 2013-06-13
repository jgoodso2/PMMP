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
            Repository.Utility.WriteLog("PopulateTable started ", System.Diagnostics.EventLogEntryType.Information);
            RunProperties contentCellProperties = table.ChildElements[3].ChildElements.ToList()[0].Descendants<RunProperties>().ToList()[0];
            table.ChildElements[3].Remove();
            
            foreach (TaskItem item in items)
            {
                TableRow tr = new TableRow();
                tr.Height = 304800L;
                tr.Append(CreateTextCell(item.ID.ToString(), contentCellProperties));
                tr.Append(CreateTextCell(item.UniqueID.ToString(), contentCellProperties));
                tr.Append(CreateTextCell(item.Task, contentCellProperties));
                if (Convert.ToInt32(item.Duration) != 0)
                {
                    tr.Append(CreateTextCell((Convert.ToInt32(item.Duration) / 4800).ToString(), contentCellProperties));
                }
                else
                {
                    tr.Append(CreateTextCell(item.Duration.ToString(), contentCellProperties));
                }
                tr.Append(CreateTextCell(item.Predecessor, contentCellProperties));
                tr.Append(CreateTextCell(item.Start.HasValue ? item.Start.Value.ToShortDateString() : String.Empty, contentCellProperties));
                tr.Append(CreateTextCell(item.Finish.HasValue ? item.Finish.Value.ToShortDateString() : String.Empty, contentCellProperties));
                //tr.Append(CreateTextCell(item.ModifiedOn.HasValue ? item.ModifiedOn.Value.ToShortDateString() : String.Empty));
                table.Append(tr);
            }
            Repository.Utility.WriteLog("PopulateTable completed successfully ", System.Diagnostics.EventLogEntryType.Information);
        }

        static TableCell CreateTextCell(string text,RunProperties runProperties, params System.Drawing.Color[] color)
        {
            try
            {
                Repository.Utility.WriteLog("CreateTextCell started ", System.Diagnostics.EventLogEntryType.Information);
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
                runProperties.Clone() as RunProperties,
                new Text(text)))),
                tableCellProperty);
                Repository.Utility.WriteLog("CreateTextCell completed successfully ", System.Diagnostics.EventLogEntryType.Information);
                return tc;
            }
            catch
            {
                return new TableCell();
            }

        }

        internal static void PopulateLateOrUpComingTasksTable(Table table, IList<TaskItem> iList, Repository.FiscalUnit fiscalMonth)
        {
            Repository.Utility.WriteLog("PopulateLateTasksTable started ", System.Diagnostics.EventLogEntryType.Information);
            RunProperties contentCellProperties = table.ChildElements[3].ChildElements.ToList()[0].Descendants<RunProperties>().ToList()[0];
            table.ChildElements[3].Remove();
            foreach (TaskItem item in iList)
            {
                //shp.Table.Cell(2, 2).Shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                TableRow tr = new TableRow();
                tr.Height = 304800L;
                tr.Append(CreateTextCell(item.UniqueID.ToString(), contentCellProperties));
                //tr.Append(CreateTextCell(item.UniqueID.ToString()));
                tr.Append(CreateTextCell(item.CA, contentCellProperties));
                tr.Append(CreateTextCell(item.Task, contentCellProperties));
                if (Convert.ToInt32(item.TotalSlack) != 0)
                {
                    tr.Append(CreateTextCell((Convert.ToInt32(item.TotalSlack) / 4800).ToString(), contentCellProperties));
                }
                else
                {
                    tr.Append(CreateTextCell(item.TotalSlack.ToString(), contentCellProperties));
                }
                tr.Append(CreateTextCell(item.Start.HasValue ? item.Start.Value.ToVeryShortDateString() : String.Empty, contentCellProperties));
                tr.Append(CreateTextCell(item.Finish.HasValue ? item.Finish.Value.ToVeryShortDateString() : String.Empty, contentCellProperties));
                TableCell baseLineStart;
                TableCell baseLineFinish;
                if (!item.BaseLineStart.HasValue)
                {
                    baseLineStart = CreateTextCell(String.Empty, contentCellProperties);
                }
                else
                {
                    if (item.Start.HasValue && item.Start.Value <= fiscalMonth.From && item.Start.Value >= fiscalMonth.To && item.Start.Value.Month > fiscalMonth.To.Month)
                    {
                        baseLineStart = CreateTextCell(item.BaseLineStart.Value.ToVeryShortDateString(), contentCellProperties,System.Drawing.Color.Red);
                    }

                    else if (item.Start.HasValue && item.Start > item.BaseLineStart)
                    {
                        baseLineStart = CreateTextCell(item.BaseLineStart.Value.ToVeryShortDateString(), contentCellProperties, System.Drawing.Color.Yellow);
                    }

                    else if (item.Start.HasValue && item.Start <= item.BaseLineStart)
                    {
                        baseLineStart = CreateTextCell(item.BaseLineStart.Value.ToVeryShortDateString(), contentCellProperties, System.Drawing.Color.Green);
                    }
                    else if (item.Start.HasValue && item.Start.Value.Year == DateTime.Now.Year && item.Start.Value.Month < DateTime.Now.Month
                        && item.BaseLineStart.Value.Month == DateTime.Now.Month)
                    {
                        baseLineStart = CreateTextCell(item.BaseLineStart.Value.ToVeryShortDateString(), contentCellProperties, System.Drawing.Color.Blue);
                    }
                    else
                    {
                        baseLineStart = CreateTextCell(item.BaseLineStart.Value.ToVeryShortDateString(), contentCellProperties);
                    }
                }

                if (!item.BaseLineFinish.HasValue)
                {
                    baseLineFinish = CreateTextCell(String.Empty, contentCellProperties);
                }
                else
                {
                    if (item.Finish.HasValue && item.Finish.Value <= fiscalMonth.From && item.Finish.Value >= fiscalMonth.To && item.Finish.Value.Month > fiscalMonth.To.Month)
                    {
                        baseLineFinish = CreateTextCell(item.BaseLineFinish.Value.ToVeryShortDateString(), contentCellProperties, System.Drawing.Color.Red);
                    }

                    else if (item.Finish.HasValue && item.Finish > item.BaseLineFinish)
                    {
                        baseLineFinish = CreateTextCell(item.BaseLineFinish.Value.ToVeryShortDateString(), contentCellProperties, System.Drawing.Color.Yellow);
                    }

                    else if (item.Finish.HasValue && item.Finish <= item.BaseLineFinish)
                    {
                        baseLineFinish = CreateTextCell(item.BaseLineFinish.Value.ToVeryShortDateString(), contentCellProperties, System.Drawing.Color.Green);
                    }
                    else if (item.Finish.HasValue && item.Finish.Value.Year == DateTime.Now.Year && item.Finish.Value.Month < DateTime.Now.Month
                        && item.BaseLineFinish.Value.Month == DateTime.Now.Month)
                    {
                        baseLineFinish = CreateTextCell(item.BaseLineFinish.Value.ToVeryShortDateString(), contentCellProperties, System.Drawing.Color.Blue);
                    }
                    else
                    {
                        baseLineFinish = CreateTextCell(item.BaseLineFinish.Value.ToVeryShortDateString(), contentCellProperties);
                    }
                }

                tr.Append(baseLineStart);
                tr.Append(baseLineFinish);
                tr.Append(Convert.ToDouble(item.Hours) != 0 ? CreateTextCell(((int)(Convert.ToDouble(item.Hours) / 60000)).ToString(), contentCellProperties) : CreateTextCell(item.Hours, contentCellProperties));
                tr.Append(CreateTextCell(item.PMT, contentCellProperties));
                tr.Append(CreateTextCell(item.ReasonRecovery, contentCellProperties));
                tr.Append(CreateTextCell("", contentCellProperties));
                if (Convert.ToInt32(item.Duration) != 0)
                {
                    tr.Append(CreateTextCell((Convert.ToInt32(item.Duration) / 4800).ToString(), contentCellProperties));
                }
                else
                {
                    tr.Append(CreateTextCell(item.Duration, contentCellProperties));
                }
                tr.Append(CreateTextCell(item.EstStart.HasValue ? item.EstStart.Value.ToVeryShortDateString() : String.Empty, contentCellProperties));
                tr.Append(CreateTextCell(item.EstFinish.HasValue ? item.EstFinish.Value.ToVeryShortDateString() : String.Empty, contentCellProperties));
                table.Append(tr);
            }
            Repository.Utility.WriteLog("PopulateLateTasksTable completed successfully ", System.Diagnostics.EventLogEntryType.Information);
        }
    }
}
