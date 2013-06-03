using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace PMMP
{
    public class BarChartUtilities
    {
        public static void LoadChartData(ChartPart chartPart, System.Data.DataTable dataTable)
        {
            Repository.Utility.WriteLog("LoadChartData started", System.Diagnostics.EventLogEntryType.Information);
            Chart chart = chartPart.ChartSpace.Elements<Chart>().First();
            BarChart bc = chart.Descendants<BarChart>().FirstOrDefault();

            if (bc != null)
            {
                BarChartSeries bcs1 = bc.Elements<BarChartSeries>().FirstOrDefault();
                BarChartSeries bcs2 = bc.Elements<BarChartSeries>().ElementAt(1);
                if (bcs1 != null && bcs2 != null)
                {
                    var categories = bcs1.Descendants<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData>().First();

                    StringReference csr = categories.Descendants<StringReference>().First();
                    csr.Formula.Text = String.Format("Sheet1!$A$2:$A${0}", dataTable.Rows.Count + 1);

                    StringCache sc = categories.Descendants<StringCache>().First();

                    CreateStringPoints(sc, dataTable.Rows.Count - 1);

                    //Series 1
                    var values1 = bcs1.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();

                    NumberReference vnr1 = values1.Descendants<NumberReference>().First();
                    vnr1.Formula.Text = String.Format("Sheet1!$B$2:$B${0}", dataTable.Rows.Count + 1);

                    NumberingCache nc1 = values1.Descendants<NumberingCache>().First();

                    CreateNumericPoints(nc1, dataTable.Rows.Count - 1);

                    //Series 2
                    var values2 = bcs2.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();

                    NumberReference vnr2 = values2.Descendants<NumberReference>().First();
                    vnr2.Formula.Text = String.Format("Sheet1!$C$2:$C${0}", dataTable.Rows.Count + 1);

                    NumberingCache nc2 = values2.Descendants<NumberingCache>().First();

                    CreateNumericPoints(nc2, dataTable.Rows.Count - 1);

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        NumericValue sv = sc.Elements<StringPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                        sv.Text = dataTable.Rows[i][0].ToString();

                        NumericValue nv1 = nc1.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                        nv1.Text = ((DateTime)dataTable.Rows[i][1]).ToOADate().ToString();

                        NumericValue nv2 = nc2.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                        nv2.Text = "10";
                    }
                }
            }
            Repository.Utility.WriteLog("LoadChartData completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        private static void CreateNumericPoints(NumberingCache nc, int count)
        {
            Repository.Utility.WriteLog("CreateNumericPoints started", System.Diagnostics.EventLogEntryType.Information);
            var np1 = nc.Elements<NumericPoint>().ElementAt(0);

            for (int i = 0; i < count; i++)
            {
                var npref = nc.Elements<NumericPoint>().ElementAt(i);

                var np = (NumericPoint)np1.Clone();
                np.Index = (UInt32)i + 1;

                nc.InsertAfter(np, npref);
            }
            Repository.Utility.WriteLog("CreateNumericPoints completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        private static void CreateStringPoints(StringCache sc, int count)
        {
            Repository.Utility.WriteLog("CreateStringPoints started", System.Diagnostics.EventLogEntryType.Information);
            var sp1 = sc.Elements<StringPoint>().ElementAt(0);

            for (int i = 0; i < count; i++)
            {
                var spref = sc.Elements<StringPoint>().ElementAt(i);

                var sp = (StringPoint)sp1.Clone();
                sp.Index = (UInt32)i + 1;

                sc.InsertAfter(sp, spref);
            }
            Repository.Utility.WriteLog("CreateStringPoints completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }
    }
}
