﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System.Data;

namespace PMMP
{
    public class BarChartUtilities
    {
        public static void LoadChartData(ChartPart chartPart, System.Data.DataTable dataTable)
        {
            Repository.Utility.WriteLog("LoadChartData started", System.Diagnostics.EventLogEntryType.Information);
            Chart chart = chartPart.ChartSpace.Elements<Chart>().First();
            BarChart bc1 = chart.Descendants<BarChart>().FirstOrDefault();
            BarChart bc2 = chart.Descendants<BarChart>().ElementAt(1);
            DateTime maxAxisvAlue=DateTime.MinValue;
            DateTime minAxisvAlue = DateTime.MaxValue;
            DateTime maxFinishValue = dataTable.AsEnumerable().Select(t=>t.Field<DateTime>("Finish")).Max();
            DateTime maxBFinishvalue = dataTable.AsEnumerable().Select(t => t.Field<DateTime>("BaseLineFinish")).Max();
            maxAxisvAlue = maxBFinishvalue > maxFinishValue ? maxBFinishvalue : maxFinishValue;

            DateTime minStartValue = dataTable.AsEnumerable().Select(t => t.Field<DateTime>("Start")).Min();
            DateTime minBStartvalue = dataTable.AsEnumerable().Select(t => t.Field<DateTime>("BaseLineStart")).Min();
            minAxisvAlue = minBStartvalue < minStartValue ? minBStartvalue : minStartValue;
            IEnumerable<ValueAxis> axes = chart.Descendants<ValueAxis>();
            foreach (ValueAxis axis in axes)
            {
                if (maxAxisvAlue != DateTime.MaxValue)
                {
                    axis.Scaling.MaxAxisValue.Val = maxAxisvAlue.ToOADate();
                }
                if (minAxisvAlue != DateTime.MinValue)
                {
                    axis.Scaling.MinAxisValue.Val = minAxisvAlue.ToOADate();
                }
                if (dataTable.Rows.Count > 0)
                {
                    axis.Elements<MajorUnit>().FirstOrDefault().Val = (axis.Scaling.MaxAxisValue.Val - axis.Scaling.MinAxisValue.Val) / dataTable.Rows.Count;
                }
            }
            

            
            if (bc1 != null && bc2 != null)
            {
                BarChartSeries bcs1 = bc1.Elements<BarChartSeries>().FirstOrDefault();
                BarChartSeries bcs2 = bc1.Elements<BarChartSeries>().ElementAt(1);
                BarChartSeries bcs3 = bc2.Elements<BarChartSeries>().FirstOrDefault();
                BarChartSeries bcs4 = bc2.Elements<BarChartSeries>().ElementAt(1);
                if (bcs1 != null && bcs2 != null)
                {
                    var categories = bcs1.Descendants<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData>().First();

                    StringReference csr = categories.Descendants<StringReference>().First();
                    csr.Formula.Text = String.Format("Sheet1!$A$2:$A${0}", dataTable.Rows.Count + 1);

                    StringCache sc = categories.Descendants<StringCache>().First();

                    CreateStringPoints(sc, dataTable.Rows.Count,false);

                    //Series 1
                    var values1 = bcs1.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();

                    NumberReference vnr1 = values1.Descendants<NumberReference>().First();
                    vnr1.Formula.Text = String.Format("Sheet1!$B$2:$B${0}", dataTable.Rows.Count + 1);

                    NumberingCache nc1 = values1.Descendants<NumberingCache>().First();

                    CreateNumericPoints(nc1, dataTable.Rows.Count,false);

                    //Series 2
                    var values2 = bcs2.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();

                    NumberReference vnr2 = values2.Descendants<NumberReference>().First();
                    vnr2.Formula.Text = String.Format("Sheet1!$C$2:$C${0}", dataTable.Rows.Count + 1);

                    NumberingCache nc2 = values2.Descendants<NumberingCache>().First();

                    CreateNumericPoints(nc2, dataTable.Rows.Count, false);

                    //Series 3
                    var values3 = bcs3.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();

                    NumberReference vnr3 = values3.Descendants<NumberReference>().First();
                    vnr3.Formula.Text = String.Format("Sheet1!$D$2:$D${0}", dataTable.Rows.Count + 1);

                    NumberingCache nc3 = values3.Descendants<NumberingCache>().First();

                    CreateNumericPoints(nc3, dataTable.Rows.Count,false);

                    //Series 4
                    var values4 = bcs4.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();

                    NumberReference vnr4 = values4.Descendants<NumberReference>().First();
                    vnr4.Formula.Text = String.Format("Sheet1!$E$2:$E${0}", dataTable.Rows.Count + 1);

                    NumberingCache nc4 = values4.Descendants<NumberingCache>().First();

                    CreateNumericPoints(nc4, dataTable.Rows.Count, false);

                    

                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        NumericValue sv = sc.Elements<StringPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                        sv.Text = dataTable.Rows[i]["Task"].ToString() + " | " + ((DateTime)dataTable.Rows[i]["Finish"]).ToString("MM/dd");

                        NumericValue nv1 = nc1.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                        NumericValue nv2 = nc2.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                        NumericValue nv3 = nc3.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                        NumericValue nv4 = nc4.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                        
                        nv1.Text = ((DateTime)dataTable.Rows[i]["Start"]).ToOADate().ToString();

                        nv2.Text = ((DateTime)dataTable.Rows[i]["Finish"] - (DateTime)dataTable.Rows[i]["Start"]).TotalDays > 4 ?
                            ((DateTime)dataTable.Rows[i]["Finish"] - (DateTime)dataTable.Rows[i]["Start"]).TotalDays.ToString() : "5";
                        
                        nv3.Text = ((DateTime)dataTable.Rows[i]["BaseLineStart"]).ToOADate().ToString();

                        nv4.Text = ((DateTime)dataTable.Rows[i]["BaseLineFinish"] - (DateTime)dataTable.Rows[i]["BaseLineStart"]).TotalDays > 4 ?
                            ((DateTime)dataTable.Rows[i]["BaseLineFinish"] - (DateTime)dataTable.Rows[i]["BaseLineStart"]).TotalDays.ToString() : "5";
                    }
                }
            }
            
            Repository.Utility.WriteLog("LoadChartData completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        private static void CreateNumericPoints(NumberingCache nc, int count,bool deleteClone)
        {
            Repository.Utility.WriteLog("CreateNumericPoints started", System.Diagnostics.EventLogEntryType.Information);
            var np1 = nc.Elements<NumericPoint>().ElementAt(0);

            
            for (int i = nc.Elements<NumericPoint>().Count() - 1; i > 0; i--)
            {
                nc.Elements<NumericPoint>().ElementAt(i).Remove();
            }
            for (int i = 0; i < count; i++)
            {
                var npref = nc.Elements<NumericPoint>().ElementAt(i);

                var np = (NumericPoint)np1.Clone();
                np.Index = (UInt32)i + 1;

                nc.InsertAfter(np, npref);
            }

            np1.Remove();
            Repository.Utility.WriteLog("CreateNumericPoints completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        private static void CreateStringPoints(StringCache sc, int count,bool deleteClone)
        {
            
            Repository.Utility.WriteLog("CreateStringPoints started", System.Diagnostics.EventLogEntryType.Information);
            var sp1 = sc.Elements<StringPoint>().ElementAt(0);

            for (int i = sc.Elements<StringPoint>().Count() - 1; i > 0; i--)
            {
                sc.Elements<StringPoint>().ElementAt(i).Remove();
            }

            for (int i = 0; i < count; i++)
            {
                var spref = sc.Elements<StringPoint>().ElementAt(i);

                var sp = (StringPoint)sp1.Clone();
                sp.Index = (UInt32)i + 1;

                sc.InsertAfter(sp, spref);
            }
            
            
                sp1.Remove();
            
            Repository.Utility.WriteLog("CreateStringPoints completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        internal static void LoadChartData(ChartPart chartPart, List<GraphDataGroup> list)
        {
            Repository.Utility.WriteLog("LoadChartData started", System.Diagnostics.EventLogEntryType.Information);
            Chart chart = chartPart.ChartSpace.Elements<Chart>().First();
            BarChart bc = chart.Descendants<BarChart>().FirstOrDefault();
            LineChart lc = chart.Descendants<LineChart>().FirstOrDefault();
            BarChartSeries bcs1 = null;
            BarChartSeries bcs2 = null;
            BarChartSeries bcs3 = null;
            BarChartSeries bcs4 = null;
            NumberingCache nc1 = null;
            NumberingCache nc2 = null;
            NumberingCache nc3 = null;
            NumberingCache nc4 = null;
            NumberingCache nc5 = null;
            NumberingCache nc6 = null;
            StringCache sc = null;
            if (bc != null)
            {
                 bcs1 = bc.Elements<BarChartSeries>().ElementAt(0);
                 bcs2 = bc.Elements<BarChartSeries>().ElementAt(1);
                 bcs3 = bc.Elements<BarChartSeries>().ElementAt(2);
                 bcs4 = bc.Elements<BarChartSeries>().ElementAt(3);
            }
                LineChartSeries lcs1 = lc.Elements<LineChartSeries>().ElementAt(0);
                LineChartSeries lcs2 = lc.Elements<LineChartSeries>().ElementAt(1);
                LineChartSeries lcs3=null;
                if(lc.Elements<LineChartSeries>().Count() > 2)
                lcs3 = lc.Elements<LineChartSeries>().ElementAt(2);
                LineChartSeries lcs4=null;
                if (lc.Elements<LineChartSeries>().Count() > 3)
                lcs4 = lc.Elements<LineChartSeries>().ElementAt(3);
                 int count = 0;
                    for (int j = 0; j < list.Count; j++)
                    {
                        try
                        {
                            GraphDataGroup graphGroup = list[j];
                            DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData categories;

                            if (graphGroup.Type == "BES" || graphGroup.Type == "BEFS" || graphGroup.Type == "BEFF" || graphGroup.Type == "BEF")
                            {
                                categories = lcs1.Descendants<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData>().ToList()[count];
                            }
                            else
                            {
                                categories = bcs1.Descendants<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData>().ToList()[count];
                            }
                            
                            

                            if (graphGroup.Type == "CS" || graphGroup.Type == "CF" ||  graphGroup.Type == "BES")
                            {

                                StringReference csr = categories.Descendants<StringReference>().First();
                                //csr.Formula.Text = String.Format("Sheet1!$A$2:$A${0}", list[j].Data.Count - 1);

                                sc = categories.Descendants<StringCache>().First();
                                CreateStringPoints(sc, list[j].Data.Count,true);
                            }

                            //Series 1
                            DocumentFormat.OpenXml.Drawing.Charts.Values values1;

                            if (graphGroup.Type == "BES" || graphGroup.Type == "BEFS" || graphGroup.Type == "BEFF" || graphGroup.Type == "BEF")
                            {
                                values1 = lcs1.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();
                            }
                            else
                            {
                                values1 = bcs1.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();
                            }

                           

                            if (graphGroup.Type == "CS" || graphGroup.Type == "CF" || graphGroup.Type == "BES" )
                            {
                                NumberReference vnr1 = values1.Descendants<NumberReference>().First();
                                //vnr1.Formula.Text = String.Format("Sheet1!$B$2:$B${0}", list[j].Data.Count - 1);
                                nc1 = values1.Descendants<NumberingCache>().First();
                                CreateNumericPoints(nc1, list[j].Data.Count,true);
                            }
                            //Series 2
                            DocumentFormat.OpenXml.Drawing.Charts.Values values2;
                            if (graphGroup.Type == "BES" || graphGroup.Type == "BEFS" || graphGroup.Type == "BEFF" || graphGroup.Type == "BEF")
                            {
                                values2 = lcs2.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();
                            }
                            else
                            {
                               values2 = bcs2.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();
                            }
                           
                            if (graphGroup.Type == "FCS" || graphGroup.Type == "FCF" || graphGroup.Type == "BEF")
                            {

                                nc2 = values2.Descendants<NumberingCache>().First();
                                NumberReference vnr2 = values2.Descendants<NumberReference>().First();
                                //vnr2.Formula.Text = String.Format("Sheet1!$C$2:$C${0}", list[j].Data.Count - 1 + 1);
                                CreateNumericPoints(nc2, list[j].Data.Count,true);
                            }

                            //Series 3
                            DocumentFormat.OpenXml.Drawing.Charts.Values values3;
                            if (graphGroup.Type == "BES" || graphGroup.Type == "BEFS" || graphGroup.Type == "BEFF" || graphGroup.Type == "BEF")
                            {
                                values3 = lcs3.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();
                            }
                            else
                            {
                                values3 = bcs3.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();
                            }
                           
                           
                            if (graphGroup.Type == "DQ" || graphGroup.Type == "DQF" || graphGroup.Type == "BEFS")
                            {
                                NumberReference vnr3 = values3.Descendants<NumberReference>().First();
                                //vnr3.Formula.Text = String.Format("Sheet1!$D$2:$D${0}", list[j].Data.Count);

                                nc3 = values3.Descendants<NumberingCache>().First();
                                CreateNumericPoints(nc3, list[j].Data.Count,true);
                            }

                            //Series 4
                            DocumentFormat.OpenXml.Drawing.Charts.Values values4;

                            if (graphGroup.Type == "BES" || graphGroup.Type == "BEFS" || graphGroup.Type == "BEFF" || graphGroup.Type == "BEF")
                            {
                                values4 = lcs4.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();
                            }

                            else
                            {
                                values4 = bcs4.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();
                            }

                            

                            if (graphGroup.Type == "FDQ" || graphGroup.Type == "FDQF" || graphGroup.Type == "BEFS")
                            {
                                NumberReference vnr4 = values4.Descendants<NumberReference>().First();
                                //vnr4.Formula.Text = String.Format("Sheet1!$E$2:$E${0}", list[j].Data.Count);

                                nc4 = values4.Descendants<NumberingCache>().First();
                                CreateNumericPoints(nc4, list[j].Data.Count,true);
                            }

                            if (graphGroup.Type == "BES" || graphGroup.Type == "BEFS" || graphGroup.Type == "BEFF" || graphGroup.Type == "BEF")
                            {
                                goto xy;
                            }
                          

                            if (graphGroup.Type == "CDQ" || graphGroup.Type == "CDQF")
                            {
                                //Series 5
                                var values5 = lcs1.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();
                                NumberReference vnr5 = values5.Descendants<NumberReference>().First();
                                //vnr5.Formula.Text = String.Format("Sheet1!$F$2:$F${0}", list[j].Data.Count);

                                nc5 = values5.Descendants<NumberingCache>().First();
                                CreateNumericPoints(nc5, list[j].Data.Count, true);
                            }

                           

                            if (graphGroup.Type == "FCDQ" || graphGroup.Type == "FCDQF")
                            {
                                //Series 6
                                var values6 = lcs2.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Values>().First();

                                NumberReference vnr6 = values6.Descendants<NumberReference>().First();
                                //vnr6.Formula.Text = String.Format("Sheet1!$G$2:$G${0}", list[j].Data.Count);

                                nc6 = values6.Descendants<NumberingCache>().First();
                                CreateNumericPoints(nc6, list[j].Data.Count, true);
                            }


                           xy: for (int i = 0; i < graphGroup.Data.Count; i++)
                            {
                                try
                                {
                                   

                                    switch (graphGroup.Type)
                                    {
                                        case "CF":
                                        case "CS":
                                        case "BES":
                                        NumericValue nv1 = nc1.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                                        nv1.Text = graphGroup.Data[i].Count.ToString();
                                             NumericValue sv = sc.Elements<StringPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                                             sv.Text = graphGroup.Data[i].Title.ToString();
                                            break;
                                   
                                        case "FCF":
                                        case "FCS":
                                             case "BEF":
                                            NumericValue nv2 = nc2.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                                            nv2.Text = graphGroup.Data[i].Count.ToString(); break;
                                        case "DQF":
                                        case "DQ":
                                             case "BEFS":
                                            NumericValue nv3 = nc3.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                                            nv3.Text = graphGroup.Data[i].Count.ToString(); break;
                                        case "FDQF":
                                        case "FDQ":
                                        case "BEFF":
                                            NumericValue nv4 = nc4.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                                            nv4.Text = graphGroup.Data[i].Count.ToString(); break;
                                        case "CDQF":
                                        case "CDQ":
                                            NumericValue nv5 = nc5.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                                            nv5.Text = graphGroup.Data[i].Count.ToString(); break;
                                        case "FCDQF":
                                        case "FCDQ":
                                            NumericValue nv6 = nc6.Elements<NumericPoint>().ElementAt(i).Elements<NumericValue>().FirstOrDefault();
                                            nv6.Text = graphGroup.Data[i].Count.ToString(); break;
                                    }
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                        catch
                        {
                            continue;
                        }

                    }
                Repository.Utility.WriteLog("LoadChartData completed successfully", System.Diagnostics.EventLogEntryType.Information);
            }
        
        }
    }
