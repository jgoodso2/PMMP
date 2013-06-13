using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;


namespace PMMP
{
    public class PresentationBuilder : IBuilder
    {
        public string FileName { get; set; }

        public byte[] OpenFile(string fileName)
        {
            Repository.Utility.WriteLog("OpenFile started ", System.Diagnostics.EventLogEntryType.Information);
            var fs = File.OpenRead(Constants.TEMPLATE_FILE_LOCATION);
            var bytes = new byte[fs.Length];
            fs.Read(bytes, 0, bytes.Length);
            fs.Close();
            Repository.Utility.WriteLog("OpenFile completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return bytes;
        }

        public object BuildDataFromDataSource(string projectGuid)
        {
            Repository.Utility.WriteLog("BuildDataFromDataSource started", System.Diagnostics.EventLogEntryType.Information);
            object data = TaskItemRepository.GetTaskGroups(projectGuid);
            Repository.Utility.WriteLog("BuildDataFromDataSource completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return data;
        }

        public MemoryStream CreateDocument(byte[] template, string projectUID)
        {
            Repository.Utility.WriteLog("CreateDocument started", System.Diagnostics.EventLogEntryType.Information);
            var ms = new MemoryStream();
            ms.Write(template, 0, (int)template.Length);

            using (var oPDoc = PresentationDocument.Open(ms, true))
            {
                var oPPart = oPDoc.PresentationPart;
                int gridSlideIndex = GetSlideindexByTitle(oPPart, "Driving Path");
                int completedSlideIndex = GetSlideindexByTitle(oPPart, "Completed Driving Path Tasks (current Fiscal Month)");
                var lateSlideIndex = GetSlideindexByTitle(oPPart, "Late Tasks (current Fiscal Month)");
                var upcomingSlideIndex = GetSlideindexByTitle(oPPart, "Upcoming Tasks (current Fiscal Month)"); 
                var chartSlideIndex = GetSlideindexByTitle(oPPart, "Chart");
                var SPDLSTartToBLSlideIndex = GetSlideindexByTitle(oPPart, "Schedule Performance – Delinquent Starts to BL");
                var SPDLFinishToBLSlideIndex = GetSlideindexByTitle(oPPart, "Schedule Performance – Delinquent Finishes to BL");
                var BEISlideIndex = GetSlideindexByTitle(oPPart, "Schedule Performance – Baseline Execution Index (BEI)"); 

                SlidePart gridSlidePart = null;
                SlidePart chartSlidePart = null;
                SlidePart lateSlidePart = null;
                SlidePart upcomingSlidePart = null;
                SlidePart completedSlidePart = null;
                SlidePart SPDLSTartToBLSlidePart = null;
                SlidePart SPDLFinishToBLSlidePart = null;
                SlidePart BEISlideIndexPart = null;
                if (gridSlideIndex > -1)
                {
                    gridSlidePart = oPPart.GetSlidePartsInOrder().ToList()[gridSlideIndex];
                }
                if (completedSlideIndex > -1)
                {
                    completedSlidePart = oPPart.GetSlidePartsInOrder().ToList()[completedSlideIndex];
                }
                if (chartSlideIndex > -1)
                {
                    chartSlidePart = oPPart.GetSlidePartsInOrder().ToList()[chartSlideIndex];
                }
                if (lateSlideIndex > -1)
                {
                    lateSlidePart = oPPart.GetSlidePartsInOrder().ToList()[lateSlideIndex];
                }
                if (upcomingSlideIndex > -1)
                {
                    upcomingSlidePart = oPPart.GetSlidePartsInOrder().ToList()[upcomingSlideIndex];
                }

                if (SPDLSTartToBLSlideIndex > -1)
                {
                    SPDLSTartToBLSlidePart = oPPart.GetSlidePartsInOrder().ToList()[SPDLSTartToBLSlideIndex];
                }

                if (SPDLFinishToBLSlideIndex > -1)
                {
                    SPDLFinishToBLSlidePart = oPPart.GetSlidePartsInOrder().ToList()[SPDLFinishToBLSlideIndex];
                }

                if (BEISlideIndex > -1)
                {
                    BEISlideIndexPart = oPPart.GetSlidePartsInOrder().ToList()[BEISlideIndex];
                }

                DPSlidePart =  gridSlidePart;
                ChartSlidePart = chartSlidePart;
                LateSlidePart =  lateSlidePart;
                UpComingSlidePart = upcomingSlidePart;
                CPSlidePart = completedSlidePart;
                SPDSSlidePart =SPDLSTartToBLSlidePart;
                SPDFSlidePart = SPDLFinishToBLSlidePart;
                SPBEISlidePart = BEISlideIndexPart;
                var taskData = TaskItemRepository.GetTaskGroups(projectUID);
                var data = taskData.TaskItemGroups;

                int[] dynamicSlideIndices = { gridSlideIndex, completedSlideIndex };
                int[] fixedSlideIndices = { lateSlideIndex, upcomingSlideIndex, chartSlideIndex, SPDLSTartToBLSlideIndex, SPDLFinishToBLSlideIndex, BEISlideIndex };

                dynamicSlideIndices = dynamicSlideIndices.Where(t => t != -1).ToArray();
                fixedSlideIndices = fixedSlideIndices.Where(t => t != -1).ToArray();
                Array.Sort(dynamicSlideIndices);
                Array.Sort(fixedSlideIndices);

                if (dynamicSlideIndices[0] < fixedSlideIndices[0])
                {
                    CreateDynamicSlides(data, dynamicSlideIndices, gridSlideIndex, completedSlideIndex, gridSlidePart, completedSlidePart, oPPart);
                    CreateFixedSlides(taskData, fixedSlideIndices, lateSlideIndex,upcomingSlideIndex, chartSlideIndex,SPDLSTartToBLSlideIndex,SPDLFinishToBLSlideIndex,BEISlideIndex, lateSlidePart,upcomingSlidePart, chartSlidePart,SPDLSTartToBLSlidePart,SPDLFinishToBLSlidePart,BEISlideIndexPart, oPPart);
                }
                else
                {
                    CreateFixedSlides(taskData, fixedSlideIndices, lateSlideIndex, upcomingSlideIndex, chartSlideIndex, SPDLSTartToBLSlideIndex, SPDLFinishToBLSlideIndex, BEISlideIndex, lateSlidePart, upcomingSlidePart, chartSlidePart, SPDLSTartToBLSlidePart, SPDLFinishToBLSlidePart, BEISlideIndexPart, oPPart);
                    CreateDynamicSlides(data, dynamicSlideIndices, gridSlideIndex, completedSlideIndex, gridSlidePart, completedSlidePart, oPPart);
                }

                int lowestSLideIndex;
                if (dynamicSlideIndices[0] < fixedSlideIndices[0])
                {
                    lowestSLideIndex = dynamicSlideIndices[0];
                }
                else
                {
                    lowestSLideIndex = fixedSlideIndices[0];
                }
                for (int i = 0; i < (dynamicSlideIndices.Count() + fixedSlideIndices.Count()); i++)
                {
                    PresentationUtilities.DeleteSlide(oPDoc, lowestSLideIndex);
                }

                PresentationUtilities.MoveSlide(oPDoc, lowestSLideIndex, PresentationUtilities.CountSlides(oPDoc) - 1);
                
                int createdCount = 0;
                if (dynamicSlideIndices[0] < fixedSlideIndices[0])
                {
                    CreateDynamicSlidesData(data, dynamicSlideIndices, gridSlideIndex, lowestSLideIndex, ref createdCount, completedSlideIndex, oPPart);
                    CreateFixedSlidesData(taskData, fixedSlideIndices, lateSlideIndex,upcomingSlideIndex, chartSlideIndex, lowestSLideIndex, SPDLSTartToBLSlideIndex, SPDLFinishToBLSlideIndex,BEISlideIndex, ref createdCount, oPPart);
                }
                else
                {
                    CreateFixedSlidesData(taskData, fixedSlideIndices, lateSlideIndex,upcomingSlideIndex, chartSlideIndex, lowestSLideIndex,SPDLSTartToBLSlideIndex,SPDLFinishToBLSlideIndex,BEISlideIndex, ref createdCount, oPPart);
                    CreateDynamicSlidesData(data, dynamicSlideIndices, gridSlideIndex, lowestSLideIndex, ref createdCount, completedSlideIndex, oPPart);
                }
                Repository.Utility.WriteLog("CreateDocument completed successfully", System.Diagnostics.EventLogEntryType.Information);
                return ms;
            }
        }

        private void CreateFixedSlidesData(TaskGroupData taskData, int[] fixedSlideIndices, int lateSlideIndex,int upcomingSlideIndex, int chartSlideIndex, int lowestSLideIndex, int SPDLSTartToBLSlideIndex, int SPDLFinishToBLSlideIndex,int BEISlideIndex,ref int createdCount, PresentationPart oPPart)
        {
            Repository.Utility.WriteLog("CreateFixedSlidesData started", System.Diagnostics.EventLogEntryType.Information);
            foreach (int slideIndex in fixedSlideIndices)
            {
                if (slideIndex == lateSlideIndex)
                {
                    CreateLateSlides(taskData, lowestSLideIndex, ref createdCount, oPPart);
                }
                if (slideIndex == upcomingSlideIndex)
                {
                    CreateUpComingSlides(taskData, lowestSLideIndex, ref createdCount, oPPart);
                }
                if (slideIndex == chartSlideIndex)
                {
                    CreateChartSlides(taskData, lowestSLideIndex, ref createdCount, oPPart);
                }
                if (slideIndex == SPDLSTartToBLSlideIndex)
                {
                    CreateSPDLToBLSlides(taskData, lowestSLideIndex, ref createdCount, oPPart,ChartType.SPBaseLineStart);
                }
                if (slideIndex == SPDLFinishToBLSlideIndex)
                {
                    CreateSPDLToBLSlides(taskData, lowestSLideIndex, ref createdCount, oPPart, ChartType.SPBaseLineFinish);
                }
                if (slideIndex == BEISlideIndex)
                {
                    CreateSPDLToBLSlides(taskData, lowestSLideIndex, ref createdCount, oPPart, ChartType.SPBEI);
                }
                
            }
            Repository.Utility.WriteLog("CreateFixedSlidesData completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        private void CreateUpComingSlides(TaskGroupData taskData, int lowestSLideIndex, ref int createdCount, PresentationPart oPPart)
        {
            Repository.Utility.WriteLog("CreateupComingSlides started", System.Diagnostics.EventLogEntryType.Information);
            try
            {
                if (UpComingSlidePart == null)
                    return;
                #region upComingSlides

                if (taskData.UpComingTaskGroups == null || taskData.UpComingTaskGroups.Count == 0)
                {
                    createdCount++;
                    return;
                }

                foreach (TaskItemGroup upComingTaskgroup in taskData.UpComingTaskGroups)
                {
                    try
                    {
                        SlidePart upComingSlidePart = oPPart.GetSlidePartsInOrder().ToList()[lowestSLideIndex + createdCount];
                        var table = upComingSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Table>().FirstOrDefault();

                        if (table != null && upComingTaskgroup.TaskItems.Count > 0)
                        {
                            TableUtilities.PopulateLateOrUpComingTasksTable(table, upComingTaskgroup.TaskItems, taskData.FiscalPeriod);
                        }
                        createdCount++;
                    }
                    catch
                    {
                        continue;
                    }
                }

                #endregion
            }
            catch (Exception ex)
            {
                Repository.Utility.WriteLog(string.Format("CreateupComingSlides had an error and the error message={0}", ex.Message), System.Diagnostics.EventLogEntryType.Information);
            }
            Repository.Utility.WriteLog("CreateupComingSlides completed", System.Diagnostics.EventLogEntryType.Information);
        }

       

        private void CreateSPDLToBLSlides(TaskGroupData taskData, int lowestSLideIndex, ref int createdCount, PresentationPart oPPart, ChartType chartType)
        {
            Repository.Utility.WriteLog("CreateLateSlides started", System.Diagnostics.EventLogEntryType.Information);
            try
            {
                #region LateSlides
                if (chartType == ChartType.SPBaseLineStart && SPDSSlidePart == null)
                    return;
                if (chartType == ChartType.SPBaseLineFinish && SPDFSlidePart == null)
                    return;
                if (chartType == ChartType.SPBEI && SPBEISlidePart == null)
                    return;
                SlidePart chartSlidePart = oPPart.GetSlidePartsInOrder().ToList()[lowestSLideIndex + createdCount];

                if (chartSlidePart.ChartParts.ToList().Count > 0)
                {
                   
                    var chartPart = chartSlidePart.ChartParts.ToList()[0];

                    foreach (IdPartPair part in chartPart.Parts)
                    {
                        var spreadsheet = chartPart.GetPartById(part.RelationshipId) as EmbeddedPackagePart;

                        if (spreadsheet != null)
                        {
                            using (var oSDoc = SpreadsheetDocument.Open(spreadsheet.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite), true))
                            {
                                var workSheetPart = oSDoc.WorkbookPart.GetPartsOfType<WorksheetPart>().FirstOrDefault();
                                var sheetData = workSheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                                switch (chartType)
                                {
                                    case ChartType.SPBaseLineStart:
                                        if (taskData.SPDLSTartToBL != null && taskData.SPDLSTartToBL.Count > 0 && taskData.SPDLSTartToBL[0].Data != null)
                                        {
                                            WorkbookUtilities.ReplicateRow(sheetData, 2, taskData.SPDLSTartToBL[0].Data.Count - 1);
                                            WorkbookUtilities.LoadGraphSheetData(sheetData, taskData.SPDLSTartToBL, 1, 0);
                                            BarChartUtilities.LoadChartData(chartPart, taskData.SPDLSTartToBL);
                                        }
                                        break;
                                    case ChartType.SPBaseLineFinish:
                                        if (taskData.SPDLFinishToBL != null && taskData.SPDLFinishToBL.Count > 0 && taskData.SPDLFinishToBL[0].Data != null)
                                        {
                                            WorkbookUtilities.ReplicateRow(sheetData, 2, taskData.SPDLFinishToBL[0].Data.Count - 1);
                                            WorkbookUtilities.LoadGraphSheetData(sheetData, taskData.SPDLFinishToBL, 1, 0);
                                            BarChartUtilities.LoadChartData(chartPart, taskData.SPDLFinishToBL);
                                        }
                                        break;
                                    case ChartType.SPBEI:
                                        if (taskData.BEIData != null && taskData.BEIData.Count > 0 && taskData.BEIData[0].Data != null)
                                        {
                                            WorkbookUtilities.ReplicateRow(sheetData, 2, taskData.BEIData[0].Data.Count - 1);
                                            WorkbookUtilities.LoadGraphSheetData(sheetData, taskData.BEIData, 1, 0);
                                            BarChartUtilities.LoadChartData(chartPart, taskData.BEIData);
                                        }
                                        break;
                                }
                            }

                            break;
                        }
                    }
                }

                #endregion
            }
            catch (Exception ex)
            {
                Repository.Utility.WriteLog(string.Format("CreateLateSlides had an error and the error message={0}", ex.Message), System.Diagnostics.EventLogEntryType.Information);
            }
            createdCount++;
            Repository.Utility.WriteLog("CreateLateSlides completed", System.Diagnostics.EventLogEntryType.Information);
        }

        private void CreateDynamicSlidesData(IList<TaskItemGroup> data, int[] dynamicSlideIndices, int gridSlideIndex, int lowestSLideIndex, ref int createdCount, int completedSlideIndex, PresentationPart oPPart)
        {
            Repository.Utility.WriteLog("CreateDynamicSlidesData started", System.Diagnostics.EventLogEntryType.Information);
            if (data.Count == 0)
            {
                createdCount += 2;
                return;
            }
            for (int i = 0; i < data.Count; i++)
            {
                var group = data[i];
                foreach (int slideIndex in dynamicSlideIndices)
                {
                    if (slideIndex == gridSlideIndex)
                    {
                        CreateGridSlide(oPPart, lowestSLideIndex, createdCount, group);
                        createdCount++;
                    }

                    if (slideIndex == completedSlideIndex)
                    {
                        CreateCompletedSlides(group, lowestSLideIndex + createdCount, oPPart);
                        createdCount++;
                    }
                }
            }
            Repository.Utility.WriteLog("CreateDynamicSlidesData completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        private void CreateChartSlides(TaskGroupData taskData, int lowestSLideIndex, ref int createdCount, PresentationPart oPPart)
        {
            Repository.Utility.WriteLog("CreateChartSlides started", System.Diagnostics.EventLogEntryType.Information);
            try
            {
                if (ChartSlidePart == null)
                    return;
                if (taskData.ChartsData == null || taskData.ChartsData.Keys.Count == 0 || !taskData.ChartsData.Keys.Any(t=>t.StartsWith("Show On")))
                {
                    createdCount++;
                    return;
                }
                
                if (taskData.ChartsData != null)
                {
                    foreach (string key in taskData.ChartsData.Keys)
                    {
                        try
                        {
                            if (key.StartsWith("Show On"))
                            {
                                //Get all Tasks related to  Driving path
                                TaskItemGroup newGroup = new TaskItemGroup() { ChartTaskItems = taskData.ChartsData[key] };
                                var chartDataTable = newGroup.GetChartDataTable(key);

                                #region Charts
                                if (chartDataTable != null)
                                {

                                    SlidePart chartSlidePart = oPPart.GetSlidePartsInOrder().ToList()[lowestSLideIndex + createdCount];

                                    if (chartSlidePart.ChartParts.ToList().Count > 0)
                                    {
                                        createdCount++;
                                        var chartPart = chartSlidePart.ChartParts.ToList()[0];

                                        foreach (IdPartPair part in chartPart.Parts)
                                        {
                                            var spreadsheet = chartPart.GetPartById(part.RelationshipId) as EmbeddedPackagePart;

                                            if (spreadsheet != null)
                                            {
                                                using (var oSDoc = SpreadsheetDocument.Open(spreadsheet.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite), true))
                                                {
                                                    var workSheetPart = oSDoc.WorkbookPart.GetPartsOfType<WorksheetPart>().FirstOrDefault();
                                                    var sheetData = workSheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                                                    WorkbookUtilities.ReplicateRow(sheetData, 2, chartDataTable.Rows.Count - 1);
                                                    WorkbookUtilities.LoadSheetData(sheetData, chartDataTable, 1, 0);
                                                    BarChartUtilities.LoadChartData(chartPart, chartDataTable);
                                                }

                                                break;
                                            }
                                        }

                                        var titleShape = chartSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().ToList();
                                        if (titleShape.Count > 0)
                                        {
                                            titleShape[0].TextBody = new DocumentFormat.OpenXml.Presentation.TextBody(
                                                                  new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                                                                  new DocumentFormat.OpenXml.Drawing.ListStyle(),
                                                                  new DocumentFormat.OpenXml.Drawing.Paragraph(
                                                                  new DocumentFormat.OpenXml.Drawing.Run(
                                                                  new DocumentFormat.OpenXml.Drawing.RunProperties() { FontSize = 3600 },
                                                                  new DocumentFormat.OpenXml.Drawing.Text { Text = key.Replace("Show On_", "") })));
                                        }

                                    }
                                #endregion
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Repository.Utility.WriteLog(string.Format("CreateChartSlides had an error and the error message={0}", ex.Message), System.Diagnostics.EventLogEntryType.Information);
                            continue;
                        }



                    }

                }
            }
            catch (Exception ex)
            {
                Repository.Utility.WriteLog(string.Format("CreateChartSlides had an error and the error message={0}", ex.Message), System.Diagnostics.EventLogEntryType.Information);
            }

            Repository.Utility.WriteLog("CreateChartSlides completed", System.Diagnostics.EventLogEntryType.Information);
        }

        private void CreateLateSlides(TaskGroupData taskData, int lowestSLideIndex, ref int createdCount, PresentationPart oPPart)
        {
            Repository.Utility.WriteLog("CreateLateSlides started", System.Diagnostics.EventLogEntryType.Information);
            try
            {
                if (LateSlidePart == null)
                    return;
                #region LateSlides

                if (taskData.LateTaskGroups == null || taskData.LateTaskGroups.Count == 0)
                {
                    createdCount++;
                    return;
                }

                foreach (TaskItemGroup lateTaskgroup in taskData.LateTaskGroups)
                {
                    try
                    {
                        SlidePart lateSlidePart = oPPart.GetSlidePartsInOrder().ToList()[lowestSLideIndex + createdCount];
                        var table = lateSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Table>().FirstOrDefault();

                        if (table != null && lateTaskgroup.TaskItems.Count > 0)
                        {
                            TableUtilities.PopulateLateOrUpComingTasksTable(table, lateTaskgroup.TaskItems, taskData.FiscalPeriod);
                        }
                        createdCount++;
                    }
                    catch
                    {
                        continue;
                    }
                }

                #endregion
            }
            catch (Exception ex)
            {
                Repository.Utility.WriteLog(string.Format("CreateLateSlides had an error and the error message={0}", ex.Message), System.Diagnostics.EventLogEntryType.Information);
            }
            Repository.Utility.WriteLog("CreateLateSlides completed", System.Diagnostics.EventLogEntryType.Information);
        }

        private void CreateCompletedSlides(TaskItemGroup group, int completedSlideIndex, PresentationPart oPPart)
        {
            Repository.Utility.WriteLog("CreateCompletedSlides started", System.Diagnostics.EventLogEntryType.Information);
            try
            {
                if (CPSlidePart == null)
                    return;
                if (group.CompletedTaskgroups != null)
                {
                    IList<TaskItemGroup> CompletedTasks = group.CompletedTaskgroups;
                    if (CompletedTasks != null && CompletedTasks.Count() > 0)
                    {
                        foreach (TaskItemGroup completedTaskgroup in CompletedTasks)
                        {
                            SlidePart completedSlidePart = oPPart.GetSlidePartsInOrder().ToList()[completedSlideIndex];
                            DocumentFormat.OpenXml.Presentation.TextBody numberPart = (DocumentFormat.OpenXml.Presentation.TextBody)completedSlidePart.Slide.CommonSlideData.ShapeTree.Elements().ToList()[3].Elements().ToList()[2];
                            int bulletCount = numberPart.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>().Count();
                            foreach (string task in completedTaskgroup.TaskItems.Select(t => t.Task))
                            {
                                
                                DocumentFormat.OpenXml.Drawing.Paragraph paragraph = (DocumentFormat.OpenXml.Drawing.Paragraph)numberPart.Elements().ToList()[3].Clone();
                                for (int i = paragraph.Elements().Count() - 1; i >=1; i--)
                                {
                                    if (paragraph.ElementAt(i).GetType() == typeof(DocumentFormat.OpenXml.Drawing.Run))
                                    {
                                        paragraph.ElementAt(i).Remove();
                                    }
                                }
                                (paragraph.Elements().ToList()[0] as DocumentFormat.OpenXml.Drawing.Run).Text = new DocumentFormat.OpenXml.Drawing.Text(task);
                                numberPart.Append(paragraph);
                            }

                            //var titleShape = completedSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().ToList();
                            //if (titleShape.Count > 0)
                            //{
                            //    titleShape[0].TextBody = new DocumentFormat.OpenXml.Presentation.TextBody(
                            //                          new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                            //                          new DocumentFormat.OpenXml.Drawing.ListStyle(),
                            //                          new DocumentFormat.OpenXml.Drawing.Paragraph(
                            //                          new DocumentFormat.OpenXml.Drawing.Run(
                            //                          new DocumentFormat.OpenXml.Drawing.RunProperties() { FontSize = 3600 },
                            //                          new DocumentFormat.OpenXml.Drawing.Text { Text = group.Title })));
                            //}

                            DocumentFormat.OpenXml.Presentation.TextBody numPart = (DocumentFormat.OpenXml.Presentation.TextBody)completedSlidePart.Slide.CommonSlideData.ShapeTree.Elements().ToList()[3].Elements().ToList()[2];
                            for (int i = bulletCount - 2; i >= 0; i--)
                            {
                                ((DocumentFormat.OpenXml.Drawing.Paragraph)numPart.Elements().ToList()[3]).Remove();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Repository.Utility.WriteLog(string.Format("CreateCompletedSlides had an error and the error message={0}", ex.Message), System.Diagnostics.EventLogEntryType.Information);
            }
            Repository.Utility.WriteLog("CreateCompletedSlides completed", System.Diagnostics.EventLogEntryType.Information);
        }

        private void CreateGridSlide(PresentationPart oPPart, int gridSlideIndex, int i, TaskItemGroup group)
        {
            Repository.Utility.WriteLog("CreateGridSlide started", System.Diagnostics.EventLogEntryType.Information);
            try
            {
                if (DPSlidePart == null)
                    return;
                SlidePart gridSlidePart = oPPart.GetSlidePartsInOrder().ToList()[gridSlideIndex + i];

                var dataTable = group.TaskItemsDataTable;
                var table = gridSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Table>().FirstOrDefault();

                if (table != null && group.TaskItems != null && group.TaskItems.Count > 0)
                {
                    TableUtilities.PopulateTable(table, group.TaskItems);
                }
                //else
                //{
                //    DocumentFormat.OpenXml.OpenXmlElement parent = table.Parent;
                //    parent.ReplaceChild(new DocumentFormat.OpenXml.Presentation.Text("No Data Avialable"), table);
                //    //table.Remove();
                //}

                var titleShape = gridSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().ToList();
                if (titleShape.Count > 0)
                {
                    titleShape[0].TextBody = new DocumentFormat.OpenXml.Presentation.TextBody(
                                          new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                                          new DocumentFormat.OpenXml.Drawing.ListStyle(),
                                          new DocumentFormat.OpenXml.Drawing.Paragraph(
                                          new DocumentFormat.OpenXml.Drawing.Run(
                                          new DocumentFormat.OpenXml.Drawing.RunProperties() { FontSize = 3600 },
                                          new DocumentFormat.OpenXml.Drawing.Text { Text = group.Title })));
                }
            }
            catch (Exception ex)
            {
                Repository.Utility.WriteLog(string.Format("CreateGridSlide had an error and the error message={0}", ex.Message), System.Diagnostics.EventLogEntryType.Information);
            }
            Repository.Utility.WriteLog("CreateGridSlide completed", System.Diagnostics.EventLogEntryType.Information);
        }

        private void CreateFixedSlides(TaskGroupData taskData, int[] dynamicSlideIndices, int lateSlideIndex,int upComingSlideIndex, int chartSlideIndex, int SPDLSTartToBLSlideIndex, int SPDLFinishToBLSlideIndex,int BEISlideIndex, SlidePart lateSlidePart,SlidePart upComingSlidePart, SlidePart chartSlidePart, SlidePart SPDLSTartToBLSlidePart, SlidePart SPDLFinishToBLSlidePart,SlidePart BEISlidePart, PresentationPart oPPart)
        {
            Repository.Utility.WriteLog("CreateFixedSlides started", System.Diagnostics.EventLogEntryType.Information);
            foreach (int slideIndex in dynamicSlideIndices)
            {
                if (lateSlideIndex == slideIndex)
                {
                    if (taskData.LateTaskGroups != null && taskData.LateTaskGroups.Count > 0)
                    {
                        if (lateSlidePart != null)
                        {
                            foreach (TaskItemGroup group in taskData.LateTaskGroups)
                            {
                                var newLateSlidePart = lateSlidePart.CloneSlide(SlideType.Late);
                                oPPart.AppendSlide(newLateSlidePart);
                            }
                        }
                    }
                    else
                    {
                        if (lateSlidePart != null)
                        {
                            var newLateSlidePart = lateSlidePart.CloneSlide(SlideType.Late);
                            oPPart.AppendSlide(newLateSlidePart);
                        }
                    }
                }
                if (upComingSlideIndex == slideIndex)
                {
                    if (taskData.UpComingTaskGroups != null && taskData.UpComingTaskGroups.Count > 0)
                    {
                        if (upComingSlidePart != null)
                        {
                            foreach (TaskItemGroup group in taskData.UpComingTaskGroups)
                            {
                                var newupComingSlidePart = upComingSlidePart.CloneSlide(SlideType.Late);
                                oPPart.AppendSlide(newupComingSlidePart);
                            }
                        }
                    }
                    else
                    {
                        if (upComingSlidePart != null)
                        {
                            var newupComingSlidePart = upComingSlidePart.CloneSlide(SlideType.Late);
                            oPPart.AppendSlide(newupComingSlidePart);
                        }
                    }
                }

                if (SPDLSTartToBLSlideIndex == slideIndex)
                {
                        if (SPDLSTartToBLSlidePart != null)
                        {
                                var newSPDLStartToBLSlidePart = SPDLSTartToBLSlidePart.CloneSlide(SlideType.Chart);
                                oPPart.AppendSlide(newSPDLStartToBLSlidePart);
                        }
                }

                if (SPDLFinishToBLSlideIndex == slideIndex)
                {
                    if (SPDLFinishToBLSlidePart != null)
                    {
                        var newSPDLFinishToBLSlidePart = SPDLFinishToBLSlidePart.CloneSlide(SlideType.Chart);
                        oPPart.AppendSlide(newSPDLFinishToBLSlidePart);
                    }
                }

                if (BEISlideIndex == slideIndex)
                {
                    if (BEISlidePart != null)
                    {
                        var newBEISlidePart = BEISlidePart.CloneSlide(SlideType.Chart);
                        oPPart.AppendSlide(newBEISlidePart);
                    }
                }

                if (chartSlideIndex == slideIndex)
                {
                    if (taskData.ChartsData != null && taskData.ChartsData.Count > 0)
                    {
                        if (chartSlidePart != null && taskData.ChartsData.Keys.Any(t => t.StartsWith("Show On") == true))
                        {
                            foreach (string chartType in taskData.ChartsData.Keys)
                            {
                                if (chartType.StartsWith("Show On"))
                                {
                                    var newChartSlidePart = chartSlidePart.CloneSlide(SlideType.Chart);
                                    oPPart.AppendSlide(newChartSlidePart);
                                }
                            }
                        }
                        else
                        {
                            var newChartSlidePart = chartSlidePart.CloneSlide(SlideType.Chart);
                            oPPart.AppendSlide(newChartSlidePart);
                        }
                    }

                }
            }
            Repository.Utility.WriteLog("CreateFixedSlides completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        private void CreateDynamicSlides(IList<TaskItemGroup> data, int[] dynamicSlideIndices, int gridSlideIndex, int completedSlideIndex, SlidePart gridSlidePart, SlidePart completedSlidePart, PresentationPart oPPart)
        {
            Repository.Utility.WriteLog("CreateDynamicSlides started", System.Diagnostics.EventLogEntryType.Information);
            if (data.Count < 1)
            {
                foreach (int slideIndex in dynamicSlideIndices)
                {
                    if (slideIndex == gridSlideIndex)
                    {
                        var newGridSlidePart = gridSlidePart.CloneSlide(SlideType.Grid);
                        oPPart.AppendSlide(newGridSlidePart);
                    }
                    if (slideIndex == completedSlideIndex)
                    {
                        var newCompletedSlidePart = completedSlidePart.CloneSlide(SlideType.Completed);
                        oPPart.AppendSlide(newCompletedSlidePart);
                    }
                }
            }
            for (int i = 0; i < data.Count; i++)
            {
                foreach (int slideIndex in dynamicSlideIndices)
                {
                    if (slideIndex == gridSlideIndex)
                    {
                        if (gridSlidePart != null)
                        {
                            var newGridSlidePart = gridSlidePart.CloneSlide(SlideType.Grid);
                            oPPart.AppendSlide(newGridSlidePart);
                        }
                    }
                    if (slideIndex == completedSlideIndex)
                    {
                        if (data[i].CompletedTaskgroups != null)
                        {
                            if (completedSlidePart != null)
                            {
                                if (data[i].CompletedTaskgroups.Count == 0)
                                {
                                    var newCompletedSlidePart = completedSlidePart.CloneSlide(SlideType.Completed);
                                    oPPart.AppendSlide(newCompletedSlidePart);
                                }
                                else
                                {
                                    foreach (TaskItemGroup group in data[i].CompletedTaskgroups)
                                    {
                                        var newCompletedSlidePart = completedSlidePart.CloneSlide(SlideType.Completed);
                                        oPPart.AppendSlide(newCompletedSlidePart);
                                    }
                                }
                            }
                        }
                        else
                        {
                            var newCompletedSlidePart = completedSlidePart.CloneSlide(SlideType.Completed);
                            oPPart.AppendSlide(newCompletedSlidePart);
                        }
                    }
                }

            }
            Repository.Utility.WriteLog("CreateDynamicSlides completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        private int GetSlideindexByTitle(PresentationPart oPPart, string title)
        {
            Repository.Utility.WriteLog("GetSlideindexByTitle started", System.Diagnostics.EventLogEntryType.Information);
            List<SlidePart> slides = oPPart.GetSlidePartsInOrder().ToList();
            int count = 0;
            foreach (var gridSlidePart in slides)
            {
                var titleShape = gridSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().ToList();
                if (titleShape.Count > 0)
                {
                    if (titleShape[0].TextBody.InnerText.Trim().ToUpper() == title.Trim().ToUpper())
                    {
                        Repository.Utility.WriteLog("GetSlideindexByTitle completed successfully", System.Diagnostics.EventLogEntryType.Information);
                        return count;
                    }
                }
                count++;
            }
            Repository.Utility.WriteLog("GetSlideindexByTitle completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return -1;
        }

        public SlidePart DPSlidePart { get; set; }
        public SlidePart CPSlidePart { get; set; }
        public SlidePart LateSlidePart { get; set; }
        public SlidePart UpComingSlidePart { get; set; }
        public SlidePart ChartSlidePart { get; set; }
        public SlidePart SPDSSlidePart { get; set; }
        public SlidePart SPDFSlidePart { get; set; }
        public SlidePart SPBEISlidePart { get; set; }
    }

}
