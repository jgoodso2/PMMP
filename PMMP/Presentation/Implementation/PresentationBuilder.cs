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
            var fs = File.OpenRead(Constants.TEMPLATE_FILE_LOCATION);
            var bytes = new byte[fs.Length];
            fs.Read(bytes, 0, bytes.Length);
            fs.Close();
            return bytes;
        }

        public object BuildDataFromDataSource(string projectGuid)
        {
            return TaskItemRepository.GetTaskGroups(projectGuid);
        }

        public MemoryStream CreateDocument(byte[] template, string projectUID)
        {

            var ms = new MemoryStream();
            ms.Write(template, 0, (int)template.Length);

            using (var oPDoc = PresentationDocument.Open(ms, true))
            {
                var oPPart = oPDoc.PresentationPart;
                int gridSlideIndex = GetSlideindexByTitle(oPPart, "Driving Path");
                int completedSlideIndex = GetSlideindexByTitle(oPPart, "Complete Tasks");
                var lateSlideIndex = GetSlideindexByTitle(oPPart, "Late Tasks"); ;
                var chartSlideIndex = GetSlideindexByTitle(oPPart, "Chart"); ;
                SlidePart gridSlidePart = null;
                SlidePart chartSlidePart = null;
                SlidePart lateSlidePart = null;
                SlidePart completedSlidePart = null;
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

                var taskData = TaskItemRepository.GetTaskGroups(projectUID);
                var data = taskData.TaskItemGroups;
               
                int[] dynamicSlideIndices = { gridSlideIndex, completedSlideIndex};
                int[] fixedSlideIndices = { lateSlideIndex, chartSlideIndex  };
                
                Array.Sort(dynamicSlideIndices);
                Array.Sort(fixedSlideIndices);
                
                if (dynamicSlideIndices[0] < fixedSlideIndices[0])
                {
                    CreateDynamicSlides(data,dynamicSlideIndices,gridSlideIndex,completedSlideIndex,gridSlidePart,completedSlidePart,oPPart);
                    CreateFixedSlides(taskData, fixedSlideIndices,lateSlideIndex,chartSlideIndex,lateSlidePart,chartSlidePart,oPPart);
                }
                else
                {
                    CreateFixedSlides(taskData, fixedSlideIndices, lateSlideIndex, chartSlideIndex, lateSlidePart, chartSlidePart, oPPart);
                    CreateDynamicSlides(data, dynamicSlideIndices, gridSlideIndex, completedSlideIndex, gridSlidePart, completedSlidePart, oPPart);
                }
               
                
                PresentationUtilities.DeleteSlide(oPDoc, 2);
                PresentationUtilities.DeleteSlide(oPDoc, 2);
                PresentationUtilities.DeleteSlide(oPDoc, 2);
                PresentationUtilities.DeleteSlide(oPDoc, 2);
                PresentationUtilities.MoveSlide(oPDoc, 2, PresentationUtilities.CountSlides(oPDoc) - 1);
                int lowestSLideIndex = dynamicSlideIndices[0];
                int createdCount = 0;
                if (dynamicSlideIndices[0] < fixedSlideIndices[0])
                {
                    CreateDynamicSlidesData(data, dynamicSlideIndices, gridSlideIndex, lowestSLideIndex, ref createdCount, completedSlideIndex, oPPart);
                    CreateFixedSlidesData(taskData, fixedSlideIndices, lateSlideIndex, chartSlideIndex, lowestSLideIndex, ref createdCount, oPPart);
                }
                else
                {
                    CreateFixedSlidesData(taskData, fixedSlideIndices, lateSlideIndex, chartSlideIndex, lowestSLideIndex, ref createdCount, oPPart);
                    CreateDynamicSlidesData(data, dynamicSlideIndices, gridSlideIndex, lowestSLideIndex, ref createdCount, completedSlideIndex, oPPart);
                }
                return ms;
            }
        }

        private void CreateFixedSlidesData(TaskGroupData taskData, int[] fixedSlideIndices, int lateSlideIndex, int chartSlideIndex, int lowestSLideIndex, ref int createdCount, PresentationPart oPPart)
        {
            foreach (int slideIndex in fixedSlideIndices)
            {
                if (slideIndex == lateSlideIndex)
                {
                    CreateLateSLides(taskData, lowestSLideIndex, ref createdCount, oPPart);
                }

                if (slideIndex == chartSlideIndex)
                {
                    CreateChartSlides(taskData, lowestSLideIndex, ref createdCount, oPPart);
                }
            }
        }

        private void CreateDynamicSlidesData(IList<TaskItemGroup> data, int[] dynamicSlideIndices, int gridSlideIndex, int lowestSLideIndex, ref int createdCount, int completedSlideIndex, PresentationPart oPPart)
        {
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
        }

        private void CreateChartSlides(TaskGroupData taskData, int lowestSLideIndex, ref int createdCount, PresentationPart oPPart)
        {
            try
            {
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
                        catch
                        {
                            continue;
                        }



                    }

                }
            }
            catch
            {
            }
        }

        private void CreateLateSLides(TaskGroupData taskData, int lowestSLideIndex, ref int createdCount, PresentationPart oPPart)
        {
            try
            {
                #region LateSlides

                foreach (TaskItemGroup lateTaskgroup in taskData.LateTaskGroups)
                {
                    try
                    {
                        lateSlidePart = oPPart.GetSlidePartsInOrder().ToList()[lowestSLideIndex + createdCount];
                        var table = lateSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Table>().FirstOrDefault();

                        if (table != null && lateTaskgroup.TaskItems.Count > 0)
                        {
                            TableUtilities.PopulateLateTasksTable(table, lateTaskgroup.TaskItems, taskData.FiscalPeriod);
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
            catch
            {

            }
        }

        private void CreateCompletedSlides(TaskItemGroup group,int completedSlideIndex,PresentationPart oPPart)
        {
            try
            {
                if (group.CompletedTaskgroups != null)
                {
                    IList<TaskItemGroup> CompletedTasks = group.CompletedTaskgroups;
                    if (CompletedTasks != null && CompletedTasks.Count() > 0)
                    {
                        foreach (TaskItemGroup completedTaskgroup in CompletedTasks)
                        {
                            SlidePart completedSlidePart = oPPart.GetSlidePartsInOrder().ToList()[completedSlideIndex];
                            DocumentFormat.OpenXml.Presentation.TextBody numberPart = (DocumentFormat.OpenXml.Presentation.TextBody)completedSlidePart.Slide.CommonSlideData.ShapeTree.Elements().ToList()[3].Elements().ToList()[2];

                            foreach (string task in completedTaskgroup.TaskItems.Select(t => t.Task))
                            {
                                DocumentFormat.OpenXml.Drawing.Paragraph paragraph = (DocumentFormat.OpenXml.Drawing.Paragraph)numberPart.Elements().ToList()[3].Clone();
                                (paragraph.Elements().ToList()[0] as DocumentFormat.OpenXml.Drawing.Run).Text = new DocumentFormat.OpenXml.Drawing.Text(task);
                                numberPart.Append(paragraph);
                            }

                            var titleShape = completedSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().ToList();
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

                            DocumentFormat.OpenXml.Presentation.TextBody numPart = (DocumentFormat.OpenXml.Presentation.TextBody)completedSlidePart.Slide.CommonSlideData.ShapeTree.Elements().ToList()[3].Elements().ToList()[2];
                            ((DocumentFormat.OpenXml.Drawing.Paragraph)numPart.Elements().ToList()[3]).Remove();
                            ((DocumentFormat.OpenXml.Drawing.Paragraph)numPart.Elements().ToList()[3]).Remove();
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void CreateGridSlide(PresentationPart oPPart,int gridSlideIndex,int i,TaskItemGroup group)
        {
            try
            {
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
            catch
            {
            }
        }

        private void CreateFixedSlides(TaskGroupData taskData, int[] dynamicSlideIndices,int lateSlideIndex,int chartSlideIndex,SlidePart lateSlidePart,SlidePart chartSlidePart,PresentationPart oPPart)
        {
            foreach (int slideIndex in dynamicSlideIndices)
            {
                    if (lateSlideIndex == slideIndex)
                    {
                        if (taskData.LateTaskGroups != null)
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
                    }

                    if (chartSlideIndex == slideIndex)
                    {
                        if (taskData.ChartsData != null)
                        {
                            if (chartSlidePart != null)
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
                        }
                    }
                }
        }

        private void CreateDynamicSlides(IList<TaskItemGroup> data, int[] dynamicSlideIndices,int gridSlideIndex,int completedSlideIndex,SlidePart gridSlidePart,SlidePart completedSlidePart,PresentationPart oPPart)
        {
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
        }

        private int GetSlideindexByTitle(PresentationPart oPPart, string title)
        {
            List<SlidePart> slides = oPPart.GetSlidePartsInOrder().ToList();
            int count = 0;
            foreach (var gridSlidePart in slides)
            {
                var titleShape = gridSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().ToList();
                if (titleShape.Count > 0)
                {
                    if (titleShape[0].TextBody.InnerText.Trim().ToUpper() == title.Trim().ToUpper())
                    {
                        return count;
                    }
                }
                count++;
            }
            return -1;
        }

        public SlidePart lateSlidePart { get; set; }
    }
}
