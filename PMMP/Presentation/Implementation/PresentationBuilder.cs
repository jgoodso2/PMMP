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

        public MemoryStream CreateDocument(byte[] template,string projectUID)
        {

            var ms = new MemoryStream();
            ms.Write(template, 0, (int)template.Length);

            using (var oPDoc = PresentationDocument.Open(ms, true))
            {
                //Make indexes configurable
                var gridSlideIndex = 2;
                var completedSlideIndex = 3;
                var lateSlideIndex = 4;
                var chartSlideIndex = 5;
                var oPPart = oPDoc.PresentationPart;
                var gridSlidePart = oPPart.GetSlidePartsInOrder().ToList()[gridSlideIndex];
                var completedSlidePart = oPPart.GetSlidePartsInOrder().ToList()[completedSlideIndex];
                var chartSlidePart = oPPart.GetSlidePartsInOrder().ToList()[chartSlideIndex];
                var lateSlidePart = oPPart.GetSlidePartsInOrder().ToList()[lateSlideIndex];
                var taskData = TaskItemRepository.GetTaskGroups(projectUID);
                var data = taskData.TaskItemGroups;
                int noOfCompletedSlides = 0,noOfGridSlides=0;
                for (int i = 0; i < data.Count; i++)
                {
                    var newGridSlidePart = gridSlidePart.CloneSlide(SlideType.Grid);
                    oPPart.AppendSlide(newGridSlidePart);
                    noOfGridSlides++;
                    if (data[i].CompletedTaskgroups != null)
                    {
                        foreach (TaskItemGroup group in data[i].CompletedTaskgroups)
                        {
                            var newCompletedSlidePart = completedSlidePart.CloneSlide(SlideType.Completed);
                            oPPart.AppendSlide(newCompletedSlidePart);
                            noOfCompletedSlides++;
                        }
                    }

                    
                }

                if (taskData.LateTaskGroups != null)
                {
                    foreach (TaskItemGroup group in taskData.LateTaskGroups)
                    {
                        var newLateSlidePart = lateSlidePart.CloneSlide(SlideType.Late);
                        oPPart.AppendSlide(newLateSlidePart);
                    }
                }

                if (taskData.ChartsData != null)
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
               
                

                PresentationUtilities.DeleteSlide(oPDoc, 2);
                PresentationUtilities.DeleteSlide(oPDoc, 2);
                PresentationUtilities.DeleteSlide(oPDoc, 2);
                PresentationUtilities.DeleteSlide(oPDoc, 2);
                PresentationUtilities.MoveSlide(oPDoc, 2, PresentationUtilities.CountSlides(oPDoc) - 1);

               
                var completedSlidesCreated = 0;
                for (int i = 0; i < data.Count; i++)
                {
                    gridSlidePart = oPPart.GetSlidePartsInOrder().ToList()[gridSlideIndex + completedSlidesCreated + i];
                    var group = data[i];
                    var dataTable = group.TaskItemsDataTable;
                    var table = gridSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Table>().FirstOrDefault();

                    if (table != null && group.TaskItems != null  && group.TaskItems.Count > 0)
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

                    if (data[i].CompletedTaskgroups != null)
                    {
                        IList<TaskItemGroup> CompletedTasks = data[i].CompletedTaskgroups;
                        if (CompletedTasks != null && CompletedTasks.Count() > 0)
                        {
                            foreach (TaskItemGroup completedTaskgroup in CompletedTasks)
                            {
                                completedSlidePart = oPPart.GetSlidePartsInOrder().ToList()[gridSlideIndex + completedSlidesCreated + i + 1];
                                DocumentFormat.OpenXml.Presentation.TextBody numberPart = (DocumentFormat.OpenXml.Presentation.TextBody)completedSlidePart.Slide.CommonSlideData.ShapeTree.Elements().ToList()[3].Elements().ToList()[2];

                                foreach (string task in completedTaskgroup.TaskItems.Select(t => t.Task))
                                {
                                    DocumentFormat.OpenXml.Drawing.Paragraph paragraph = (DocumentFormat.OpenXml.Drawing.Paragraph)numberPart.Elements().ToList()[3].Clone();
                                    (paragraph.Elements().ToList()[0] as DocumentFormat.OpenXml.Drawing.Run).Text = new DocumentFormat.OpenXml.Drawing.Text(task);
                                    numberPart.Append(paragraph);
                                }
                                
                                    titleShape = completedSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().ToList();
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
                                    completedSlidesCreated++;

                                    DocumentFormat.OpenXml.Presentation.TextBody numPart = (DocumentFormat.OpenXml.Presentation.TextBody)completedSlidePart.Slide.CommonSlideData.ShapeTree.Elements().ToList()[3].Elements().ToList()[2];
                                        ((DocumentFormat.OpenXml.Drawing.Paragraph)numPart.Elements().ToList()[3]).Remove();
                                        ((DocumentFormat.OpenXml.Drawing.Paragraph)numPart.Elements().ToList()[3]).Remove();
                            }
                        }
                    }

                    
                   
                }


                 #region LateSlides
                int lateSlidesCreated=0, noOfLateSlides=0;
                foreach (TaskItemGroup lateTaskgroup in taskData.LateTaskGroups)
                {
                    lateSlidePart = oPPart.GetSlidePartsInOrder().ToList()[gridSlideIndex + noOfGridSlides + noOfCompletedSlides + lateSlidesCreated];
                    var table = lateSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Table>().FirstOrDefault();

                    if (table != null && lateTaskgroup.TaskItems.Count > 0)
                    {
                        TableUtilities.PopulateLateTasksTable(table, lateTaskgroup.TaskItems);
                    }
                    lateSlidesCreated++;
                    noOfLateSlides++;
                }
                                
                 #endregion
                if (taskData.ChartsData != null)
                {
                    int count = 0;
                    foreach (string key in taskData.ChartsData.Keys)
                    {
                       if(key.StartsWith("Show On"))
                       {
                        //Get all Tasks related to  Driving path
                            TaskItemGroup group = new TaskItemGroup() { ChartTaskItems = taskData.ChartsData[key] };
                            var chartDataTable = group.GetChartDataTable(key);
                       
                        #region Charts
                            if (chartDataTable != null)
                            {

                                chartSlidePart = oPPart.GetSlidePartsInOrder().ToList()[gridSlideIndex + noOfGridSlides + noOfCompletedSlides + noOfLateSlides + count];

                                if (chartSlidePart.ChartParts.ToList().Count > 0)
                                {
                                    count++;
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
                                                              new DocumentFormat.OpenXml.Drawing.Text { Text = key })));
                                    }




                                }
                        #endregion
                            }
                        }



                    }

                }
            }

            return ms;
        }

        public SlidePart lateSlidePart { get; set; }
    }
}
