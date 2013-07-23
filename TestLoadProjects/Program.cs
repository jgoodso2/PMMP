using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using PMMP;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml.Serialization;
using System.Xml;
using System.Data;
using System.ComponentModel;
using DocumentFormat.OpenXml;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using DocumentFormat.OpenXml.Drawing.Charts;


namespace TestLoadProjects
{
    class Program
    {
    //    public CopyChartFromXlsx2Pptx(string SourceFile, string TargetFile, string targetppt)
    //{

    //    ChartPart chartPart;

    //    ChartPart newChartPart;

    //    //SlidePart slidepartbkMark = null;

    //    string chartPartIdBookMark = null;

    //    File.Copy(TargetFile, targetppt, true);

    //    //Powerpoint document

    //    using (OpenXmlPkg.PresentationDocument pptPackage = OpenXmlPkg.PresentationDocument.Open(targetppt, true))
    //    {

    //        OpenXmlPkg.PresentationPart presentationPart = pptPackage.PresentationPart;

    //        var secondSlidePart = pptPackage.PresentationPart.SlideParts.Skip(0).First();  // this will retrieve your second slide

    //        chartPart = secondSlidePart.ChartParts.First();

    //        chartPartIdBookMark = secondSlidePart.GetIdOfPart(chartPart);

    //        //Console.WriteLine("ID:"+chartPartIdBookMark.ToString());

    //        secondSlidePart.DeletePart(chartPart);

    //        secondSlidePart.Slide.Save();

    //        newChartPart = secondSlidePart.AddNewPart<ChartPart>(chartPartIdBookMark);

    //        ChartPart saveXlsChart = null;

    //        using (SpreadsheetDocument xlsDocument = SpreadsheetDocument.Open(SourceFile.ToString(), true))
    //        {

    //            WorkbookPart xlsbookpart = xlsDocument.WorkbookPart;

    //            foreach (var worksheetPart in xlsDocument.WorkbookPart.WorksheetParts)
    //            {

    //                if (worksheetPart.DrawingsPart != null)

    //                    if (worksheetPart.DrawingsPart.ChartParts.Any())
    //                    {
    //                        saveXlsChart = worksheetPart.DrawingsPart.ChartParts.First();
    //                    }

    //            }

    //            newChartPart.FeedData(saveXlsChart.GetStream());
    //            //newChartPart.FeedData(
    //            secondSlidePart.Slide.Save();

    //            xlsDocument.Close();

    //            pptPackage.Close();

    //        }

    //    }
    //}
        static void Main(string[] args)
        {
            try
            {
                 ChartPart saveXlsChart = null;
                 
                using (Stream stream = new FileStream("POC1.pptx", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                 {
                     using (var oPDoc = PresentationDocument.Open(stream,true,new OpenSettings()))
                     {
                         var oPPart = oPDoc.PresentationPart;
                         SlidePart chartSlidePart = oPPart.GetSlidePartsInOrder().ToList()[0];

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
                                         DocumentFormat.OpenXml.Spreadsheet.Row row = WorkbookUtilities.GetRow(sheetData, 3);
                                         Cell dataCell1 = WorkbookUtilities.GetCell(row, 3);
                                         if (dataCell1.DataType != null && dataCell1.DataType == CellValues.SharedString)
                                             dataCell1.DataType = CellValues.String;
                                         dataCell1.CellValue.Text = "10";
                                         if (workSheetPart.DrawingsPart != null)

                                             if (workSheetPart.DrawingsPart.ChartParts.Any())
                                             {
                                                 saveXlsChart = workSheetPart.DrawingsPart.ChartParts.First();
                                             }
                                         //chartPart.ChartSpace.ChildElements[2].ChildElements[1].InnerXml = saveXlsChart.ChartSpace.ChildElements[2].ChildElements[1].InnerXml;
                                        // chartPart.FeedData(saveXlsChart.GetStream());
                                         //workSheetPart.Worksheet.Save();
                                         
                                        
                                         //chartSlidePart.DeletePart(chartPart);
                                         
                                        
                                         //chartSlidePart.FeedData(oSDoc.WorkbookPart.GetStream());
                                         
                                     }
                                     using (var oSDoc = SpreadsheetDocument.Open(spreadsheet.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite), true))
                                     {
                                         var workSheetPart = oSDoc.WorkbookPart.GetPartsOfType<WorksheetPart>().FirstOrDefault();
                                         var sheetData = workSheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                                         //var newChartPart = chartSlidePart.AddNewPart<ChartPart>(chartPartIdBookMark);
                                         //newChartPart.FeedData(saveXlsChart.GetStream());
                                         //chartSlidePart.Slide.Save();
                                         //chartSlidePart.FeedData(oSDoc.WorkbookPart.GetStream());

                                     }
                                 }
                                 foreach (SlideMasterPart master in oPPart.SlideMasterParts)
                                 {
                                     master.SlideMaster.Save();

                                 }

                                 //chartPart.EmbeddedPackagePart.FeedData(spreadsheet.GetStream());
                             }
                             chartSlidePart.Slide.Save();
                             
                             // SaveStreamToFile
                             //         ("POC1.pptx",ms);

                         }
                     }
                 }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occured and the exception message ={0}", ex.Message);
            }
            Console.ReadKey();
        }

        //public static DataTable ToDataTable<T>(IList<T> data)
        //{
        //    PropertyDescriptorCollection props =
        //        TypeDescriptor.GetProperties(typeof(T));
        //    DataTable table = new DataTable();
        //    for (int i = 0; i < props.Count; i++)
        //    {
        //        PropertyDescriptor prop = props[i];
        //        table.Columns.Add(prop.Name, prop.PropertyType);
        //    }
        //    object[] values = new object[props.Count];
        //    foreach (T item in data)
        //    {
        //        for (int i = 0; i < values.Length; i++)
        //        {
        //            values[i] = props[i].GetValue(item);
        //        }
        //        table.Rows.Add(values);
        //    }
        //    return table;
        //}

        //public static MemoryStream SerializeObject(DataTable ds)// for given List<object>
        //{
        //    try
        //    {
        //        byte[] data = null;
        //        using (MemoryStream stream = new MemoryStream())
        //        {
        //            IFormatter bf = new BinaryFormatter();
        //            ds.RemotingFormat = SerializationFormat.Binary;
        //            bf.Serialize(stream, ds);
        //            return stream;
        //        }
                
        //    }
        //    catch (Exception e) { System.Console.WriteLine(e); return null; }
        //}
    }
}
