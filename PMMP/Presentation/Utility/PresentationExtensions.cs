using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.IO;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Reflection;


namespace PMMP
{
    public static class PresentationExtensions
    {
        static char[] hexDigits = {

'0', '1', '2', '3', '4', '5', '6', '7',

'8', '9', 'A', 'B', 'C', 'D', 'E', 'F'};
        public static bool InCurrentFiscalMonth(this DateTime date,Repository.FiscalMonth months)
        {
            if (date >= months.From && date <= months.To)
                return true;
            else
                return false;
        }

        public static string ToVeryShortDateString(this DateTime date)
        {
            string format = "M/d";
            return date.ToString(format);
        }

        public static string ToHexString(this System.Drawing.Color color)
        {
            byte[] bytes = new byte[3];
            bytes[0] = color.R;
            bytes[1] = color.G;
            bytes[2] = color.B;
            char[] chars = new char[bytes.Length * 2];
            for (int i = 0; i < bytes.Length; i++)
            {
                int b = bytes[i];
                chars[i * 2] = hexDigits[b >> 4];
                chars[i * 2 + 1] = hexDigits[b & 0xF];
            }
            return new string(chars);
        }


        

        public static IEnumerable<SlidePart> GetSlidePartsInOrder(this PresentationPart presentationPart)
        {
            Repository.Utility.WriteLog("GetSlidePartsInOrder started", System.Diagnostics.EventLogEntryType.Information);
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            var objects =  slideIdList.ChildElements
                .Cast<SlideId>()
                .Select(x => presentationPart.GetPartById(x.RelationshipId))
                .Cast<SlidePart>();
            Repository.Utility.WriteLog("GetSlidePartsInOrder completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return objects;
        }

        public static SlidePart CloneSlide(this SlidePart templatePart, SlideType type)
        {
            Repository.Utility.WriteLog("CloneSlide started", System.Diagnostics.EventLogEntryType.Information);
            // find the presentationPart: makes the API more fluent
            var presentationPart = templatePart.GetParentParts().OfType<PresentationPart>().Single();

            int i = presentationPart.SlideParts.Count();

            // clone slide contents
            var slidePartClone = presentationPart.AddNewPart<SlidePart>("newSlide" + i);
            slidePartClone.FeedData(templatePart.GetStream(FileMode.Open));

            // copy layout part
            slidePartClone.AddPart(templatePart.SlideLayoutPart, templatePart.GetIdOfPart(templatePart.SlideLayoutPart));

            if (type == SlideType.Grid)
            {
               
                foreach (IdPartPair part in templatePart.Parts)
                {
                    var tPart = templatePart.GetPartById(part.RelationshipId);

                    var embeddedPackagePart = tPart as EmbeddedPackagePart;
                    if (embeddedPackagePart != null)
                    {
                        var newPart = slidePartClone.AddEmbeddedPackagePart(embeddedPackagePart.ContentType);
                        newPart.FeedData(embeddedPackagePart.GetStream());
                        slidePartClone.ChangeIdOfPart(newPart, templatePart.GetIdOfPart(embeddedPackagePart));
                    }

                    var vmlDrawingPart = tPart as VmlDrawingPart;
                    if (vmlDrawingPart != null)
                    {
                        var newPart = slidePartClone.AddNewPart<VmlDrawingPart>();
                        newPart.FeedData(vmlDrawingPart.GetStream());

                        var drawingImg = vmlDrawingPart.ImageParts.ToList()[0];
                        var newImgPart = newPart.AddImagePart(drawingImg.ContentType, templatePart.GetIdOfPart(drawingImg));
                        newImgPart.FeedData(drawingImg.GetStream());
                    }

                    var imagePart = tPart as ImagePart;
                    if (imagePart != null)
                    {
                        var newPart = slidePartClone.AddImagePart(imagePart.ContentType, templatePart.GetIdOfPart(imagePart));
                        newPart.FeedData(imagePart.GetStream());
                    }
                   
                }
               
            }
            else
            {
                foreach (ChartPart cpart in templatePart.ChartParts)
                {
                    ChartPart newcpart = slidePartClone.AddNewPart<ChartPart>(templatePart.GetIdOfPart(cpart));
                    newcpart.FeedData(cpart.GetStream());
                    // copy the embedded excel file
                    EmbeddedPackagePart epart = newcpart.AddEmbeddedPackagePart(cpart.EmbeddedPackagePart.ContentType);
                    epart.FeedData(cpart.EmbeddedPackagePart.GetStream());
                    // link the excel to the chart
                    newcpart.ChartSpace.GetFirstChild<ExternalData>().Id = newcpart.GetIdOfPart(epart);
                    newcpart.ChartSpace.Save();
                }
                
            }
            Repository.Utility.WriteLog("CloneSlide completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return slidePartClone;
        }



        public static void AppendSlide(this PresentationPart presentationPart, SlidePart newSlidePart)
        {
            Repository.Utility.WriteLog("AppendSlide started", System.Diagnostics.EventLogEntryType.Information);
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // find the highest id
            uint maxSlideId = slideIdList.ChildElements.Cast<SlideId>().Max(x => x.Id.Value);
            // Insert the new slide into the slide list after the previous slide.
            var id = maxSlideId + 1;

            SlideId newSlideId = new SlideId();
            slideIdList.Append(newSlideId);
            newSlideId.Id = id;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(newSlidePart);
            Repository.Utility.WriteLog("AppendSlide completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        public static string GetString(this CustomFieldType e)
        {
            switch (e)
            {
                case CustomFieldType.CA: return "Cost Account";
                case CustomFieldType.EstFinish: return "CAM Finish";
                case CustomFieldType.EstStart: return "CAM Start";
                case CustomFieldType.PMT: return "PMT";
                case CustomFieldType.ReasonRecovery: return "Reason _Recovery";
                case CustomFieldType.ShowOn: return "Show On";
            }
            return null;
        }
    }
}
