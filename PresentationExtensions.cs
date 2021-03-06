﻿using System;
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

        public static string StringValueOf(this Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());
            Microsoft.SharePoint.Linq.ChoiceAttribute[] attributes =
                (Microsoft.SharePoint.Linq.ChoiceAttribute[])fi.GetCustomAttributes(
                typeof(Microsoft.SharePoint.Linq.ChoiceAttribute), false);
            if (attributes.Length > 0)
            {
                return attributes[0].Value;
            }
            else
            {
                return value.ToString();
            }
        }

        public static string ExceptChars(this string str, IEnumerable<char> toExclude)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < str.Length; i++)
            {
                char c = str[i];
                if (!toExclude.Contains(c))
                    sb.Append(c);
            }
            return sb.ToString();
        }

        public static IEnumerable<SlidePart> GetSlidePartsInOrder(this PresentationPart presentationPart)
        {
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            return slideIdList.ChildElements
                .Cast<SlideId>()
                .Select(x => presentationPart.GetPartById(x.RelationshipId))
                .Cast<SlidePart>();
        }

        public static SlidePart CloneSlide(this SlidePart templatePart, SlideType type)
        {
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

            return slidePartClone;
        }



        public static void AppendSlide(this PresentationPart presentationPart, SlidePart newSlidePart)
        {
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // find the highest id
            uint maxSlideId = slideIdList.ChildElements.Cast<SlideId>().Max(x => x.Id.Value);
            // Insert the new slide into the slide list after the previous slide.
            var id = maxSlideId + 1;

            SlideId newSlideId = new SlideId();
            slideIdList.Append(newSlideId);
            newSlideId.Id = id;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(newSlidePart);
        }
    }
}
