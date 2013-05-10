using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PMMP
{
    public class PresentationUtilities
    {
        public static int CountSlides(PresentationDocument presentationDocument)
        {
            if (presentationDocument == null)
                throw new ArgumentNullException("presentationDocument");

            int slidesCount = 0;

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart != null)
                slidesCount = presentationPart.SlideParts.Count();

            return slidesCount;
        }

        public static void MoveSlide(PresentationDocument presentationDocument, int from, int to)
        {
            if (presentationDocument == null)
                throw new ArgumentNullException("presentationDocument");

            int slidesCount = CountSlides(presentationDocument);

            if (from < 0 || from >= slidesCount)
                throw new ArgumentOutOfRangeException("from");

            if (to < 0 || from >= slidesCount || to == from)
                throw new ArgumentOutOfRangeException("to");

            PresentationPart presentationPart = presentationDocument.PresentationPart;
            Presentation presentation = presentationPart.Presentation;
            SlideIdList slideIdList = presentation.SlideIdList;
            SlideId sourceSlide = slideIdList.ChildElements[from] as SlideId;
            SlideId targetSlide = null;

            if (to == 0)
                targetSlide = null;

            if (from < to)
                targetSlide = slideIdList.ChildElements[to] as SlideId;
            else
                targetSlide = slideIdList.ChildElements[to - 1] as SlideId;

            sourceSlide.Remove();
            slideIdList.InsertAfter(sourceSlide, targetSlide);

            presentation.Save();
        }

        public static void DeleteSlide(PresentationDocument presentationDocument, int slideIndex)
        {
            if (presentationDocument == null)
                throw new ArgumentNullException("presentationDocument");

            int slidesCount = CountSlides(presentationDocument);

            if (slideIndex < 0 || slideIndex >= slidesCount)
                throw new ArgumentOutOfRangeException("slideIndex");

            PresentationPart presentationPart = presentationDocument.PresentationPart;
            Presentation presentation = presentationPart.Presentation;
            SlideIdList slideIdList = presentation.SlideIdList;
            SlideId slideId = slideIdList.ChildElements[slideIndex] as SlideId;
            string slideRelId = slideId.RelationshipId;

            slideIdList.RemoveChild(slideId);

            if (presentation.CustomShowList != null)
            {
                foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                {
                    if (customShow.SlideList != null)
                    {
                        LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                        foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                            if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                                slideListEntries.AddLast(slideListEntry);

                        foreach (SlideListEntry slideListEntry in slideListEntries)
                            customShow.SlideList.RemoveChild(slideListEntry);
                    }
                }
            }

            presentation.Save();
            SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;
            presentationPart.DeletePart(slidePart);
        }
    }
}
