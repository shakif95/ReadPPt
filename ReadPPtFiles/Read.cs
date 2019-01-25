using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice;
using NetOffice.OfficeApi.Enums;
using PowerPoint = NetOffice.PowerPointApi;


namespace ReadPPtFiles
{
    public static class Read
    {
        public static StringBuilder Start(string filePath)
        {
            StringBuilder stringBuilder = new StringBuilder();

            PowerPoint.Application powerApplication = new PowerPoint.Application();
            PowerPoint.Presentations pptPresentations = powerApplication.Presentations;

            PowerPoint.Presentation pptPresentation = pptPresentations.Open(filePath,
                                                          MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

            PowerPoint.Slides pptSlides = pptPresentation.Slides;

            if(pptSlides != null)
            {
                //var slidesCount = pptSlides.Count;

                foreach(PowerPoint.Slide slide in pptSlides)
                {
                    foreach(PowerPoint.Shape shape in slide.Shapes)
                    {
                        if(shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            PowerPoint.TextRange range = shape.TextFrame.TextRange;
                            if (range != null && range.Length > 0)
                            {
                                stringBuilder.Append(" " + range.Text);
                                
                            }
                        }
                    }
                }
            }
            return stringBuilder;
        }
    }
}
