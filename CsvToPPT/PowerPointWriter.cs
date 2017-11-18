using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace CsvToPPT
{
    class PowerPointWriter
    {
        public PowerPointWriter() { }

        public void makePresentation(List<WorkItemInfo> workItems)
        {
            PowerPoint.Application objApp;
            PowerPoint._Presentation objPres;
            PowerPoint.Slides objSlides;
            PowerPoint._Slide objSlide;
            PowerPoint.TextRange objTextRng;

            //Create a new presentation based on a template.
            objApp = new PowerPoint.Application();
            objPres = objApp.Presentations.Add(MsoTriState.msoTrue);
            objApp.Visible = MsoTriState.msoTrue;
            objSlides = objPres.Slides;

            //Build Slide #1:
            objSlide = objSlides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);

            foreach (var info in workItems)
            {
                addObject(objSlide, info);
            }

            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Font.Size = 8;
        }

        private void addObject(PowerPoint._Slide slide, WorkItemInfo info)
        {
            PowerPoint.Shape shp = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 50, 50, Constants.WIDTH, info.Cost * Constants.HEIGHT_PER_COST);
            shp.TextFrame.TextRange.Text = info.Id + "\n" + info.Summary;
            shp.TextFrame.TextRange.Font.Size = 6;
        }
    }
}
