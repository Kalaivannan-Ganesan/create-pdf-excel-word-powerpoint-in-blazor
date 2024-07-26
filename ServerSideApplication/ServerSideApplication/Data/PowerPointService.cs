using Syncfusion.Presentation;

namespace ServerSideApplication.Data
{
    public class PowerPointService
    {
        /// <summary>
        /// Create a 'Hello World' PowerPoint document
        /// </summary>
        /// <returns>Return the created PowerPoint document as stream</returns>
        public MemoryStream CreatePowerPoint()
        {
            //Create a PowerPoint file
            using (IPresentation pptxFile = Presentation.Create())
            {
                //Add a slide to the PowerPoint Presentation
                ISlide firstSlide = pptxFile.Slides.Add(SlideLayoutType.Blank);

                //Add a textbox in a slide by specifying its position and size
                IShape textShape = firstSlide.AddTextBox(100, 75, 756, 200);

                //Add a paragraph into the textShape
                IParagraph paragraph = textShape.TextBody.AddParagraph();

                //Set the horizontal alignment of paragraph
                paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;

                //Add a textPart in the paragraph
                ITextPart textPart = paragraph.AddTextPart("Hello World");

                //Save the document as a stream and retrun the stream.
                using (MemoryStream pptxStream = new MemoryStream())
                {
                    //Save the created PowerPoint document to MemoryStream.
                    pptxFile.Save(pptxStream);

                    return pptxStream;
                }
            }
        }
    }
}
