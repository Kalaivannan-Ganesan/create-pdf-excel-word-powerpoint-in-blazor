using System.IO;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;

namespace ServerSideApplication.Data
{
    public class PDFService
    {
        /// <summary>
        /// Create the 'Hello World' PDF document
        /// </summary>
        /// <returns>Return the created PDF document as stream</returns>
        public MemoryStream CreatePDF()
        {
            //Create a new PDF document
            using (PdfDocument document = new PdfDocument())
            {
                //Add a page to the document
                PdfPage page = document.Pages.Add();

                //Create PDF graphics for the page
                PdfGraphics graphics = page.Graphics;

                //Set the standard font
                PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 20);

                //Draw the text
                graphics.DrawString("Hello World!", font, PdfBrushes.Black, new Syncfusion.Drawing.PointF(0, 0));

                //Save the document as a stream and retrun the stream.
                using (MemoryStream stream = new MemoryStream())
                {
                    //Save the created PDF document to MemoryStream.
                    document.Save(stream);

                    return stream;
                }
            }
        }
    }
}
