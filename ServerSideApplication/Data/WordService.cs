using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace ServerSideApplication.Data
{
    public class WordService
    {

        /// <summary>
        /// Create a 'Hello World' word document
        /// </summary>
        /// <returns>Return the created word document as stream</returns>
        public MemoryStream CreateWord()
        {
            //Create an instance of WordDocument (Empty Word Document)
            using (WordDocument document = new WordDocument())
            {
                //Add a section & a paragraph in the empty document
                document.EnsureMinimal();

                //Append text to the last paragraph of the document
                document.LastParagraph.AppendText("Hello World");

                //Save the document as a stream and retrun the stream.
                using (MemoryStream stream = new MemoryStream())
                {
                    //Save the created Word document to MemoryStream.
                    document.Save(stream, FormatType.Docx);

                    return stream;
                }
            }
        }
    }
}
