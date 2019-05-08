using Syncfusion.XlsIO;
using System.IO;

namespace ServerSideApplication.Data
{
    public class ExcelService
    {
        /// <summary>
        /// Create a 'Hello World' excel document
        /// </summary>
        /// <returns>Return the created excel document as stream</returns>
        public MemoryStream CreateExcel()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {

                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Assigns default application version
                application.DefaultVersion = ExcelVersion.Excel2013;

                //A new workbook is created equivalent to creating a new workbook in Excel
                //Create a workbook with 1 worksheet
                IWorkbook workbook = application.Workbooks.Create(1);

                //Access first worksheet from the workbook
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding text to a cell
                worksheet.Range["A1"].Text = "Hello World";

                //Save the document as a stream and retrun the stream.
                using (MemoryStream stream = new MemoryStream())
                {
                    //Save the created Excel document to MemoryStream.
                    workbook.SaveAs(stream);

                    return stream;
                }

            }
        }
    }
}
