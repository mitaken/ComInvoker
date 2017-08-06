using Microsoft.Office.Interop.Excel;

namespace ComInvoker.Sample
{
    internal class SampleExcel : Invoker
    {
        /// <summary>
        /// Excel instance
        /// </summary>
        private Application excel;

        /// <summary>
        /// Create excel instance
        /// </summary>
        internal SampleExcel()
        {
            excel = Invoke<Application>(new Application());
            excel.Visible = true;
        }


        internal void Write1To100()
        {
            //Get Workbooks
            var workbooks = Invoke<Workbooks>(excel.Workbooks);
            //Add Workbook
            var workbook = Invoke<Workbook>(workbooks.Add());
            //Get Worksheets
            var worksheets = InvokeEnumurator<Worksheet>(workbook.Sheets);
            foreach (var worksheet in worksheets)
            {
                var cells = Invoke<Range>(worksheet.Cells);
                for (var i = 1; i < 1000; i++)
                {
                    //Use real type for intellisense when return dynamic
                    Range cell = Invoke<Range>(cells[i, 1]);
                    cell.Value = i;
                }
            }

            workbook.Close(SaveChanges: false);
        }

        /// <summary>
        /// Override dispose (for Quit)
        /// </summary>
        /// <param name="disposing"></param>
        protected override void Dispose(bool disposing)
        {
            //Exit excel
            excel?.Quit();

            //Call parent dispose
            base.Dispose(disposing);
        }
    }
}
