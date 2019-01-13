namespace ComInvoker.Sample
{
    internal class SampleExcel : Invoker
    {
        /// <summary>
        /// Excel instance
        /// </summary>
        private dynamic excel;

        /// <summary>
        /// Create excel instance
        /// </summary>
        internal SampleExcel()
        {
            excel = InvokeFromProgID("Excel.Application");//Application
            excel.Visible = true;
        }


        internal void Write1To100()
        {
            //Get Workbooks
            var workbooks = Invoke(excel.Workbooks);//Workbooks
            //Add Workbook
            var workbook = Invoke(workbooks.Add());//Workbook
            //Get Worksheets
            var worksheets = InvokeEnumurator(workbook.Worksheets);//IEnumerable<Worksheet>
            foreach (var worksheet in worksheets)
            {
                var cells = Invoke(worksheet.Cells);//Range
                for (var i = 1; i < 1000; i++)
                {
                    dynamic cell = Invoke(cells[i, 1]);//Range
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
