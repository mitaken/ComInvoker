using System;

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
            var type = Type.GetTypeFromProgID("Excel.Application");
            if (type == null) throw new TypeLoadException("Excel does not installed");

            excel = Invoke<dynamic>(Activator.CreateInstance(type));//Application
            excel.Visible = true;
        }


        internal void Write1To100()
        {
            //Get Workbooks
            var workbooks = Invoke<dynamic>(excel.Workbooks);//Workbooks
            //Add Workbook
            var workbook = Invoke<dynamic>(workbooks.Add());//Workbook
            //Get Worksheets
            var worksheets = InvokeEnumurator<dynamic>(workbook.Sheets);//Worksheet
            foreach (var worksheet in worksheets)
            {
                var cells = Invoke<dynamic>(worksheet.Cells);//Range
                for (var i = 1; i < 1000; i++)
                {
                    dynamic cell = Invoke<dynamic>(cells[i, 1]);//Range
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
