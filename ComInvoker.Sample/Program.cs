using System;

namespace ComInvoker.Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var excel = new SampleExcel())
            {
                excel.Write1To100();
                Console.WriteLine("Press any key to quit excel");
                Console.ReadKey();
            }
        }
    }
}
