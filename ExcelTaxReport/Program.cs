using ExcelTaxReport.Reports;

namespace ExcelTaxReport
{
    class Program
    {
        static void Main(string[] args)
        {
            Report1 report = new Report1(TestData.GetSampleOrder());
            report.CreateReport();
        }
    }
}
