using ExcelTaxReport.Reports;

namespace ExcelTaxReport
{
    /// <summary>
    /// Author: Eldis Rice
    /// Date: 05/18/2018
    /// Description: A tax research Excel worksheet generator.
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Report1 report = new Report1(TestData.GetSampleOrder());
            report.CreateReport();
        }
    }
}
