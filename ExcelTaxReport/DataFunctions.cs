using ExcelTaxReport;
using ExcelTaxReport.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTaxReport
{
    public class DataFunctions
    {
        public static string DateToString(DateTime dt)
        {
            string date_val = dt.ToString("MM/dd/yyyy");
            if (dt == DateTime.MinValue)
            {
                date_val = string.Empty;
            }
            return date_val;
        }
    }
}
