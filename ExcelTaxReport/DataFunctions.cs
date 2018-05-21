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

        public static string TownCity(List<TaxAuthorityPaymentRecord> tax_records)
        {
            string taxtype = string.Empty;
            foreach (TaxAuthorityPaymentRecord record in tax_records)
            {
                if (record.tax_type == "City" || record.tax_type == "Town" || record.tax_type == "Township")
                {
                    taxtype = (string.Compare(taxtype, string.Empty) == 0) ? record.tax_authority.name : "; " + record.tax_authority.name;
                }
                else if (record.tax_type == "School")
                {
                    taxtype = record.tax_authority.name;
                }
            }
            return taxtype;




        }
    }
}
