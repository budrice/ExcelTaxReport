using System;

namespace ExcelTaxReport.Models
{
    /// <summary>
    /// tax information
    /// </summary>
    public class TaxInformation
    {
        public string jurisdiction_name { get; set; }
        public string jurisdiction_type { get; set; }
        public string exemptions { get; set; }
        public decimal tax_rate { get; set; }
        public decimal milage_rate { get; set; }
        public DateTime milage_next_due { get; set; }
    }
}
