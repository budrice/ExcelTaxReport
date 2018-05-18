using System;
using System.Collections.Generic;

namespace ReportBuilder.Models
{
    public class TaxAuthorityPaymentRecord
    {
        public string tax_type { get; set; }
        public string additional_data { get; set; }
        public byte other_auth { get; set; }
        public string research_notes { get; set; }
        public byte lump_sum { get; set; }
        public byte prior_year_del { get; set; }
        public byte any_exemptions { get; set; }
        public byte ex_homestead { get; set; }
        public byte ex_disabled { get; set; }
        public byte ex_veteran { get; set; }
        public byte ex_mortgage { get; set; }
        public byte ex_star { get; set; }
        public byte ex_elderly { get; set; }
        public string ex_other { get; set; }
        public decimal milage_rate { get; set; }
        public decimal assessed_value { get; set; }
        public DateTime milage_next_due { get; set; }
        public decimal land_value { get; set; }
        public decimal improved_value { get; set; }
        public decimal total_value { get; set; }
        public byte unincorporated { get; set; }
        public string lawsuit { get; set; }
        public string lawsuit_case { get; set; }
    }
}
