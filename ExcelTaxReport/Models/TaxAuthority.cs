namespace ExcelTaxReport.Models
{
    /// <summary>
    /// tax authority
    /// </summary>
    public class TaxAuthority
    {
        public string name { get; set; }
        public string payment_address_1 { get; set; }
        public string payment_address_2 { get; set; }
        public string payment_city { get; set; }
        public string payment_state { get; set; }
        public string payment_zip { get; set; }
        public string payment_string_address
        {
            get
            {
                string string_address = string.Empty;
                string_address = this.payment_address_1;
                if (this.payment_address_2 != null)
                {
                    string_address = string_address + ", " + this.payment_address_2;
                }
                return string_address;
            }
        }
        public string payment_city_state_zip
        {
            get
            {
                string city_state_zip = string.Empty;
                city_state_zip = this.payment_city + ", " + this.payment_state + " " + this.payment_zip;
                return city_state_zip;
            }
        }
        public string payment_name { get; set; }
        public string payment_phone { get; set; }
        public string payment_ext { get; set; }
        public string payment_phone_string
        {
            get
            {
                string phone_number = string.Empty;
                phone_number = this.payment_phone + " Ext: " + this.payment_ext;
                return phone_number;
            }
        }
        public string current_tax_year { get; set; }
        public string schedule { get; set; }
        public string fiscal_year { get; set; }
        public decimal duplicate_bill_fee { get; set; }
        public string discounts { get; set; }
        public string ta_other_notes { get; set; }
    }
}
