using System;

namespace ExcelTaxReport.Models
{
    /// <summary>
    /// payment Installment
    /// </summary>
    /// <remarks>
    /// A payment installment to tax authority payment.
    /// Checkboxes are byte values: 0 equals false; 1 equals true
    /// </remarks>
    public class PaymentInstallment
    {
        public decimal base_amount { get; set; }
        public byte is_delinquent { get; set; }
        public byte is_estimate { get; set; }
        public decimal delinquent_amount { get; set; }
        public DateTime date_paid { get; set; }
        public decimal amount_due { get; set; }
        //public DateTime date_first_due { get; set; }
        //public DateTime date_last_due { get; set; }
        public DateTime date_due { get; set; }
        public int install { get; set; }
        public decimal paid { get; set; }
        public byte is_partial { get; set; }
        public DateTime date_good_thru { get; set; }
        public string year { get; set; }
        public byte is_exempt { get; set; }
        public decimal one_month { get; set; }
        public decimal two_month { get; set; }
        public string status
        {
            get
            {
                String status = String.Empty;

                if ((this.base_amount - this.paid) == 0)
                {
                    status = "Paid";
                }
                else
                {
                    status = "Owing";
                }
                return status;
            }
        }
    }
}
