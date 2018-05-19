using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTaxReport.Models
{
    public class PaymentInstallment
    {
        public decimal base_amount { get; set; }
        public byte is_delinquent { get; set; }
        public byte is_estimate { get; set; }
        public decimal delinquent_amount { get; set; }
        public DateTime date_paid { get; set; }
        public decimal amount_due { get; set; }
        public DateTime date_first_due { get; set; }
        public DateTime date_last_due { get; set; }
        public DateTime date_final_due { get; set; }
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
