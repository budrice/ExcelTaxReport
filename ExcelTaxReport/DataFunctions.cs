using ExcelTaxReport.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace ExcelTaxReport
{
    /// <summary>
    /// DataFunctions project shared functions
    /// </summary>
    public class DataFunctions
    {
        /// <summary>
        /// DateToString
        /// </summary>
        /// <param name="dt"></param>
        /// <returns>Returns date string value or empty string.</returns>
        public static string DateToString(DateTime dt)
        {
            string date_val = dt.ToString("MM/dd/yyyy");
            if (dt == DateTime.MinValue)
            {
                date_val = string.Empty;
            }
            return date_val;
        }

        /// <summary>
        /// TownCity
        /// </summary>
        /// <param name="tax_records"></param>
        /// <returns>Returns string tax type.</returns>
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

        /// <summary>
        /// SchoolDistrict
        /// </summary>
        /// <param name="records"></param>
        /// <returns>Returns string school district or empty string.</returns>
        public static string SchoolDistrict(List<TaxAuthorityPaymentRecord> records)
        {
            string school_district = string.Empty;
            foreach (TaxAuthorityPaymentRecord record in records)
            {
                if (record.tax_type == "School")
                {
                    school_district = record.tax_authority.name;
                }
            }
            return school_district;
        }

        /// <summary>
        /// HasExemptions
        /// </summary>
        /// <param name="records"></param>
        /// <param name="checkboxes"></param>
        /// <returns>Returns string for wingdings checkboxes representing Yes/No or a string for Yes or No.</returns>
        public static string HasExemptions(List<TaxAuthorityPaymentRecord> records, bool checkboxes = true)
        {
            string has_exemptions = (checkboxes) ? "x No p Yes" : "No";
            foreach (TaxAuthorityPaymentRecord record in records)
            {
                if (record.ex_disabled == 1
                 || record.ex_elderly == 1
                 || record.ex_homestead == 1
                 || record.ex_mortgage == 1
                 || record.ex_star == 1
                 || record.ex_veteran == 1
                 || record.ex_other.Length > 0)
                {
                    has_exemptions = (checkboxes) ? "p No x Yes" : "Yes";
                }
            }
            return has_exemptions;
        }

        /// <summary>
        /// ExemptionString
        /// </summary>
        /// <param name="records"></param>
        /// <param name="state"></param>
        /// <returns>Returns string of exemptions separated by semi-colons.</returns>
        public static string ExemptionString(List<TaxAuthorityPaymentRecord> records, string state)
        {
            int ex_disabled = 0;
            int ex_elderly = 0;
            int ex_homestead = 0;
            int ex_mortgage = 0;
            int ex_star = 0;
            int ex_veteran = 0;
            string other = string.Empty;
            string exemptions = string.Empty;

            foreach (TaxAuthorityPaymentRecord record in records)
            {
                if(record.ex_disabled == 1)
                {
                    ex_disabled = 1;
                }
                if (record.ex_elderly == 1)
                {
                    ex_elderly = 1;
                }
                if (record.ex_homestead == 1)
                {
                    ex_homestead = 1;
                }
                if (record.ex_mortgage == 1)
                {
                    ex_mortgage = 1;
                }
                if (record.ex_star == 1)
                {
                    ex_star = 1;
                }
                if (record.ex_veteran == 1)
                {
                    ex_veteran = 1;
                }
                
            }
            if (ex_homestead == 1)
            {
                if(state == "CA")
                {
                    exemptions = (exemptions.Length > 0) ? "; " + "HomeOwners Exempt" : "HomeOwners Exempt";
                }
                else
                {
                    exemptions = (exemptions.Length > 0) ? "; " + "Homestead Exempt" : "Homestead Exempt";
                }
                ex_homestead = 1;
            }
            if (ex_disabled == 1)
            {
                exemptions = (exemptions.Length > 0) ? "; " + "Disabled Exempt" : "Disabled Exempt";
            }
            if (ex_veteran == 1)
            {
                exemptions = (exemptions.Length > 0) ? "; " + "Veteran Exempt" : "Veteran Exempt";
            }
            if (ex_mortgage == 1)
            {
                exemptions = (exemptions.Length > 0) ? "; " + "Mortgage Exempt" : "Mortgage Exempt";
            }
            if (ex_star == 1)
            {
                exemptions = (exemptions.Length > 0) ? "; " + "Star Exempt" : "Star Exempt";
            }
            if (ex_elderly == 1)
            {
                exemptions = (exemptions.Length > 0) ? "; " + "Elderly Exempt" : "Elderly Exempt";
            }
            return exemptions;
        }

        /// <summary>
        /// StrInString
        /// </summary>
        /// <param name="stringtocheck"></param>
        /// <param name="value"></param>
        /// <returns>Boolean representing found string value in string</returns>
        public static bool StrInString(string stringtocheck, string value)
        {
            bool found = false;
            int j = stringtocheck.IndexOf(value);
            if (j >= 0)
            {
                found = true;
            }
            return found;
        }

        /// <summary>
        /// StrInString
        /// </summary>
        /// <param name="stringtocheck"></param>
        /// <param name="value"></param>
        /// <returns>Boolean representing found value in array of strings</returns>
        public static bool StrInString(string stringtocheck, string[] value)
        {
            bool found = false;
            int len = value.Length;
            for (int i = 0; i < len; i++)
            {
                string val = value[i];
                int j = stringtocheck.IndexOf(val);
                if (j >= 0)
                {
                    found = true;
                }
            }
            return found;
        }

        /// <summary>
        /// IsDelinquent
        /// </summary>
        /// <param name="installments"></param>
        /// <param name="checkboxes"></param>
        /// <returns>Returns string for wingdings checkboxes representing Yes/No or a string for Yes or No.</returns>
        public static string IsDelinquent(List<PaymentInstallment> installments, bool checkboxes = true)
        {
            string is_delq = (checkboxes) ? "x No p Yes" : "No";
            foreach (PaymentInstallment install in installments)
            {
                if (install.is_delinquent == 1)
                {
                    is_delq = (checkboxes) ? "p No x Yes" : "Yes";
                }
            }
            return is_delq;
        }

        /// <summary>
        /// DelinquencyDescription
        /// </summary>
        /// <param name="installments"></param>
        /// <returns>Returns note.</returns>
        public static string DelinquencyDescription(List<PaymentInstallment> installments)
        {
            string note = string.Empty;
            foreach (PaymentInstallment install in installments)
            {
                string year = install.year;
                string num = install.install.ToString();
                string bamt = String.Format("{0:C}", install.base_amount);
                string dtdue = install.date_due.ToString("MM/dd/yyyy");
                string delamt = String.Format("{0:C}", install.delinquent_amount);
                string gthru = install.date_good_thru.ToString("MM/dd/yyyy");

                if (note.Length > 0)
                {
                    note = note + "; ";
                }
                if (install.is_delinquent == 1)
                {
                    note = note + year + " Installment " + num + ", Base Amount: " + bamt + " originally due " + dtdue + " delinquent in the amount of " + delamt + ". Good through " + gthru + ".\n";
                }
            }
            return note;
        }

        /// <summary>
        /// StringHeight
        /// </summary>
        /// <param name="text"></param>
        /// <param name="font_size"></param>
        /// <param name="col_width_px"></param>
        /// <returns>Returns string height float value.</returns>
        public static float StringHeight(string text, int font_size, int col_width_px)
        {
            Font font = new Font("Arial", font_size, FontStyle.Regular);
            Bitmap bit = new Bitmap(2000, 2000);
            Graphics graphic = Graphics.FromImage(bit);
            SizeF str_size = new SizeF();
            str_size = graphic.MeasureString(text, font, col_width_px);
            return str_size.Height;
        }

        /// <summary>
        /// TotalBilled
        /// </summary>
        /// <param name="installments"></param>
        /// <returns>Returns the decimal value for total base_amount.</returns>
        public static decimal TotalBilled(List<PaymentInstallment> installments)
        {
            decimal total = 0;
            int i = 0;
            int count = installments.Count();
            for (i = 0; i < count; i++)
            {
                string a = installments.ElementAt(0).year;
                string b = installments.ElementAt(i).year;
                if (String.Compare(a, b) == 0)
                {
                    total = total + installments.ElementAt(i).base_amount;

                }
            }
            return total;
        }

        /// <summary>
        /// IsPaid
        /// </summary>
        /// <param name="date_paid"></param>
        /// <param name="checkboxes"></param>
        /// <returns>Returns string for wingdings checkboxes representing Paid/Unpaid or a string for Paid or Owing.</returns>
        public static string IsPaid(DateTime date_paid, bool checkboxes = true)
        {
            string is_paid = (checkboxes) ? "x Paid p Unpaid" : "Paid";
            if (date_paid == DateTime.MinValue)
            {
                is_paid = (checkboxes) ? "p Paid x Unpaid" : "Owing";
            }
            return is_paid;
        }
    }
}
