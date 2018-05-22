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
    }
}
