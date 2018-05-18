using System;
using System.Collections.Generic;

namespace ReportBuilder.Models
{
    public class ParcelInformation
    {
        public int id { get; set; }
        public string researcher { get; set; }
        public string client_po_number { get; set; }
        public string address { get; set; }
        public string city { get; set; }
        public string state { get; set; }
        public string zip_code { get; set; }
        public string address_string
        {
            get
            {
                return this.address + ", " + this.city + ", " + this.state + " " + this.zip_code;
            }
        }
        public string searched_address { get; set; }
        public string searched_city { get; set; }
        public string searched_state { get; set; }
        public string searched_zip_code { get; set; }
        public string searched_address_string
        {
            get
            {
                return this.searched_address + ", " + this.searched_city + ", " + this.searched_state + " " + searched_zip_code;
            }
        }
        public string county { get; set; }
        public string owner_1 { get; set; }
        public string owner_2 { get; set; }
        public string owners
        {
            get
            {
                string owners = this.owner_1;
                if (this.owner_2 != null && string.Compare(this.owner_2, string.Empty) != 0)
                {
                    owners = owners + " and " + this.owner_2;
                }
                return owners;
            }
        }
        public string legal_desc { get; set; }
        public string class_code { get; set; }
        public DateTime effective_date { get; set; }
        public string parcel_number { get; set; }
        public string cad_number { get; set; }
        public string acreage { get; set; }
        public string hoa { get; set; }
        public string hoa_note { get; set; }
        public DateTime date_received { get; set; }
        public decimal land_value { get; set; }
        public decimal improvement_value { get; set; }
        public string assessed_owner_1 { get; set; }
        public string assessed_owner_2 { get; set; }
        public string assessed_owners
        {
            get
            {
                string assessed_owners = this.assessed_owner_1;
                if (this.assessed_owner_2 != null && string.Compare(this.assessed_owner_2, string.Empty) != 0)
                {
                    assessed_owners = assessed_owners + " and " + this.assessed_owner_2;
                }
                return assessed_owners;
            }
        }
        public string assessed_address
        {
            get
            {
                string assessed_address = string.Empty;
                assessed_address = this.address + ", " + this.city + ", " + this.state + " " + this.zip_code;
                return assessed_address;
            }
        }
        public decimal assessed_valuation { get; set; }
        
    }

}