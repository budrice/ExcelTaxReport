using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelTaxReport.Models;

namespace ExcelTaxReport
{
    class TestData
    {
        public static ClientOrder GetSampleOrder()
        {
            ClientOrder client_order = new ClientOrder();

            ReportConfig report_config = new ReportConfig();
            report_config.disclaimer = "We have made every effort to ensure the accuracy of this tax information.  However, due to the frequency with which municipalities revise their fees and other specifications, we cannot assume liability for any discrepancy in the taxes.  In the event that tax amounts have changed, please notify us so we can update our records.  Possible revenue bond charges for sewer and water pursuant to state statutes and local ordinances when connection to the system is made by the owner.  The exact current and continuing charges depend on all the facts.  Contact local officials for details.  This report is based on best available information at the time.  This is for informational purposes only and will not appear on title policy.\r\nPatent Pending ";
            client_order.report_config = report_config;

            ClientConfig client_config = new ClientConfig();
            client_config.base_path = @"D:\Programming\0_crap";
            client_config.report_name = "Tax_Research.xlsx";
            client_config.logo = @"D:\Programming\0_crap\logo.jpg";
            client_order.client_config = client_config;

            ParcelInformation parcel = new ParcelInformation();
            parcel.client_po_number = "U NTN-ARS-12345";
            parcel.researcher = "erice";
            parcel.owner_1 = "Buddy Rice";
            parcel.owner_2 = "Ashley Gonzalez";
            parcel.address = "427 8th Ave N";
            parcel.city = "Saint Petersburg";
            parcel.state = "FL";
            parcel.zip_code = "33701";
            parcel.county = "Pinellas";

            parcel.searched_address = "427 8th Ave N";
            parcel.searched_city = "Saint Petersburg";
            parcel.searched_state = "FL";
            parcel.searched_zip_code = "33701";

            parcel.assessed_owner_1 = "Buddy Rice";
            parcel.assessed_owner_2 = "Ashley Gonzalez";
            parcel.assessed_valuation = 150000.00M;

            parcel.parcel_number = "18-31-17-77814-001-0060";
            parcel.effective_date = DateTime.Parse("2017-06-01");
            parcel.legal_desc = "SAFFORD'S ADD REVISED BLK 1, E 75FT OF LOTS 6 AND 7"




            return client_order;
        }
    }
}
