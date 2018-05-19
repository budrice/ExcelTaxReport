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
            parcel.parcel_number = "18-31-17-77814-001-0060";
            parcel.effective_date = DateTime.Parse("2018-05-01");
            parcel.legal_desc = "SAFFORD'S ADD REVISED BLK 1, E 75FT OF LOTS 6 AND 7";
            parcel.class_code = "Apartments (10 units to 49 units)";

            TaxAuthorityPaymentRecord payment_record = new TaxAuthorityPaymentRecord();
            payment_record.tax_type = "County";
            payment_record.additional_data = "";
            payment_record.research_notes = "";
            payment_record.lump_sum = 1;
            payment_record.ex_homestead = 0;
            payment_record.ex_disabled = 0;
            payment_record.ex_veteran = 1;
            payment_record.ex_mortgage = 0;
            payment_record.ex_star = 0;
            payment_record.ex_elderly = 0;
            payment_record.ex_other = string.Empty;
            payment_record.assessed_value = 408980.00M;
            payment_record.land_value = 225600.00M;
            payment_record.improved_value = 45100.00M;
            payment_record.unincorporated = 0;
            payment_record.lawsuit = string.Empty;
            payment_record.lawsuit_case = string.Empty;

            TaxInformation tax_info = new TaxInformation();
            tax_info.jurisdiction_name = "";
            tax_info.jurisdiction_type = "";
            tax_info.tax_rate = 0.00M;
            tax_info.exemptions = "";
            tax_info.milage_rate = 0.00M;
            tax_info.milage_next_due = DateTime.Parse("2018-12-01");

            PaymentInstallment install_1 = new PaymentInstallment();
            install_1.amount_due = 0.00M;
            install_1.base_amount = 0.00M;
            install_1.date_due = DateTime.Parse("2017-12-31");
            install_1.date_good_thru = DateTime.Parse("2018-02-31");
            install_1.date_paid = DateTime.Parse("2018-02-31");
            install_1.delinquent_amount = 0.00M;
            install_1.install = 1;
            install_1.is_delinquent = 0;
            install_1.is_estimate = 0;
            install_1.is_exempt = 0;
            install_1.is_partial = 0;
            install_1.one_month = 0.00M;
            install_1.two_month = 0.00M;
            payment_record.installments.Add(install_1);

            PaymentInstallment install_2 = new PaymentInstallment();
            install_2.amount_due = 0.00M;
            install_2.base_amount = 0.00M;
            install_2.date_due = DateTime.Parse("2017-12-31");
            install_2.date_good_thru = DateTime.Parse("2018-02-31");
            install_2.date_paid = DateTime.Parse("2018-02-31");
            install_2.delinquent_amount = 0.00M;
            install_2.install = 1;
            install_2.is_delinquent = 0;
            install_2.is_estimate = 0;
            install_2.is_exempt = 0;
            install_2.is_partial = 0;
            install_2.one_month = 0.00M;
            install_2.two_month = 0.00M;
            payment_record.installments.Add(install_2);

            PaymentInstallment install_3 = new PaymentInstallment();
            install_3.amount_due = 0.00M;
            install_3.base_amount = 0.00M;
            install_3.date_due = DateTime.Parse("2017-12-31");
            install_3.date_good_thru = DateTime.Parse("2018-02-31");
            install_3.date_paid = DateTime.Parse("2018-02-31");
            install_3.delinquent_amount = 0.00M;
            install_3.install = 1;
            install_3.is_delinquent = 0;
            install_3.is_estimate = 0;
            install_3.is_exempt = 0;
            install_3.is_partial = 0;
            install_3.one_month = 0.00M;
            install_3.two_month = 0.00M;
            payment_record.installments.Add(install_3);

            PaymentInstallment install_4 = new PaymentInstallment();
            install_4.amount_due = 0.00M;
            install_4.base_amount = 0.00M;
            install_4.date_due = DateTime.Parse("2017-12-31");
            install_4.date_good_thru = DateTime.Parse("2018-02-31");
            install_4.date_paid = DateTime.Parse("2018-02-31");
            install_4.delinquent_amount = 0.00M;
            install_4.install = 1;
            install_4.is_delinquent = 0;
            install_4.is_estimate = 0;
            install_4.is_exempt = 0;
            install_4.is_partial = 0;
            install_4.one_month = 0.00M;
            install_4.two_month = 0.00M;
            payment_record.installments.Add(install_4);

            payment_record.tax_information = tax_info;
            parcel.payment_records.Add(payment_record);
            client_order.Parcels.Add(parcel);

            return client_order;
        }
    }
}
