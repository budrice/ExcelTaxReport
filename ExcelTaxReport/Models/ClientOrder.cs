using System.Collections.Generic;

namespace ExcelTaxReport.Models
{
    /// <summary>
    /// client order
    /// </summary>
    /// <remarks>
    /// A client order contains the List of parcels. Each parcel will be a separate document.
    /// </remarks>
    public class ClientOrder
    {
        public ReportConfig report_config { get; set; }
        public ClientConfig client_config { get; set; }
        public List<ParcelInformation> Parcels { get; set; }

        public ClientOrder()
        {
            this.report_config = new ReportConfig();
            this.client_config = new ClientConfig();
            this.Parcels = new List<ParcelInformation>();
        }
    }
}
