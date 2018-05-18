using System;
using System.Collections.Generic;

namespace ReportBuilder.Models
{
    public class ClientOrder
    {
        public ReportConfig report_config { get; set; }
        public ClientConfig client_config { get; set; }

        public ClientOrder()
        {
            this.report_config = new ReportConfig();
            this.client_config = new ClientConfig();
        }
    }
}
