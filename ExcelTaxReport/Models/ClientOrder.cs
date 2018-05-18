using System;
using System.Collections.Generic;

namespace ReportBuilder.Models
{
    public class ClientOrder
    {
        public ReportConfig report_config { get; set; }

        public ClientOrder()
        {
            this.report_config = new ReportConfig();
        }
    }
}
