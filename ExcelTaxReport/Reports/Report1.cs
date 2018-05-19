using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelTaxReport;
using ExcelTaxReport.Models;
using System.Runtime.InteropServices;
using System.Diagnostics;


namespace ExcelTaxReport.Reports
{
    class Report1 : IReport
    {
        public Report1(ClientOrder client_order)
        {
            
        }

        public bool CreateReport()
        {
            return false;
        }

    }
}
