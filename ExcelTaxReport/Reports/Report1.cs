using ExcelTaxReport;
using ExcelTaxReport.Models;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;




namespace ExcelTaxReport.Reports
{
    public class Report1: WriteExcel, IReport
    {
        public Report1(ClientOrder client_order)
        {
            this.client_order = client_order;
        }

        object misValue = System.Reflection.Missing.Value;
        private string filepath = string.Empty;
        private bool gridlines = true;

        public bool CreateReport()
        {
            Filepath();
            NewExcel();
            SetMargins();

            SaveExcel();
            CloseExcel();

            return false;
        }

        private void Filepath()
        {
            ClientConfig config = this.client_order.client_config;
            filepath = config.base_path + config.report_name;
        }

        private void NewExcel()
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Add(misValue);
            sheet1 = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            sheet1.Name = "Tax Research";
        }

        private void SetMargins()
        {
            xlApp.Windows.Application.ActiveWindow.DisplayGridlines = gridlines;
            sheet1.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperLegal;
            sheet1.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            sheet1.PageSetup.LeftMargin = 10.0;
            sheet1.PageSetup.RightMargin = 10.0;
            sheet1.PageSetup.TopMargin = 30.0;
            sheet1.PageSetup.BottomMargin = 0.2;
            sheet1.PageSetup.Zoom = false;
            sheet1.PageSetup.FitToPagesTall = 1;
        }

        private void SaveExcel()
        {
            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(filepath, Excel.XlFileFormat.xlOpenXMLWorkbook);
            xlApp.DisplayAlerts = true;

        }

        private void CloseExcel()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            xlWorkbook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.FinalReleaseComObject(sheet1);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
