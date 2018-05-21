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
        private double default_row_height = 15;
        int currentrow = 0;

        public bool CreateReport()
        {
            foreach(ParcelInformation parcel in this.client_order.Parcels)
            {
                Filepath();
                NewExcel();
                SetMargins();
                Header(parcel);
                SaveExcel();
                CloseExcel();
            }
            

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

        private void ColumnWidths()
        {
            sheet1.Columns["A:A"].ColumnWidth = 11.57;
            sheet1.Columns["B:B"].ColumnWidth = 8.43;
            sheet1.Columns["C:C"].ColumnWidth = 9.57;
            sheet1.Columns["D:D"].ColumnWidth = 10.14;
            sheet1.Columns["E:E"].ColumnWidth = 8.14;
            sheet1.Columns["F:F"].ColumnWidth = 7.86;
            sheet1.Columns["G:G"].ColumnWidth = 8.43;
            sheet1.Columns["H:H"].ColumnWidth = 9.57;
            sheet1.Columns["I:I"].ColumnWidth = 8.86;
            sheet1.Columns["J:J"].ColumnWidth = 6;
            sheet1.Columns["K:K"].ColumnWidth = 8.43;
        }
        
        private void Header(ParcelInformation parcel)
        {
            currentrow++;
            sheet1.Range["A" + currentrow, "K" + currentrow].Merge();
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            this.CellValue("A" + currentrow, "Tax Certification", 10, Excel.XlHAlign.xlHAlignCenter, font: "Calibri");

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["E" + currentrow, "F" + currentrow].Merge();
            sheet1.Range["G" + currentrow, "H" + currentrow].Merge();
            this.CellValue("E" + currentrow, "Verified as of:", 10, Excel.XlHAlign.xlHAlignRight, font: "Calibri");
            this.CellValue("G" + currentrow, DataFunctions.DateToString(parcel.effective_date));

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = 6.75;

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "F" + currentrow].Merge();
            sheet1.Range["H" + currentrow, "I" + currentrow].Merge();
            sheet1.Range["J" + currentrow, "K" + currentrow].Merge();
            this.CellValue("A" + currentrow, "PO Number:", 8, font: "Calibri");
            this.CellValue("C" + currentrow, parcel.client_po_number, 8);
            this.CellValue("H" + currentrow, "Assessed Valuation:", 8, font: "Calibri");
            this.CellValue("J" + currentrow, string.Format("{0:C}", parcel.assessed_valuation), 8, Excel.XlHAlign.xlHAlignRight);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "F" + currentrow].Merge();
            sheet1.Range["I" + currentrow, "K" + currentrow].Merge();
            this.CellValue("A" + currentrow, "Property Owner:", 8, font: "Calibri");
            this.CellValue("C" + currentrow, parcel.assessed_owners, 8);
            this.CellValue("H" + currentrow, "County:", 8, font: "Calibri");
            this.CellValue("I" + currentrow, parcel.county, 8, Excel.XlHAlign.xlHAlignCenter);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "F" + currentrow].Merge();
            this.CellValue("A" + currentrow, "Tax Address:", 8, font: "Calibri");
            this.CellValue("C" + currentrow, parcel.assessed_address, 8);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "F" + currentrow].Merge();
            this.CellValue("A" + currentrow, "Town/City:", 8, font: "Calibri");
            this.CellValue("C" + currentrow, DataFunctions.TownCity(parcel.payment_records), 8);



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
