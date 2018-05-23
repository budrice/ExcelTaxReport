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

        object misValue = System.Reflection.Missing.Value;// used to create the worksheet
        private string filepath = string.Empty;
        private bool gridlines = true;
        private double default_row_height = 15;
        int currentrow = 0;
        int delq_section_flag = 0;

        public bool CreateReport()
        {
            foreach(ParcelInformation parcel in client_order.Parcels)
            {
                Filepath();
                NewExcel();
                SetMargins();
                Header(parcel);
                Content(parcel);
                SaveExcel();
                CloseExcel();
            }
            

            return false;
        }

        private void Filepath()
        {
            ClientConfig config = client_order.client_config;
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
            DrawBorder("A" + currentrow, "K" + (currentrow + 11));
            sheet1.Range["A" + currentrow, "K" + currentrow].Merge();
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            CellValue("A" + currentrow, "Tax Certification", 10, Excel.XlHAlign.xlHAlignCenter, font: "Calibri");

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["E" + currentrow, "F" + currentrow].Merge();
            sheet1.Range["G" + currentrow, "H" + currentrow].Merge();
            CellValue("E" + currentrow, "Verified as of:", 10, Excel.XlHAlign.xlHAlignRight, font: "Calibri");
            CellValue("G" + currentrow, DataFunctions.DateToString(parcel.effective_date));

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = 6.75;

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "F" + currentrow].Merge();
            sheet1.Range["H" + currentrow, "I" + currentrow].Merge();
            sheet1.Range["J" + currentrow, "K" + currentrow].Merge();
            CellValue("A" + currentrow, "PO Number:", 8, font: "Calibri");
            CellValue("C" + currentrow, parcel.client_po_number, 8);
            CellValue("H" + currentrow, "Assessed Valuation:", 8, font: "Calibri");
            CellValue("J" + currentrow, string.Format("{0:C}", parcel.assessed_valuation), 8, Excel.XlHAlign.xlHAlignRight);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "F" + currentrow].Merge();
            sheet1.Range["I" + currentrow, "K" + currentrow].Merge();
            CellValue("A" + currentrow, "Property Owner:", 8, font: "Calibri");
            CellValue("C" + currentrow, parcel.assessed_owners, 8);
            CellValue("H" + currentrow, "County:", 8, font: "Calibri");
            CellValue("I" + currentrow, parcel.county, 8, Excel.XlHAlign.xlHAlignCenter);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "F" + currentrow].Merge();
            CellValue("A" + currentrow, "Tax Address:", 8, font: "Calibri");
            CellValue("C" + currentrow, parcel.assessed_address, 8);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "F" + currentrow].Merge();
            CellValue("A" + currentrow, "Town/City:", 8, font: "Calibri");
            CellValue("C" + currentrow, DataFunctions.TownCity(parcel.payment_records), 8);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "F" + currentrow].Merge();
            CellValue("A" + currentrow, "Parcel ID:", 8, font: "Calibri");
            CellValue("C" + currentrow, parcel.parcel_number, 8);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "F" + currentrow].Merge();
            sheet1.Range["I" + currentrow, "K" + currentrow].Merge();
            CellValue("A" + currentrow, "School District:", 8, font: "Calibri");
            CellValue("C" + currentrow, DataFunctions.SchoolDistrict(parcel.payment_records), 8);
            CellValue("H" + currentrow, "Class Code:", 8, font: "Calibri");
            CellValue("I" + currentrow, parcel.class_code, 8);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = 9;

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = default_row_height;
            sheet1.Range["B" + currentrow, "C" + currentrow].Merge();
            sheet1.Range["E" + currentrow, "K" + currentrow].Merge();
            DrawGrid("B" + currentrow, "C" + currentrow);
            CellValue("A" + currentrow, "Exemptions:", 8, font: "Calibri");
            Checkboxes(currentrow, "B", DataFunctions.HasExemptions(parcel.payment_records));
            CellValue("D" + currentrow, "Description:", 8, font: "Calibri");
            CellValue("E" + currentrow, DataFunctions.ExemptionString(parcel.payment_records, parcel.state), 8);
        }

        private void Content(ParcelInformation parcel)
        {
            foreach(TaxAuthorityPaymentRecord record in parcel.payment_records)
            {
                Valuation(record, parcel);
                if (delq_section_flag == 0)
                {
                    Delinquent(record);
                    delq_section_flag++;
                }
                PaymentInstallments(record);
            }
        }

        private void Valuation(TaxAuthorityPaymentRecord record, ParcelInformation parcel)
        {
            string first_six = parcel.client_po_number.Substring(0, 6);
            string[] array = new string[3];
            array[0] = "NTN";
            array[1] = "TCTI";
            array[2] = "CTCSD";

            currentrow++;
            if (DataFunctions.StrInString(first_six, array))
            {
                ColorMergedRow(currentrow, 7.5);

                currentrow++;
                sheet1.Cells[currentrow, "A"].RowHeight = 15;
                sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
                sheet1.Range["H" + currentrow, "I" + currentrow].Merge();
                sheet1.Range["J" + currentrow, "K" + currentrow].Merge();
                CellValue("A" + currentrow, "Millage Rate Information", 8, font: "Calibri");
                DrawBorder("A" + currentrow, "B" + currentrow);
                CellValue("C" + currentrow, "Millage Rate:", 8, Excel.XlHAlign.xlHAlignRight, font: "Calibri");
                CellValue("D" + currentrow, String.Format("{0:C}", record.tax_information.milage_rate), 8);
                DrawBorder("A" + currentrow, "B" + currentrow);
                DrawBorder("A" + (currentrow + 1), "B" + (currentrow + 1));
                DrawBorder("C" + currentrow, "D" + (currentrow + 1));
                DrawBorder("E" + currentrow, "G" + (currentrow + 1));
                DrawBorder("H" + currentrow, "I" + (currentrow + 1));
                DrawBorder("J" + currentrow, "K" + (currentrow + 1));

                CellValue("H" + currentrow, "Assessed Value:", 8, Excel.XlHAlign.xlHAlignRight, font: "Calibri");
                CellValue("J" + currentrow, String.Format("{0:C}", record.assessed_value), 8);

                currentrow++;
                sheet1.Cells[currentrow, "A"].RowHeight = 15;
                CellValue("A" + currentrow, "Next due Date:", 8, Excel.XlHAlign.xlHAlignRight, font: "Calibri");
                CellValue("B" + currentrow, DataFunctions.DateToString(record.tax_information.milage_next_due), 8);
                CellValue("C" + currentrow, "Land:", 8, Excel.XlHAlign.xlHAlignRight, font: "Calibri");
                CellValue("D" + currentrow, String.Format("{0:C}", record.land_value), 8);
                sheet1.Range["E" + currentrow, "F" + currentrow].Merge();
                CellValue("E" + currentrow, "Improvement:", 8, Excel.XlHAlign.xlHAlignCenter, font: "Calibri");
                CellValue("G" + currentrow, String.Format("{0:C}", record.improved_value), 6);
                CellValue("I" + currentrow, "Total:", 8, Excel.XlHAlign.xlHAlignCenter, font: "Calibri");
                sheet1.Range["J" + currentrow, "K" + currentrow].Merge();
                CellValue("J" + currentrow, String.Format("{0:C}", record.total_value), 8);
            }
            else
            {
                ColorMergedRow(currentrow, 7.5);
            }
        }

        private void Delinquent(TaxAuthorityPaymentRecord record)
        {
            currentrow++;
            this.ColorMergedRow(currentrow, 7.5);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = 15;
            sheet1.Range["A" + currentrow, "F" + currentrow].Merge();
            CellValue("A" + currentrow, "DELINQUENT TAXES", 8, Excel.XlHAlign.xlHAlignCenter, font: "Cambri", bold: true, underline: true);
            CellValue("H" + currentrow, "Payable to:", 8, font: "Calibri", bold: true, underline: true);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = 15;
            sheet1.Range["A" + currentrow, "C" + currentrow].Merge();
            sheet1.Range["D" + currentrow, "G" + currentrow].Merge();
            sheet1.Range["H" + currentrow, "K" + currentrow].Merge();
            CellValue("A" + currentrow, "Delinquencies have been verified with:", 8, font: "Calibri");
            CellValue("H" + currentrow, record.tax_authority.payment_string_address, 8);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = 15;
            sheet1.Range["A" + currentrow, "B" + currentrow].Merge();
            sheet1.Range["C" + currentrow, "D" + currentrow].Merge();
            sheet1.Range["H" + currentrow, "K" + currentrow].Merge();
            CellValue("A" + currentrow, "TAXES DELINQUENT?", 8, font: "Calibri");
            Checkboxes(currentrow, "C", DataFunctions.IsDelinquent(record.installments));
            CellValue("H" + currentrow, record.tax_authority.payment_city_state_zip, 8);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = 15;
            sheet1.Range["A" + currentrow, "C" + currentrow].Merge();
            sheet1.Range["D" + currentrow, "G" + currentrow].Merge();
            sheet1.Range["H" + currentrow, "K" + currentrow].Merge();
            CellValue("A" + currentrow, "Description of delinquent taxes:", 8, font: "Calibri");
            

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = 15;
            sheet1.Range["A" + currentrow, "G" + currentrow].Merge();
            sheet1.Range["I" + currentrow, "K" + currentrow].Merge();
            sheet1.Cells[currentrow, "A"].WrapText = true;
            CellValue("H" + currentrow, "Phone:", 8, font: "Calibri");
            CellValue("I" + currentrow, record.tax_authority.payment_phone_string, 8);

            DrawBorder("A" + (currentrow - 5), "G" + (currentrow - 2));
            DrawBorder("H" + (currentrow - 5), "K" + (currentrow - 1));
            DrawBorder("A" + (currentrow -1), "G" + currentrow);
            DrawBorder("H" + currentrow, "K" + currentrow);
            DrawBorder("C" + (currentrow - 2), "D" + (currentrow - 2));

            if (String.Compare(DataFunctions.IsDelinquent(record.installments, false), "Yes") == 0)
            {
                string desc = DataFunctions.DelinquencyDescription(record.installments);
                CellValue("H" + (currentrow - 3), record.tax_authority.payment_string_address, 8);
                CellValue("H" + (currentrow - 2), record.tax_authority.payment_city_state_zip, 8);
                CellValue("I" + currentrow, record.tax_authority.payment_phone_string, 8);

                CellValue("A" + currentrow, desc, 8);
                sheet1.Cells[currentrow, "A"].WrapText = true;
                var height = DataFunctions.StringHeight(desc, 8, 600);
                sheet1.Cells[currentrow, "A"].RowHeight = height;
                CellValue("H" + (currentrow - 3), record.tax_authority.payment_string_address, 8);
            }
        }

        private void PaymentInstallments(TaxAuthorityPaymentRecord record)
        {
            currentrow++;
            this.ColorMergedRow(currentrow, 7.5);

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = 15;
            sheet1.Range["D" + currentrow, "E" + currentrow].Merge();
            sheet1.Range["H" + currentrow, "I" + currentrow].Merge();
            sheet1.Range["J" + currentrow, "K" + currentrow].Merge();
            CellValue("A" + currentrow, "Tax Year:", 8, font: "Calibri");
            CellValue("B" + currentrow, record.installments[0].year, 8);
            CellValue("C" + currentrow, "Tax Type:", 8, font: "Calibri");
            CellValue("D" + currentrow, record.tax_type, 8);
            CellValue("F" + currentrow, "Fiscal Year:", 6, font: "Calibri");
            CellValue("G" + currentrow, record.tax_authority.fiscal_year, 6, font: "Calibri");
            CellValue("H" + currentrow, "Installment Info:", 8, font: "Calibri");
            CellValue("J" + currentrow, record.tax_authority.installment_info, 8, font: "Calibri");

            currentrow++;
            sheet1.Cells[currentrow, "A"].RowHeight = 15;
            sheet1.Range["B" + currentrow, "K" + currentrow].Merge();
            CellValue("A" + currentrow, "Total Tax Billed:");
            CellValue("B" + currentrow, string.Format("{0:C}", DataFunctions.TotalBilled(record.installments)));
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
