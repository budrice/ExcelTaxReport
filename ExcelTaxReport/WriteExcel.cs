using System;
using System.Drawing;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;
using ExcelTaxReport.Models;
using ExcelTaxReport.Reports;
namespace ExcelTaxReport
{
    public class WriteExcel
    {
        public ClientOrder client_order;

        public Excel.Application xlApp;
        public Excel.Workbook xlWorkbook;
        public Excel.Workbooks xlWorkbooks;
        public Excel.Worksheet sheet1;
        public Excel.Worksheet sheet2;
        public Excel.Worksheet sheet3;

        public void ColorMergedRow(int row, double height = 15, string start_col = "A", string end_col = "K", Color? color = null, bool border = true)
        {
            Color rgb = color ?? Color.FromArgb(192, 192, 192);
            sheet1.Range[start_col + row, end_col + row].Merge();
            if (border == true)
            {
                sheet1.Range[start_col + row, end_col + row].Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }
            sheet1.Cells[row, start_col].Interior.Color = ColorTranslator.ToOle(rgb);
            sheet1.Cells[row, start_col].RowHeight = height;
        }

        public void CellValue(string cell, string value, int fontsize = 10, Excel.XlHAlign halign = Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign valign = Excel.XlVAlign.xlVAlignBottom, Color? font_color = null, string font = "Arial", bool bold = false, bool underline = false)
        {
            int row;
            Int32.TryParse(new String(cell.Where(Char.IsDigit).ToArray()), out row);
            string col = new String(cell.Where(Char.IsLetter).ToArray());
            Color font_rgb = font_color ?? Color.FromArgb(0, 0, 0);

            sheet1.Cells[row, col].WrapText = true;
            sheet1.Cells[row, col].Font.Name = font;
            sheet1.Cells[row, col].Font.Size = fontsize;
            sheet1.Cells[row, col].Font.Bold = bold;
            sheet1.Cells[row, col].Font.Underline = underline;
            sheet1.Cells[row, col].Font.Color = ColorTranslator.ToOle(font_rgb);
            sheet1.Cells[row, col] = value;
            sheet1.Cells[row, col].HorizontalAlignment = halign;
            sheet1.Cells[row, col].VerticalAlignment = valign;
        }

        public void Checkboxes(int row, string col, string val, int fontsize = 10)
        {
            int indice_of_first_space = val.IndexOf(" ", 0);
            int indice_of_second_space = val.IndexOf(" ", indice_of_first_space + 1);
            int string_length = val.Length;
            int char_length_between_checkboxes = indice_of_first_space + indice_of_second_space;
            sheet1.Cells[row, col] = val;
            sheet1.Cells[row, col].Font.Size = fontsize;
            sheet1.Cells[row, col].Characters(0, 1).Font.Name = "Wingdings";
            sheet1.Cells[row, col].Characters(2, char_length_between_checkboxes).Font.Name = "Arial";
            sheet1.Cells[row, col].Characters(char_length_between_checkboxes + 1, 1).Font.Name = "Wingdings";
            sheet1.Cells[row, col].Characters(char_length_between_checkboxes + 2, string_length).Font.Name = "Arial";
        }

    }
}
