using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Core;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Xml;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;

namespace WindowsFormsApp1
{
    class CreateExcelFile
    {


        string time;

        XmlNodeList nodeList;
        LoadXML XML;
        LoadTemplate template;

        Microsoft.Office.Interop.Excel.Application xlApp;
        Workbooks xlWorkBooks;
        Workbook xlWorkBook;
        Sheets xlWorkSheets;
        Worksheet xlWorkSheet;
        public CreateExcelFile(LoadXML x, LoadTemplate t)
        {
            XML = x;
            template = t;
            try
            {
                //Start the Microsoft Excel Application

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }
                xlWorkBooks = xlApp.Workbooks;
                xlWorkBook = xlWorkBooks.Add();
                xlWorkSheets = xlWorkBook.Worksheets;
                xlWorkSheet = xlWorkSheets.get_Item(1);

                int totalRows = template.allCells.GetLength(0);
                int totalColumns = template.allCells.GetLength(1);

                
                
                for (int row = 0; row < totalRows ; row++)
                {
                    for (int column = 0; column < totalColumns; column++) // the plus one is because the excel columns and rows start at 1
                    {
                        writeData(template.allCells[row, column]);
                    }
                    Console.WriteLine("Finished writing row " + row);
                }

                setRowsAndColumnDimensions(totalRows, totalColumns);
                mergeCells();

                xlWorkBook.SaveAs(System.AppDomain.CurrentDomain.BaseDirectory + "Ebom_testing" + getTime() + ".xlsx");

            }
            finally
            {
                Console.WriteLine("Finished Saving Excel File");
                Marshal.FinalReleaseComObject(xlWorkSheet);
                Marshal.FinalReleaseComObject(xlWorkSheets);
                xlWorkBook.Close();
                Marshal.FinalReleaseComObject(xlWorkBook);
                xlWorkBooks.Close();
                Marshal.FinalReleaseComObject(xlWorkBooks);
                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp); // excel objects don't releast comObjects to excel so you have to force it
            }
        }

        public void writeData(LoadTemplate.myCell cell)
        {


            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Interior.Color = cell.color;
            //16777215 color of white cell
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].HorizontalAlignment = cell.horizontalAlignment;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].VerticalAlignment = cell.verticalAlignment;

            if (cell.topLineStyle == XlLineStyle.xlLineStyleNone) cell.topWeight = XlBorderWeight.xlHairline;
            if (cell.rightLineStyle == XlLineStyle.xlLineStyleNone) cell.rightWeight = XlBorderWeight.xlHairline;
            if (cell.bottomLineStyle == XlLineStyle.xlLineStyleNone) cell.bottomWeight = XlBorderWeight.xlHairline;
            if (cell.leftLineStyle == XlLineStyle.xlLineStyleNone) cell.leftWeight = XlBorderWeight.xlHairline;

            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeTop).LineStyle = cell.topLineStyle;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeTop).Weight = cell.topWeight;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeRight).LineStyle = cell.rightLineStyle;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeRight).Weight = cell.rightWeight;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeBottom).LineStyle = cell.bottomLineStyle;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeBottom).Weight = cell.bottomWeight;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeLeft).LineStyle = cell.leftLineStyle;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeLeft).Weight = cell.leftWeight;

            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex] = cell.text;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Name = cell.name;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Size = cell.size;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Color = cell.fontColor;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Bold = cell.bold;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Italic = cell.italic;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Underline = cell.underline;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Strikethrough = cell.strikeThrough;
        }

        public void setRowsAndColumnDimensions(int totalRows, int totalColumns)
        {

            for (int column = 0; column < totalColumns; column++)
            {
                xlWorkSheet.Rows[template.allCells[0, column].rowIndex].RowHeight = template.allCells[0, column].height; // sets cell height for just the 1st row
                xlWorkSheet.Columns[column + 1].ColumnWidth = template.allCells[0, column].width; // sets cell width for every column
            }
            Console.WriteLine("Finished Setting Column Widths");

            for (int row = 1; row < totalRows; row++)
            {
                xlWorkSheet.Rows[row].RowHeight = template.allCells[row, totalColumns - 1].height; // sets the cell height for every column except row 1
            }
            Console.WriteLine("Finished Setting row heights");
        }
        public void mergeCells()
        {
            foreach (List<int> temp in template.mergedArea)
            {
                Range range = xlWorkSheet.Range[xlWorkSheet.Cells[temp[0], temp[1]], xlWorkSheet.Cells[temp[2], temp[3]]];
               

                range.Merge(false);
            }
        
        }

        private string getTime()
        {
            string datePatt = @"hh.mm.ss.ff";
            DateTime saveUtcNow = DateTime.UtcNow;
            return saveUtcNow.ToString(datePatt);
        }


    }
}
