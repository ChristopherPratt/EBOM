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
                string excelTemplate = System.AppDomain.CurrentDomain.BaseDirectory + "New EBOM " + template.time + ".xlsx";

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }
                xlWorkBooks = xlApp.Workbooks;
                try
                {
                    xlWorkBook = xlWorkBooks.Open(excelTemplate, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false); //open the template file!
                }
                catch (Exception e) { return; }
                xlWorkSheets = xlWorkBook.Worksheets;
                xlWorkSheet = xlWorkSheets.get_Item(1);

                //int totalRows = template.allCells.GetLength(0);
                //int totalColumns = template.allCells.GetLength(1);

                
                
                //for (int row = 0; row < totalRows ; row++)
                //{
                //    for (int column = 0; column < totalColumns; column++) // the plus one is because the excel columns and rows start at 1
                //    {
                //        if (template.allCells[row, column].text.Contains("[TextHere]")) writeData(template.allCells[row, column]);
                //    }
                //    Console.WriteLine("Finished writing row " + (row + 1));
                //}

                foreach (LoadTemplate.myCell cell in template.titleBlock) writeData(cell);
                foreach (LoadTemplate.myCell cell in template.headerRow) writeData(cell);
                Console.WriteLine("Finished writing Title Block");

                for (int a = 0; a < template.bodyRows.Count; a++)
                {
                    Range range = xlWorkSheet.Range[xlWorkSheet.Cells[template.bodyRows[a][0].rowIndex, template.bodyRows[a][0].columnIndex], xlWorkSheet.Cells[template.bodyRows[a][template.bodyRows[a].Count - 1].rowIndex, template.bodyRows[a][template.bodyRows[a].Count-1].columnIndex]]; // get whole row
                    range.Interior.Color = template.bodyColors[a % template.bodyColors.Count]; // change color of whole row
                    if (template.bodyRows[a][template.bodyRows[a].Count - 1].rightLineStyle != XlLineStyle.xlLineStyleNone) // check if we need to set borders
                    {
                        for (int b = 0; b < template.bodyRows[a].Count; b++)//set borders
                        {
                            xlWorkSheet.Cells[template.bodyRows[a][b].rowIndex, template.bodyRows[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeRight).LineStyle = template.bodyRows[a][b].rightLineStyle;
                            xlWorkSheet.Cells[template.bodyRows[a][b].rowIndex, template.bodyRows[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeRight).Weight = template.bodyRows[a][b].rightWeight;
                            xlWorkSheet.Cells[template.bodyRows[a][b].rowIndex, template.bodyRows[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeBottom).LineStyle = template.bodyRows[a][b].bottomLineStyle;
                            xlWorkSheet.Cells[template.bodyRows[a][b].rowIndex, template.bodyRows[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeBottom).Weight = template.bodyRows[a][b].bottomWeight;
                            xlWorkSheet.Cells[template.bodyRows[a][b].rowIndex, template.bodyRows[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeLeft).LineStyle = template.bodyRows[a][b].leftLineStyle;
                            xlWorkSheet.Cells[template.bodyRows[a][b].rowIndex, template.bodyRows[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeLeft).Weight = template.bodyRows[a][b].leftWeight;
                        }
                    }
                }


                int c = 0;
                foreach (List<List<LoadTemplate.myCell>> partGroups in XML.sorted)
                {        
                    foreach(List<LoadTemplate.myCell> cellRows in partGroups)
                    {
                        foreach (LoadTemplate.myCell cell in cellRows)
                            writeData(cell);
                    }
                    Console.WriteLine("Finished writing row: " + ++c);
                    
                }
                Console.WriteLine("Total Part Count: " + XML.totalPartCount);
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
            //if (!cell.moreThanText)
            //{
                xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex] = cell.info;
            //}
            //else
            //{
            //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex] = cell.info;
            //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Interior.Color = cell.color;
            //    if (cell.rightLineStyle != XlLineStyle.xlLineStyleNone) // only set borders if right border is different than default
            //    {
            //        xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeRight).LineStyle = cell.rightLineStyle;
            //        xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeRight).Weight = cell.rightWeight;
            //        xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeBottom).LineStyle = cell.bottomLineStyle;
            //        xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeBottom).Weight = cell.bottomWeight;
            //        xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeLeft).LineStyle = cell.leftLineStyle;
            //        xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeLeft).Weight = cell.leftWeight;
            //    }                
            //}            
        }

        //public void writeData(LoadTemplate.myCell cell)
        //{


        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Interior.Color = cell.color;
        //    //16777215 color of white cell
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].HorizontalAlignment = cell.horizontalAlignment;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].VerticalAlignment = cell.verticalAlignment;

        //    //if (cell.topLineStyle == XlLineStyle.xlLineStyleNone) cell.topWeight = XlBorderWeight.xlHairline;
        //    //if (cell.rightLineStyle == XlLineStyle.xlLineStyleNone) cell.rightWeight = XlBorderWeight.xlHairline;
        //    //if (cell.bottomLineStyle == XlLineStyle.xlLineStyleNone) cell.bottomWeight = XlBorderWeight.xlHairline;
        //    //if (cell.leftLineStyle == XlLineStyle.xlLineStyleNone) cell.leftWeight = XlBorderWeight.xlHairline;

        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeTop).LineStyle = cell.topLineStyle;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeTop).Weight = cell.topWeight;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeRight).LineStyle = cell.rightLineStyle;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeRight).Weight = cell.rightWeight;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeBottom).LineStyle = cell.bottomLineStyle;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeBottom).Weight = cell.bottomWeight;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeLeft).LineStyle = cell.leftLineStyle;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Borders(XlBordersIndex.xlEdgeLeft).Weight = cell.leftWeight;

        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex] = cell.text;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Name = cell.name;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Size = cell.size;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Color = cell.fontColor;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Bold = cell.bold;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Italic = cell.italic;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Underline = cell.underline;
        //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Strikethrough = cell.strikeThrough;
        //}

        //public void setRowsAndColumnDimensions(int totalRows, int totalColumns)
        //{

        //    for (int column = 0; column < totalColumns; column++)
        //    {
        //        xlWorkSheet.Rows[template.allCells[0, column].rowIndex].RowHeight = template.allCells[0, column].height; // sets cell height for just the 1st row
        //        xlWorkSheet.Columns[column + 1].ColumnWidth = template.allCells[0, column].width; // sets cell width for every column
        //    }
        //    Console.WriteLine("Finished Setting Column Widths");

        //    for (int row = 1; row < totalRows; row++)
        //    {
        //        xlWorkSheet.Rows[row].RowHeight = template.allCells[row, totalColumns - 1].height; // sets the cell height for every column except row 1
        //    }
        //    Console.WriteLine("Finished Setting row heights");
        //}
        //public void mergeCells()
        //{
        //    foreach (List<int> temp in template.mergedArea)
        //    {
        //        Range range = xlWorkSheet.Range[xlWorkSheet.Cells[temp[0], temp[1]], xlWorkSheet.Cells[temp[2], temp[3]]];
               

        //        range.Merge(false);
        //    }
        
        //}

        private string getTime()
        {
            string datePatt = @"hh.mm.ss.ff";
            DateTime saveUtcNow = DateTime.UtcNow;
            return saveUtcNow.ToString(datePatt);
        }

        //public CreateExcelFile(LoadXML x, LoadTemplate t)
        //{
        //    XML = x;
        //    template = t;
        //    try
        //    {
        //        //Start the Microsoft Excel Application

        //        xlApp = new Microsoft.Office.Interop.Excel.Application();
        //        if (xlApp == null)
        //        {
        //            MessageBox.Show("Excel is not properly installed!!");
        //            return;
        //        }
        //        xlWorkBooks = xlApp.Workbooks;
        //        xlWorkBook = xlWorkBooks.Add();
        //        xlWorkSheets = xlWorkBook.Worksheets;
        //        xlWorkSheet = xlWorkSheets.get_Item(1);

        //        int totalRows = template.allCells.GetLength(0);
        //        int totalColumns = template.allCells.GetLength(1);



        //        for (int row = 0; row < totalRows; row++)
        //        {
        //            for (int column = 0; column < totalColumns; column++) // the plus one is because the excel columns and rows start at 1
        //            {
        //                writeData(template.allCells[row, column]);
        //            }
        //            Console.WriteLine("Finished writing row " + (row + 1));
        //        }

        //        setRowsAndColumnDimensions(totalRows, totalColumns);
        //        mergeCells();

        //        xlWorkBook.SaveAs(System.AppDomain.CurrentDomain.BaseDirectory + "Ebom_tesbvting" + getTime() + ".xlsx");

        //    }
        //    finally
        //    {
        //        Console.WriteLine("Finished Saving Excel File");
        //        Marshal.FinalReleaseComObject(xlWorkSheet);
        //        Marshal.FinalReleaseComObject(xlWorkSheets);
        //        xlWorkBook.Close();
        //        Marshal.FinalReleaseComObject(xlWorkBook);
        //        xlWorkBooks.Close();
        //        Marshal.FinalReleaseComObject(xlWorkBooks);
        //        xlApp.Quit();
        //        Marshal.FinalReleaseComObject(xlApp); // excel objects don't releast comObjects to excel so you have to force it
        //    }
        //}


    }
}
