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

namespace EBOMCreationTool
{
    class CreateExcelFile
    {


        

        XmlNodeList nodeList;
        LoadXML XML;
        LoadTemplate template;

        string ExportFileName;

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
                //string excelTemplate = System.AppDomain.CurrentDomain.BaseDirectory + "template.xlsx";
                ExportFileName = System.IO.Path.ChangeExtension(XML.xmlFile, null) + ".xlsx";


                xlApp = template.xlAppOpen;
                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }
                xlWorkBooks = template.xlWorkBooks2;
                try
                {
                    xlWorkBook = template.xlWorkBook2;
                }
                catch (Exception e) { return; }
                xlWorkSheet = template.xlWorkSheet2;

                xlWorkSheet.Cells[1, template.columnEnd+1] = "";
                xlWorkSheet.Cells[1, template.columnEnd] = "";
                xlWorkSheet.Cells[template.rowEnd, 1] = "";
                foreach (LoadTemplate.myCell cell in template.titleBlock) writeData(cell);
                foreach (LoadTemplate.myCell cell in template.headerRow) writeData(cell);
                for (int a = template.bodyRowStart; a < template.rowEnd; a++)
                {
                    for (int b = 1; b < template.columnEnd; b++)
                    {
                        xlWorkSheet.Cells[a, b] = ""; // clear all artifacts from body and footer section.
                        xlWorkSheet.Cells[a, b].Interior.Color = 16777215;
                    }
                }
                Console.WriteLine("Finished writing Title Block");

                for (int a = 0; a < XML.sorted.Count; a++)
                {
                    Range range = xlWorkSheet.Range[xlWorkSheet.Cells[XML.sorted[a][0].rowIndex, XML.sorted[a][0].columnIndex], xlWorkSheet.Cells[XML.sorted[a][XML.sorted[a].Count - 1].rowIndex, XML.sorted[a][XML.sorted[a].Count-1].columnIndex]]; // get whole row
                    range.Interior.Color = template.bodyColors[a % template.bodyColors.Count]; // change color of whole row
                    if (XML.sorted[a][XML.sorted[a].Count - 1].rightLineStyle != XlLineStyle.xlLineStyleNone) // check if we need to set borders
                    {
                        for (int b = 0; b < template.bodyRows[a].Count; b++)//set borders
                        {
                            xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeRight).LineStyle = XML.sorted[a][b].rightLineStyle;
                            xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeRight).Weight = XML.sorted[a][b].rightWeight;
                            xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XML.sorted[a][b].bottomLineStyle;
                            xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeBottom).Weight = XML.sorted[a][b].bottomWeight;
                            xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XML.sorted[a][b].leftLineStyle;
                            xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeLeft).Weight = XML.sorted[a][b].leftWeight;
                        }
                    }
                }


                int c = 0;
                foreach (List<LoadTemplate.myCell> cellRows in XML.sorted)
                {        
                    foreach(LoadTemplate.myCell cell in cellRows) writeData(cell);
                    Console.WriteLine("Finished writing row: " + ++c);
                }
                Console.WriteLine("Total Part Count: " + XML.totalPartCount);
                if (c != XML.totalPartCount) MessageBox.Show("WARNING!\nNot all parts exported to BOM from xml.");
                //xlWorkBook.SaveAs(System.AppDomain.CurrentDomain.BaseDirectory + "New EBOM1 " + template.time + ".xlsx");
                bool excelFileIsOpen = false;
                do
                {
                    try
                    {
                        xlWorkBook.SaveAs(ExportFileName);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Please close the Excel file with the same name as the .XML file that was chosen.");
                        excelFileIsOpen = true;
                        continue;
                    }
                    excelFileIsOpen = false;
                }
                while (excelFileIsOpen);
                

            }
            finally
            {
                Console.WriteLine("Finished Saving Excel File");
                Marshal.FinalReleaseComObject(xlWorkSheet);
                //Marshal.FinalReleaseComObject(xlWorkSheets);
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

            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex] = cell.info;
        }


        private string getTime()
        {
            string datePatt = @"hh.mm.ss.ff";
            DateTime saveUtcNow = DateTime.UtcNow;
            return saveUtcNow.ToString(datePatt);
        }

     

    }
}
