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
    class LoadTemplate
    {
        public class myCell
        {
            public int rowIndex { get; set; }
            public int columnIndex { get; set; }

            //public class myFont
            //{
            //    public System.Object name { get; set; }
            //    public System.Object size { get; set; }
            //    public System.Object color { get; set; }
            //    public System.Object bold { get; set; }
            //    public System.Object italic { get; set; }
            //    public System.Object underline { get; set; }
            //    public System.Object strikeThrough { get; set; }                
            //}

            //public List<myFont> complexWords;

            public System.Object name { get; set; }
            public System.Object size { get; set; }
            public System.Object fontColor { get; set; }
            public System.Object bold { get; set; }
            public System.Object italic { get; set; }
            public System.Object underline { get; set; }
            public System.Object strikeThrough { get; set; }

            public System.Object horizontalAlignment;
            public System.Object verticalAlignment;

            public string text { get; set; }
            public double width { get; set; }
            public double height { get; set; }
            public double color { get; set; }

            public bool merge { get; set; }
            //public Border topLineStyle { get; set; }
            //public Border topWeight { get; set; }
            //public Border rightLineStyle { get; set; }
            //public Border rightWeight { get; set; }
            //public Border bottomLineStyle { get; set; }
            //public Border bottomWeight { get; set; }
            //public Border leftLineStyle { get; set; }
            //public Border leftWeight { get; set; }
            public XlLineStyle topLineStyle { get; set; }
            public XlBorderWeight topWeight { get; set; }
            public XlLineStyle rightLineStyle { get; set; }
            public XlBorderWeight rightWeight { get; set; }
            public XlLineStyle bottomLineStyle { get; set; }
            public XlBorderWeight bottomWeight { get; set; }
            public XlLineStyle leftLineStyle { get; set; }
            public XlBorderWeight leftWeight { get; set; }
        }

        public myCell[,] allCells;

        Microsoft.Office.Interop.Excel.Application xlAppOpen;
        Workbooks xlWorkBooks2;
        Workbook xlWorkBook2;
        Worksheet xlWorkSheet2;

        public LoadTemplate()
        {

            //time = getTime();
            Thread gui = new Thread(delegate ()
            {
                try
                {
                    readExcelFile(1000,1000); // choosing unthinkably huge number since i want to be able to cover any size template
                }                               // breaks in the loops that use those numbers prevent inefficiency.
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            
            //    try
            //    {
            //        saveExcelFile();
            //    }
            //    finally
            //    {
            //        GC.Collect();
            //        GC.WaitForPendingFinalizers();
            //    }
            });
            gui.Name = "gui";
            gui.Start();



        }
        public void readExcelFile(int totalRows, int totalColumns)
        {
            //allCells = new myCell[totalRows, totalColumns];

            try
            {
                //string excelTemplate = System.AppDomain.CurrentDomain.BaseDirectory + "Ebom_testing" + ".xlsx";
                string excelTemplate = System.AppDomain.CurrentDomain.BaseDirectory + "template" + ".xlsx";

                if (!File.Exists(excelTemplate))
                {
                    throw new Exception("Excel template not found in " + excelTemplate);
                }
                xlAppOpen = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBooks2 = xlAppOpen.Workbooks;

                try
                {
                    xlWorkBook2 = xlWorkBooks2.Open(excelTemplate, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false); //open the template file!
                }
                catch (Exception e) { return; }
                xlWorkSheet2 = (Worksheet)xlWorkBook2.Worksheets.get_Item(1); //worksheet to write data to

                getCells(totalRows, totalColumns);


            }
            finally
            {
                Console.WriteLine("Finished reading Excel File");
                Marshal.FinalReleaseComObject(xlWorkSheet2);
                xlWorkBook2.Close();
                Marshal.FinalReleaseComObject(xlWorkBook2);
                xlWorkBooks2.Close();
                Marshal.FinalReleaseComObject(xlWorkBooks2);
                xlAppOpen.Quit();
                Marshal.FinalReleaseComObject(xlAppOpen); // excel objects don't release comObjects to excel so you have to force it
            }
        }

        public void getCells(int totalRows, int totalColumns)
        {
            int cellBuffer = 5;
            bool firstRowFound = false;
            int longestRow = 0;
            int emptyColumnCellCount = 0;
            int emptyRowCellCount = 0;
            List<List<myCell>> tempAllCells = new List<List<myCell>>();
            List<myCell> temp;
            myCell tempCell;
            for (int row = 1; row < totalRows + 1; row++)
            {
                temp = new List<myCell>();
                emptyColumnCellCount = 0;
                for (int column = 1; column < totalColumns + 1; column++) // the plus one is because the excel columns and rows start at 1
                {
                    tempCell = new myCell();
                    tempCell = getAllCellProperties(xlWorkSheet2.Cells[row, column], row, column);
                    temp.Add(tempCell);
                    // detecting if the cell is uneddited and empty
                    if (tempCell.text.Equals("") // no text in cell
                        && tempCell.merge == false // not a merged cell
                        && tempCell.color == 16777215 // 16777215 is white background cell color
                        && tempCell.rightLineStyle == XlLineStyle.xlLineStyleNone // default border line style
                        && tempCell.rightWeight == XlBorderWeight.xlThin) // default border weight
                                                                          // only caring about right side of border since we are reading from left to right and there might be valid cells above and beneath.
                        emptyColumnCellCount++;
                    else emptyColumnCellCount = 0;
                    if (emptyColumnCellCount >= cellBuffer)
                    {
                        if (firstRowFound)
                            if (column - cellBuffer < longestRow) continue;
                            else
                            {
                                longestRow = column - cellBuffer;
                                tempAllCells.Add(temp);
                                //for (int a = 0; a < temp.Count - 5; a++) allCells[row, a] = temp[a];
                                break;
                            }
                        else
                        {
                            firstRowFound = true;
                            longestRow = column - cellBuffer;
                            tempAllCells.Add(temp);
                            //for (int a = 0; a < temp.Count - 5; a++) allCells[row, a] = temp[a];
                            break;
                        }
                    }
                }
                Console.WriteLine("Finished reading row " + row);
                if (emptyColumnCellCount- cellBuffer == longestRow)
                {
                    emptyRowCellCount++;
                    if (emptyRowCellCount >= cellBuffer) break;
                }
                else emptyRowCellCount = 0;
                
            }
            allCells = new myCell[tempAllCells.Count - cellBuffer, longestRow];
            for (int a = 0; a < tempAllCells.Count - cellBuffer; a++)
                for (int b = 0; b < tempAllCells[a].Count - cellBuffer; b++)
                    allCells[a, b] = tempAllCells[a][b];

        }


        public myCell getAllCellProperties(Range cell, int row, int column)
        {
            myCell tempCell = new myCell();

            tempCell.rowIndex = row;
            tempCell.columnIndex = column;

            //tempCell.complexWords = GetStringProperties(cell[1,1], cell[1, 1].Text);
            //tempCell.complexWords = new List<myCell.myFont>();
            //string text = cell[1, 1].Text;
            //for (int a = 0; a < text.Length+1; a++)
            //{
            //    tempCell.complexWords.Add(new myCell.myFont());
            //    Characters temp = cell[1, 1].Characters(a,1);
            //    tempCell.complexWords[a].name = temp.Font.Name;
            //    tempCell.complexWords[a].size = temp.Font.Size;
            //    tempCell.complexWords[a].color = temp.Font.Color;
            //    tempCell.complexWords[a].bold = temp.Font.Bold;
            //    tempCell.complexWords[a].italic = temp.Font.Italic;
            //    tempCell.complexWords[a].underline = temp.Font.Underline;
            //    tempCell.complexWords[a].strikeThrough = temp.Font.Strikethrough;
            //}
            tempCell.horizontalAlignment = cell[1, 1].HorizontalAlignment;
            tempCell.verticalAlignment = cell[1, 1].VerticalAlignment;
            tempCell.text = cell[1, 1].Text;
            tempCell.height = cell[1, 1].RowHeight;
            tempCell.width = cell[1, 1].ColumnWidth;
            tempCell.color = cell[1, 1].Interior.Color;
            tempCell.merge = cell[1, 1].MergeCells();
            tempCell.name = cell[1, 1].Font.Name;
            tempCell.size = cell[1, 1].Font.Size;
            tempCell.fontColor = cell[1, 1].Font.Color;
            tempCell.bold = cell[1, 1].Font.Bold;
            tempCell.italic = cell[1, 1].Font.Italic;
            tempCell.underline = cell[1, 1].Font.Underline;
            tempCell.strikeThrough = cell[1, 1].Font.Strikethrough;



            tempCell.topLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeTop).LineStyle;
            tempCell.topWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeTop).Weight;
            tempCell.rightLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).LineStyle;
            tempCell.rightWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).Weight;
            tempCell.bottomLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).LineStyle;
            tempCell.bottomWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).Weight;
            tempCell.leftLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).LineStyle;
            tempCell.leftWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).Weight;
            //allCells[row - 1, column - 1] = tempCell; // the -1 is because in excel you start counting from 1 not 0 and we don't want an empty cell in the
                                                      // beginning of each column and row.
            return tempCell;
            

        }
    }
}
