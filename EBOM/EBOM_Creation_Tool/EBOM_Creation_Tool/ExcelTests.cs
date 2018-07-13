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
    class ExcelTests
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


        public static int AlphabetStart = 65;

        public static int numberOfRows = 25;
        public static int numberOfColumns = 26;
        public static int DefaultCellWidth = 0;
        public static int DefaultCellHeight = 0;
        public static int rowLabelLength = 30;


        public static List<String> ElementList = new List<String>();

        int HeaderElementCounter = 0;

        Boolean ElementListBoxMouseDown = false;
        Boolean ElementListBoxDisableFlag = false;
        string time;

        XmlNodeList nodeList;

        public ExcelTests()
        {

            time = getTime();
            Thread gui = new Thread(delegate ()
            {

                //try
                //{
                //    readExcelFile();
                //    saveExcelFile();                    
                //}
                //finally
                //{
                //    GC.Collect();
                //    GC.WaitForPendingFinalizers();
                //}
                try
                {
                    readExcelFile(2, 4);
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                try
                {
                    saveExcelFile();
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            });
            gui.Name = "gui";
            gui.Start();



        }
        Microsoft.Office.Interop.Excel.Application xlApp;
        Workbooks xlWorkBooks;
        Workbook xlWorkBook;
        Sheets xlWorkSheets;
        Worksheet xlWorkSheet;
        public void saveExcelFile()
        {

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

                /*
                 * 
             //Get Cell Default Width and Height
             Graphics g = CreateGraphics();
             System.Console.WriteLine("Column Width = " + xlWorkSheet.Columns.ColumnWidth);
             System.Console.WriteLine("Row Height = " + xlWorkSheet.Rows.RowHeight);
             DefaultCellWidth = Convert.ToInt32(xlWorkSheet.Columns.ColumnWidth * 10);
             //DefaultCellHeight = Convert.ToInt32(xlWorkSheet.Rows.RowHeight / (72 / g.DpiY));
             DefaultCellHeight = Convert.ToInt32(xlWorkSheet.Columns.ColumnWidth * 10);
             System.Console.WriteLine("Default Cell Width = " + DefaultCellWidth);
             System.Console.WriteLine("Default Cell Height = " + DefaultCellHeight);


             //Populate Element List
             //For right now this will be hardcoded
             //-----------------------------------------------------------------------
             ElementList.Add("AdditionalNote");
             ElementList.Add("ApprovedBy");
             ElementList.Add("Designator");
             ElementList.Add("Edited by");
             ElementList.Add("mfg");
             ElementList.Add("ModifiedDate");
             ElementList.Add("note");
             ElementList.Add("Package");
             ElementList.Add("Part Number");
             ElementList.Add("Quantity");
             ElementList.Add("Release No");
             ElementList.Add("Title");
             ElementList.Add("Value");
             ElementList.Add("Description");
             ElementList.Add("Cust. No");
             ElementList.Add("Product Number");
             ElementList.Add("change index");
             ElementList.Add("Product No.");

             for (int i = 0; i < ElementList.Count; i++)
             {
                 ElementListBox.Items.Add(ElementList.ElementAt(i));
             }
             ////-----------------------------------------------------------------------
             ElementListBox.SetSelected(0, true);


             //Test Excel Code
             //-----------------------------------------------------------------------------------

             */
                //xlWorkSheet.Cells[1, 1] = "Jeremy";
                //xlWorkSheet.Cells[1, 1].Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlLineStyleNone;
                //xlWorkSheet.Cells[1, 1].Borders(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlThick;

                //xlWorkSheet.Cells[1, 2] = "Galazin";
                //xlWorkSheet.Cells[2, 1] = "Testing";
                //xlWorkSheet.Cells[2, 2] = "123";
                int count = 0;
                int currentRow = 1;
                int firstRow = 1;
                bool oneTime = true;


                int totalRows = allCells.GetLength(0);
                int totalColumns = allCells.GetLength(1);


                ////////////////////// outside border area //////////////////////////
                for (int column = 0; column < totalColumns; column++)
                {
                    writeData(allCells[0, column]);
                    xlWorkSheet.Rows[allCells[0, column].rowIndex].RowHeight = allCells[0, column].height; // sets cell height for just the 1st row
                    xlWorkSheet.Columns[column + 1].ColumnWidth = allCells[0, column].width; // sets cell width for every column
                }
                //WriteToConsole("Finished writing top border");

                for (int row = 1; row < totalRows; row++)
                {
                    writeData(allCells[row, totalColumns - 1]);
                    xlWorkSheet.Rows[row].RowHeight = allCells[row, totalColumns - 1].height; // sets the cell height for every column except row 1
                }
                Console.WriteLine("Finished writing right border");

                for (int column = 0; column < totalColumns - 1; column++)
                    writeData(allCells[totalRows - 1, column]);
                Console.WriteLine("Finished writing bottom border");

                for (int row = 1; row < totalRows - 1; row++)
                    writeData(allCells[row, 0]);
                Console.WriteLine("Finished writing left border");
                ////////////////////// outside border area //////////////////////////

                ////////////////////// checkered inside border area //////////////////////////
                for (int row = 1; row < totalRows - 1; row++)
                {
                    for (int column = 1; column < totalColumns - 1; column++) // the plus one is because the excel columns and rows start at 1
                    {
                        writeCheckeredData(allCells[row, column]);
                    }
                    Console.WriteLine("Finished writing row " + row);
                }

                ////////////////////// checkered inside border area //////////////////////////


                //foreach (myCell cell in allCells)
                //{

                //    // ideas for making this faster
                //    /*
                //     * alternate populating borders in checker pattern since neighboring cells can detect borders of other cells.
                //     * just make sure you get all the cells on the outside of the rectangle of all cells first and fill in the middle 
                //     * in a checker pattern
                //     * 
                //     * put the merge check and row and column widths and heights in a separate loop, 
                //     * 
                //     * don't set the color if the color is already white.
                //     * */
                //    if (oneTime)
                //    {
                //        currentRow = cell.rowIndex;
                //        firstRow = cell.rowIndex;
                //        xlWorkSheet.Rows[cell.rowIndex].RowHeight = cell.height; // first row needs height adjusted
                //        oneTime = false;
                //    }

                //    if (firstRow == cell.rowIndex) xlWorkSheet.Columns[cell.columnIndex].ColumnWidth = cell.width;
                //    if (currentRow != cell.rowIndex)
                //    {
                //        currentRow = cell.rowIndex;
                //        xlWorkSheet.Rows[cell.rowIndex].RowHeight = cell.height;
                //    }

                //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Interior.Color = cell.color;
                //    //16777215 color of white cell
                //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].HorizontalAlignment = cell.horizontalAlignment;
                //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].VerticalAlignment = cell.verticalAlignment;

                //    if (cell.topLineStyle == XlLineStyle.xlLineStyleNone) cell.topWeight = XlBorderWeight.xlHairline;
                //    if (cell.rightLineStyle == XlLineStyle.xlLineStyleNone) cell.rightWeight = XlBorderWeight.xlHairline;
                //    if (cell.bottomLineStyle == XlLineStyle.xlLineStyleNone) cell.bottomWeight = XlBorderWeight.xlHairline;
                //    if (cell.leftLineStyle == XlLineStyle.xlLineStyleNone) cell.leftWeight = XlBorderWeight.xlHairline;

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


                //    //for (int a = cell.complexWords.Count-1 ; a > 0 ; a--)
                //    //{
                //    //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Characters(a, 1).Font.Name = cell.complexWords[a].name;
                //    //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Characters(a, 1).Font.Size = cell.complexWords[a].size;
                //    //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Characters(a, 1).Font.Color = cell.complexWords[a].color;
                //    //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Characters(a, 1).Font.Bold = cell.complexWords[a].bold;
                //    //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Characters(a, 1).Font.Italic = cell.complexWords[a].italic;
                //    //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Characters(a, 1).Font.Underline = cell.complexWords[a].underline;
                //    //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Characters(a, 1).Font.Strikethrough = cell.complexWords[a].strikeThrough;
                //    //}
                //    //for (int a = 0; a < cell.complexWords.Count; a++)
                //    //{
                //    //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex] += cell.complexWords[a].Text;
                //    //    xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Characters(a, a).Text = cell.complexWords[a].Font.Color;
                //    //}
                //    count++;
                //    if (count % 100 == 0) Console.WriteLine("Cell " + count++ + " Finished");

                //}
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




            //xlWorkSheet = null;
            //xlWorkBook = null;
            //xlWorkBooks = null;
            //xlApp = null;


            //if (xlWorkSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            //if (xlWorkBook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            //if (xlApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp); // excel objects don't releast comObjects to excel so you have to force it

            //-----------------------------------------------------------------------------------

        }

        public void writeData(myCell cell)
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

        public void writeCheckeredData(myCell cell)
        {

            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Interior.Color = cell.color;
            //16777215 color of white cell
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].HorizontalAlignment = cell.horizontalAlignment;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].VerticalAlignment = cell.verticalAlignment;




            if (cell.columnIndex % 2 == 0) // only read borders that are even column numbers to make checker pattern
            {
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
            }


            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex] = cell.text;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Name = cell.name;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Size = cell.size;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Color = cell.fontColor;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Bold = cell.bold;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Italic = cell.italic;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Underline = cell.underline;
            xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Font.Strikethrough = cell.strikeThrough;
        }

        Microsoft.Office.Interop.Excel.Application xlAppOpen;
        Workbooks xlWorkBooks2;
        Workbook xlWorkBook2;
        Worksheet xlWorkSheet2;
        public void readExcelFile(int totalRows, int totalColumns)
        {
            allCells = new myCell[totalRows, totalColumns];

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

                //for (int row = 1; row < totalRows + 1; row++)
                //{
                //    for (int column = 1; column < totalColumns + 1; column++) // the plus one is because the excel columns and rows start at 1
                //    {
                //        getAllCellProperties(xlWorkSheet2.Cells[row, column], row, column);
                //    }
                //    Console.WriteLine("Finished reading row " + row);
                //}

                ////////////////////// outside border area //////////////////////////
                for (int column = 1; column < totalColumns + 1; column++) // the plus one is because the excel columns and rows start at 1
                    getAllCellProperties(xlWorkSheet2.Cells[1, column], 1, column);
                Console.WriteLine("Finished reading top border");

                for (int row = 2; row < totalRows + 1; row++) // the plus one is because the excel columns and rows start at 1
                    getAllCellProperties(xlWorkSheet2.Cells[row, totalColumns + 1], row, totalColumns + 1);
                Console.WriteLine("Finished reading right border");

                for (int column = 1; column < totalColumns; column++) // the plus one is because the excel columns and rows start at 1
                    getAllCellProperties(xlWorkSheet2.Cells[totalRows + 1, column], totalRows + 1, column);
                Console.WriteLine("Finished reading bottom border");

                for (int row = 2; row < totalRows; row++) // the plus one is because the excel columns and rows start at 1
                    getAllCellProperties(xlWorkSheet2.Cells[row, 1], row, 1);
                Console.WriteLine("Finished reading left border");
                ////////////////////// outside border area //////////////////////////

                ////////////////////// checkered inside border area //////////////////////////
                for (int row = 2; row < totalRows; row++)
                {
                    for (int column = 2; column < totalColumns; column++) // the plus one is because the excel columns and rows start at 1
                    {
                        getAllCellProperties(xlWorkSheet2.Cells[row, column], row, column);
                    }
                    Console.WriteLine("Finished reading row " + row);
                }
                ////////////////////// checkered inside border area //////////////////////////



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

        public void getAllCellProperties(Range cell, int row, int column)
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
            allCells[row - 1, column - 1] = tempCell; // the -1 is because in excel you start counting from 1 not 0 and we don't want an empty cell in the
                                                      // beginning of each column and row.
        }

        public void getAlternatingBortderCellProperties(Range cell, int row, int column)
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


            if (column % 2 == 0) // only read borders that are even column numbers to make checker pattern
            {
                tempCell.topLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeTop).LineStyle;
                tempCell.topWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeTop).Weight;
                tempCell.rightLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).LineStyle;
                tempCell.rightWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).Weight;
                tempCell.bottomLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).LineStyle;
                tempCell.bottomWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).Weight;
                tempCell.leftLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).LineStyle;
                tempCell.leftWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).Weight;
            }

            allCells[row - 1, column - 1] = tempCell; // the -1 is because in excel you start counting from 1 not 0 and we don't want an empty cell in the
                                                      // beginning of each column and row.
        }
        //public List<myCell.myFont> GetStringProperties(Range cell, string text)
        //{
        //    List<myCell.myFont> tempWords = new List<myCell.myFont>();
        //    for (int a = 0; a < text.Length-2; a++)
        //    {
        //        tempWords.Add(new myCell.myFont());
        //        Characters temp = cell.Characters(a,a+1);
        //        tempWords[a].name = temp.Font.Name;
        //        tempWords[a].size = temp.Font.Size;
        //        tempWords[a].color = temp.Font.Color;
        //        tempWords[a].bold = temp.Font.Bold;
        //        tempWords[a].italic = temp.Font.Italic;
        //        tempWords[a].underline = temp.Font.Underline;
        //        tempWords[a].strikeThrough = temp.Font.Strikethrough;
        //    }
        //    return tempWords;
        //}
        private string getTime()
        {
            string datePatt = @"hh.mm.ss.ff";
            DateTime saveUtcNow = DateTime.UtcNow;
            return saveUtcNow.ToString(datePatt);
        }


    }
}

