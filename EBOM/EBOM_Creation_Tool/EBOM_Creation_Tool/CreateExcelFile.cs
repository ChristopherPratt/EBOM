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
        MainFrame mainframe;

        string ExportFileName;

        public CreateExcelFile(MainFrame m, LoadXML x, LoadTemplate t)
        {

            mainframe = m;
            XML = x;
            template = t;
            if (mainframe.end) return;
            try
            {
                setupExcel();
                cleanUpTemplateArtifacts();
                writeTitleBlock();
                colorAllRows();
                writeBodyInfo();
                writeFooterInfo();
                saveExcelFile();
                ////Start the Microsoft Excel Application
                ////string excelTemplate = System.AppDomain.CurrentDomain.BaseDirectory + "template.xlsx";
                //ExportFileName = System.IO.Path.ChangeExtension(XML.xmlFile, null) + ".xlsx";


                //template.xlApp = template.template.xlAppOpen;
                //if (template.xlApp == null)
                //{
                //    MessageBox.Show("Excel is not properly installed!!");
                //    return;
                //}
                //template.template.xlWorkBooks = template.template.template.xlWorkBooks2;
                //try
                //{
                //    template.xlWorkBook = template.template.xlWorkBook2;
                //}
                //catch (Exception e) { return; }
                //template.xlWorkSheet = template.template.xlWorkSheet2;

                ////template.xlWorkSheet.Cells[1, template.columnEnd+1] = "";
                //template.xlWorkSheet.Cells[1, template.columnEnd] = "";
                //template.xlWorkSheet.Cells[template.rowEnd, 1] = "";
                //foreach (LoadTemplate.myCell cell in template.titleBlock) writeData(cell);
                //foreach (LoadTemplate.myCell cell in template.headerRow) writeData(cell);
                //for (int a = template.bodyRowStart; a < template.rowEnd; a++)
                //{
                //    for (int b = 1; b < template.columnEnd; b++)
                //    {
                //        template.xlWorkSheet.Cells[a, b] = ""; // clear all artifacts from body and footer section.
                //        template.xlWorkSheet.Cells[a, b].Interior.Color = 16777215;
                //    }
                //}
                //mainframe.WriteToConsole("Finished writing Title Block");

                //for (int a = 0; a < XML.sorted.Count; a++)
                //{
                //    Range range = template.xlWorkSheet.Range[template.xlWorkSheet.Cells[XML.sorted[a][0].rowIndex, XML.sorted[a][0].columnIndex], template.xlWorkSheet.Cells[XML.sorted[a][XML.sorted[a].Count - 1].rowIndex, XML.sorted[a][XML.sorted[a].Count-1].columnIndex]]; // get whole row
                //    range.Interior.Color = template.bodyColors[a % template.bodyColors.Count]; // change color of whole row
                //    if (XML.sorted[a][XML.sorted[a].Count - 1].rightLineStyle != XlLineStyle.xlLineStyleNone) // check if we need to set borders
                //    {
                //        for (int b = 0; b < template.bodyRows[a].Count; b++)//set borders
                //        {
                //            template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeRight).LineStyle = XML.sorted[a][b].rightLineStyle;
                //            template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeRight).Weight = XML.sorted[a][b].rightWeight;
                //            template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XML.sorted[a][b].bottomLineStyle;
                //            template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeBottom).Weight = XML.sorted[a][b].bottomWeight;
                //            template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XML.sorted[a][b].leftLineStyle;
                //            template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeLeft).Weight = XML.sorted[a][b].leftWeight;
                //        }
                //    }
                //}


                //int c = 0;
                //foreach (List<LoadTemplate.myCell> cellRows in XML.sorted)
                //{        
                //    foreach(LoadTemplate.myCell cell in cellRows) writeData(cell);
                //    mainframe.WriteToConsole("Finished writing row: " + ++c);
                //}
                //mainframe.WriteToConsole("Total Part Count: " + XML.totalPartCount);
                //if (c != XML.totalPartCount) MessageBox.Show("WARNING!\nNot all parts exported to BOM from xml.");
                ////template.xlWorkBook.SaveAs(System.AppDomain.CurrentDomain.BaseDirectory + "New EBOM1 " + template.time + ".xlsx");
                //int footerOffset = 2;
                //float height = (float)template.xlWorkSheet.Range[template.xlWorkSheet.Cells[1, 1], template.xlWorkSheet.Cells[footerOffset + template.bodyRowStart + XML.sorted.Count, 1]].Height;
                //int textBoxWidthModifier = 512;
                //float textBoxHeightModifier = 48;

                //Microsoft.Office.Interop.Excel.Shape textBox = template.xlWorkSheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 30, height, 60, 30);
                //textBox.Fill.ForeColor.RGB = System.Drawing.Color.DodgerBlue.ToArgb();

                //string footerVal = "";
                //for (int i = 1; i <= template.footerList.Count; i++)
                //    footerVal += "Note " + i + '-' + template.footerList[i - 1] + '\n' + '\n';

                //textBox.TextFrame.Characters().Text = footerVal;
                //textBox.Height = template.footerList.Count * textBoxHeightModifier;
                //textBox.Width = textBoxWidthModifier;
                //bool excelFileIsOpen = false;
                //do
                //{
                //    try
                //    {
                //        template.xlWorkBook.SaveAs(ExportFileName);
                //    }
                //    catch (Exception e)
                //    {
                //        MessageBox.Show("Please close the Excel file with the same name as the .XML file that was chosen.");
                //        excelFileIsOpen = true;
                //        continue;
                //    }
                //    excelFileIsOpen = false;
                //}
                //while (excelFileIsOpen);


            }
            finally
            {
                //object misValue = System.Reflection.Missing.Value;
                mainframe.WriteToConsole("Finished Saving Excel File");
                Marshal.FinalReleaseComObject(template.xlWorkSheet);
                //template.xlWorkBook.Close(false, misValue, misValue); 
                template.xlWorkBook.Close(SaveChanges: false);
                //template.xlWorkBook.Close();
                Marshal.FinalReleaseComObject(template.xlWorkBook);
                template.xlWorkBooks.Close();
                Marshal.FinalReleaseComObject(template.xlWorkBooks);
                template.xlApp.Quit();
                Marshal.FinalReleaseComObject(template.xlApp); // excel objects don't releast comObjects to excel so you have to force it
            }
        }

        private void setupExcel()
        {
            ExportFileName = System.IO.Path.ChangeExtension(XML.xmlFile, null) + ".xlsx";
        }
        private void cleanUpTemplateArtifacts()
        {
            //template.xlWorkSheet.Cells[1, template.columnEnd+1] = "";
            template.xlWorkSheet.Cells[1, template.columnEnd] = "";
            template.xlWorkSheet.Cells[template.rowEnd, 1] = "";

            for (int a = template.bodyRowStart; a < template.rowEnd; a++)
            {
                for (int b = 1; b < template.columnEnd; b++)
                {
                    template.xlWorkSheet.Cells[a, b] = ""; // clear all artifacts from body and footer section.
                    template.xlWorkSheet.Cells[a, b].Interior.Color = 16777215;
                }
            }
        }

        private void writeTitleBlock()
        {
            foreach (LoadTemplate.myCell cell in template.titleBlock) writeData(cell);
            foreach (LoadTemplate.myCell cell in template.headerRow) writeData(cell);
            mainframe.WriteToConsole("Finished writing Title Block");
        }

        private void colorAllRows()
        {
            for (int a = 0; a < XML.sorted.Count; a++)
            {
                Range range = template.xlWorkSheet.Range[template.xlWorkSheet.Cells[XML.sorted[a][0].rowIndex, XML.sorted[a][0].columnIndex], template.xlWorkSheet.Cells[XML.sorted[a][XML.sorted[a].Count - 1].rowIndex, XML.sorted[a][XML.sorted[a].Count - 1].columnIndex]]; // get whole row
                range.Interior.Color = template.bodyColors[a % template.bodyColors.Count]; // change color of whole row
                if (XML.sorted[a][XML.sorted[a].Count - 1].rightLineStyle != XlLineStyle.xlLineStyleNone) // check if we need to set borders
                {
                    for (int b = 0; b < template.bodyRows[a].Count; b++)//set borders
                    {
                        template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeRight).LineStyle = XML.sorted[a][b].rightLineStyle;
                        template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeRight).Weight = XML.sorted[a][b].rightWeight;
                        template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XML.sorted[a][b].bottomLineStyle;
                        template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeBottom).Weight = XML.sorted[a][b].bottomWeight;
                        template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XML.sorted[a][b].leftLineStyle;
                        template.xlWorkSheet.Cells[XML.sorted[a][b].rowIndex, XML.sorted[a][b].columnIndex].Borders(XlBordersIndex.xlEdgeLeft).Weight = XML.sorted[a][b].leftWeight;
                    }
                }
            }
            mainframe.WriteToConsole("Finished coloring rows");
        }

        private void writeBodyInfo()
        {
            int c = 0;
            foreach (List<LoadTemplate.myCell> cellRows in XML.sorted)
            {
                foreach (LoadTemplate.myCell cell in cellRows) writeData(cell);
                mainframe.WriteToConsole("Finished writing row: " + ++c);
            }
            mainframe.WriteToConsole("Total Part Count: " + XML.totalPartCount);
            if (c != XML.totalPartCount) MessageBox.Show("WARNING!\nNot all parts exported to BOM from xml.");
            template.xlWorkSheet.Columns.AutoFit(); // autofit all columns in the sheet.
            for (int a = 1; a < template.columnEnd; a++)
            {
                template.xlWorkSheet.Cells[1, a].ColumnWidth = template.xlWorkSheet.Cells[1, a].ColumnWidth + 3;
            }
        }

        private void writeFooterInfo()
        {
            int footerOffset = 2;
            float height = (float)template.xlWorkSheet.Range[template.xlWorkSheet.Cells[1, 1], template.xlWorkSheet.Cells[footerOffset + template.bodyRowStart + XML.sorted.Count, 1]].Height;
            int textBoxWidthModifier = 720;
            float textBoxHeightModifier = 34;

            Microsoft.Office.Interop.Excel.Shape textBox = template.xlWorkSheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 30, height, 60, 30);
            //textBox.Fill.ForeColor.RGB = System.Drawing.Color.DodgerBlue.ToArgb();
            textBox.Fill.ForeColor.RGB = template.footerColor;

            string footerVal = "";
            for (int i = 1; i <= template.footerList.Count; i++)
                footerVal += "Note " + i + '-' + template.footerList[i - 1] + '\n' + '\n';

            textBox.TextFrame.Characters().Text = footerVal;
            textBox.Height = template.footerList.Count * textBoxHeightModifier;
            textBox.Width = textBoxWidthModifier;
        }

        public void writeData(LoadTemplate.myCell cell)
        {

            template.xlWorkSheet.Cells[cell.rowIndex, cell.columnIndex].Value = cell.info;
        }

        private void saveExcelFile()
        {
            bool excelFileIsOpen = false;
            do
            {
                try
                {
                    template.xlWorkBook.SaveAs(ExportFileName);
                }
                catch (Exception e)
                {
                    DialogResult dialogResult = MessageBox.Show("Excel file with same name already open.\n Would you like to exit without creating a new EBOM?", "Warning", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes) // exit without saving excel file
                    {
                        excelFileIsOpen = false;
                        mainframe.end = true;
                    }
                    else if (dialogResult == DialogResult.No) // retry
                    {
                        excelFileIsOpen = true;
                        continue;
                    }
                    
                }
                excelFileIsOpen = false;
            }
            while (excelFileIsOpen);
        }

        private string getTime()
        {
            string datePatt = @"hh.mm.ss.ff";
            DateTime saveUtcNow = DateTime.UtcNow;
            return saveUtcNow.ToString(datePatt);
        }



    }
}
