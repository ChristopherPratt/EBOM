using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
//using Microsoft.Office.Core;

namespace EBOM_Creation_Tool_v2
{
    public class EBOMexcelFile
    {
        public EBOMexcelFile(excelFileHandler excelFileHandler2, List<List<string>> componentList, excelSection template, string fileName, mainFrame mainFrame1, int xmlPartCount, List<string> titleBlockInfo)
        {
            populateTitleBlock(excelFileHandler2.xlWorkSheet, titleBlockInfo, template.TBrowIndex, template.TBcolumnIndex); mainFrame1.writeToConsole("Finished writing title block.");
            populateHeaderRow(excelFileHandler2.xlWorkSheet, template.Htext, template.HrowIndex, template.HcolumnIndex); mainFrame1.writeToConsole("Finished writing header row.");
            int finalCount = populateBody(excelFileHandler2.xlWorkSheet, componentList, template, mainFrame1); mainFrame1.writeToConsole("Finished writing body."); 
            mainFrame1.missingParts(compareComponentCounts(xmlPartCount, finalCount));
            createFooter(excelFileHandler2.xlWorkSheet, componentList.Count + template.rowIndex[0][0], template.footerColor, template.footerList); mainFrame1.writeToConsole("Finished writing footer.");
            if (!saveExcelFile(excelFileHandler2.xlWorkSheet, fileName))
                excelFileHandler2.xlWorkBook.Close(SaveChanges: false);
            mainFrame1.writeToConsole("Saved new EBOM file.");

        }
        private void populateTitleBlock(Worksheet excel, List<string> text, List<int> row, List<int> column)
        {
            for (int a = 0; a < text.Count; a++)
            {
                writeData(excel.Cells[row[a], column[a]], text[a]);
            }
        }
        private void populateHeaderRow(Worksheet excel, List<string> text, List<int> row, List<int> column)
        {
            for (int a = 0; a < text.Count; a++)
            {
                writeData(excel.Cells[row[a], column[a]], text[a]);
            }
        }
        private int populateBody(Worksheet excel, List<List<string>> componentList, excelSection template1, mainFrame mainFrame2)
        {
            int numberOfTemplateRows = template1.backGroundColor.Count;
            int startRow = template1.rowIndex[0][0];
            int row = 0;
            int finalCount = 0;
            int match = 0;
            for (int a = 0; a < componentList.Count; a++)// for each row
            {
                if (mainFrame.end) throw new Exception("User is closing the program");
                row = a % template1.backGroundColor.Count;

                for (int b = 0; b < componentList[a].Count; b++)// for each column
                {
                    writeData(excel.Cells[startRow + a, template1.columnIndex[row][b]], componentList[a][b]); // write cell text
                }

                for (int b = 0; b < componentList[a].Count; b+=match + 1)// for each column
                {
                    match = 0;
                    for (int c = b + 1; c < componentList[a].Count; c++) // for each column of the same row after the 'b' column
                        if (template1.backGroundColor[row][b] == template1.backGroundColor[row][c]) match++;
                        else break;
                    if (template1.backGroundColor[row][b] == 16777215)
                        continue; // skip matching segment if already default setting.
                    Range range = getRange(excel, b, b + match, a + startRow, template1.columnIndex); // get like cells
                    range.Interior.Color = template1.backGroundColor[row][b];
                }

                for (int b = 0; b < componentList[a].Count; b += match + 1)// for each column
                {
                    match = 0;
                    for (int c = b + 1; c < componentList[a].Count; c++) // for each column of the same row after the 'b' column
                        if (template1.rightLineStyle[row][b] == template1.rightLineStyle[row][c]) match++;
                        else break;
                    if (template1.rightLineStyle[row][b] == XlLineStyle.xlLineStyleNone)
                        continue; // skip matching segment if already default setting.
                    Range range = getRange(excel, b, b + match, a + startRow, template1.columnIndex); // get like cells
                    range.Borders.LineStyle = template1.rightLineStyle[row][b];
                }
                for (int b = 0; b < componentList[a].Count; b += match + 1)// for each column
                {
                    match = 0;
                    for (int c = b + 1; c < componentList[a].Count; c++) // for each column of the same row after the 'b' column
                        if (template1.rightWeight[row][b] == template1.rightWeight[row][c]) match++;
                        else break;
                    if (template1.rightWeight[row][b] == XlBorderWeight.xlThin)
                        continue; // skip matching segment if already default setting.
                    Range range = getRange(excel, b, b + match, a + startRow, template1.columnIndex); // get like cells
                    range.Borders.Weight = template1.rightWeight[row][b];
                }
                for (int b = 0; b < componentList[a].Count; b += match + 1)// for each column
                {
                    match = 0;
                    for (int c = b + 1; c < componentList[a].Count; c++) // for each column of the same row after the 'b' column
                        if (template1.bottomLineStyle[row][b] == template1.bottomLineStyle[row][c]) match++;
                        else break;
                    if (template1.bottomLineStyle[row][b] == XlLineStyle.xlLineStyleNone)
                        continue; // skip matching segment if already default setting.
                    Range range = getRange(excel, b, b + match, a + startRow, template1.columnIndex); // get like cells
                    range.Borders.LineStyle = template1.bottomLineStyle[row][b];
                }
                for (int b = 0; b < componentList[a].Count; b += match + 1)// for each column
                {
                    match = 0;
                    for (int c = b + 1; c < componentList[a].Count; c++) // for each column of the same row after the 'b' column
                        if (template1.bottomWeight[row][b] == template1.bottomWeight[row][c]) match++;
                        else break;
                    if (template1.bottomWeight[row][b] == XlBorderWeight.xlThin)
                        continue; // skip matching segment if already default setting.
                    Range range = getRange(excel, b, b + match, a + startRow, template1.columnIndex); // get like cells
                    range.Borders.LineStyle = template1.bottomWeight[row][b];
                }
                for (int b = 0; b < componentList[a].Count; b += match + 1)// for each column
                {
                    match = 0;
                    for (int c = b + 1; c < componentList[a].Count; c++) // for each column of the same row after the 'b' column
                        if (template1.leftLineStyle[row][b] == template1.leftLineStyle[row][c]) match++;
                        else break;
                    if (template1.leftLineStyle[row][b] == XlLineStyle.xlLineStyleNone)
                        continue; // skip matching segment if already default setting.
                    Range range = getRange(excel, b, b + match, a + startRow, template1.columnIndex); // get like cells
                    range.Borders.LineStyle = template1.leftLineStyle[row][b];
                }
                for (int b = 0; b < componentList[a].Count; b += match + 1)// for each column
                {
                    match = 0;
                    for (int c = b + 1; c < componentList[a].Count; c++) // for each column of the same row after the 'b' column
                        if (template1.leftWeight[row][b] == template1.leftWeight[row][c]) match++;
                        else break;
                    if (template1.leftWeight[row][b] == XlBorderWeight.xlThin)
                        continue; // skip matching segment if already default setting.
                    Range range = getRange(excel, b, b + match, a + startRow, template1.columnIndex); // get like cells
                    range.Borders.LineStyle = template1.leftWeight[row][b];
                }
                for (int b = 0; b < componentList[a].Count; b += match + 1)// for each column
                {
                    match = 0;
                    for (int c = b + 1; c < componentList[a].Count; c++) // for each column of the same row after the 'b' column
                        if (template1.textFont[row][b] == template1.textFont[row][c]) match++;
                        else break;
                    if (template1.textFont[row][b] == "Calibri")
                        continue; // skip matching segment if already default setting.
                    Range range = getRange(excel, b, b + match, a + startRow, template1.columnIndex); // get like cells
                    range.Text.Name = template1.textFont[row][b];
                }
                for (int b = 0; b < componentList[a].Count; b += match + 1)// for each column
                {
                    match = 0;
                    for (int c = b + 1; c < componentList[a].Count; c++) // for each column of the same row after the 'b' column
                        if (template1.textSize[row][b] == template1.textSize[row][c]) match++;
                        else break;
                    if (template1.textSize[row][b] == 11)
                        continue; // skip matching segment if already default setting.
                    Range range = getRange(excel, b, b + match, a + startRow, template1.columnIndex); // get like cells
                    range.Text.Size = template1.textSize[row][b];
                }
                for (int b = 0; b < componentList[a].Count; b += match + 1)// for each column
                {
                    match = 0;
                    for (int c = b + 1; c < componentList[a].Count; c++) // for each column of the same row after the 'b' column
                        if (template1.textColor[row][b] == template1.textColor[row][c]) match++;
                        else break;
                    if (template1.textColor[row][b] == 0)
                        continue; // skip matching segment if already default setting.
                    Range range = getRange(excel, b, b + match, a + startRow, template1.columnIndex); // get like cells
                    range.Text.Color = template1.textColor[row][b];
                }
                

                mainFrame2.writeToConsole("Row " + a + " finished.");
                finalCount++;
            }
            return finalCount;
        }


        private Range getRange(Worksheet excel1, int startIndex, int endIndex, int rowIndex, List<List<int>> columnIndex1)
        {
            Range startCell = excel1.Cells[rowIndex, columnIndex1[0][startIndex]];
            Range endCell = excel1.Cells[rowIndex, columnIndex1[0][endIndex]];
            return excel1.Range[startCell, endCell]; // get whole row
        }
        private void createFooter(Worksheet excel, int rowCount, int footerColor, List<string> footerList)
        {
            float height = (float)excel.Range[excel.Cells[1, 1], excel.Cells[rowCount, 1]].Height;
            int textBoxWidthModifier = 720;
            float textBoxHeightModifier = 34;

            Microsoft.Office.Interop.Excel.Shape textBox = excel.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 30, height, 60, 30);
            textBox.Fill.ForeColor.RGB = footerColor;

            string footerVal = "";
            for (int i = 1; i <= footerList.Count; i++)
                footerVal += "Note " + i + '-' + footerList[i - 1] + '\n' + '\n';

            textBox.TextFrame.Characters().Text = footerVal;
            textBox.Height = footerList.Count * textBoxHeightModifier;
            textBox.Width = textBoxWidthModifier;
        }
        private bool saveExcelFile(Worksheet excel, string fileName)
        {
            bool excelFileIsOpen = false;
            do
            {
                try
                {
                    excel.SaveAs(fileName);
                }
                catch (Exception e)
                {
                    DialogResult dialogResult = MessageBox.Show("Excel file with same name already open.\n Would you like to exit without creating a new EBOM?", "Warning", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes) // exit without saving excel file
                    {
                        return false;
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
            return false;
        }
        public void writeData(Range cell, string text)
        {
            cell[1,1].Value = text;
        }

        private bool compareComponentCounts(int xmlCount, int ExcelCount)
        {
            if (xmlCount != ExcelCount)
            {
                MessageBox.Show("Warning, XML part count different from Excel part count.\n XML part count: " + xmlCount + "\nExcel part count: " + ExcelCount);
                return false;
            }
            return true;
        }
    }
}
