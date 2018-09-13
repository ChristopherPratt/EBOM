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
        public EBOMexcelFile(excelFileHandler excelFileHandler2, List<List<string>> componentList, excelSection template, string fileName, mainFrame mainFrame1, int xmlPartCount)
        {
            populateTitleBlock(excelFileHandler2.xlWorkSheet, template.TBtext, template.TBrowIndex, template.TBcolumnIndex); mainFrame1.writeToConsole("Finished writing title block.");
            populateHeaderRow(excelFileHandler2.xlWorkSheet, template.Htext, template.HrowIndex, template.HcolumnIndex); mainFrame1.writeToConsole("Finished writing header row.");
            int finalCount = populateBody(excelFileHandler2.xlWorkSheet, componentList, template, mainFrame1); mainFrame1.writeToConsole("Finished writing body.");
            compareComponentCounts(xmlPartCount, finalCount);
            createFooter(excelFileHandler2.xlWorkSheet, componentList.Count + template.rowIndex[0][0], template.footerColor, template.footerList); mainFrame1.writeToConsole("Finished writing footer.");
            saveExcelFile(excelFileHandler2.xlWorkSheet, fileName); mainFrame1.writeToConsole("Saved new EBOM file.");
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
            for (int a = 0; a < componentList.Count; a++)// for each row
            {
                for (int b = 0; b < template1.rowIndex.Count; b++)// for each column
                {

                    row = a % template1.backGroundColor.Count;
                    writeData(excel.Cells[startRow + a, template1.columnIndex[row][b]], componentList[a][b]);
                    modifyCellStyle(excel.Cells[startRow + a, template1.columnIndex[row][b]], template1.rightLineStyle[row][b], template1.rightWeight[row][b],
                                                                                                template1.bottomLineStyle[row][b], template1.bottomWeight[row][b],
                                                                                                template1.leftLineStyle[row][b], template1.leftWeight[row][b],
                                                                                                template1.textColor[row][b], template1.backGroundColor[row][b],
                                                                                                template1.textFont[row][b], template1.textSize[row][b]);
                }
                mainFrame2.writeToConsole("Row " + a + " finished.");
                finalCount++;
            }
            return finalCount;
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
        private void saveExcelFile(Worksheet excel, string fileName)
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
                        throw e;
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
        public void writeData(Range cell, string text)
        {
            cell[1,1].Value = text;
            //modifyCellStyle(null, (XlBorderWeight)null, null, null, null, null, null, null, null, null, null);
        }
        public void modifyCellStyle(Range cell, XlLineStyle rightStyle, XlBorderWeight rightWeight,
                                                    XlLineStyle bottomStyle, XlBorderWeight bottomWeight,
                                                    XlLineStyle leftStyle, XlBorderWeight leftWeight,
                                                    double textColor, double backGroundColor, string textFont, int textSize)
        {
            if (rightStyle != null && rightStyle != XlLineStyle.xlLineStyleNone) {cell[1,1].Borders(XlBordersIndex.xlEdgeRight).LineStyle = rightStyle; }
            if (rightWeight != null && rightWeight != XlBorderWeight.xlHairline) { cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).Weight = rightWeight; }
            if (bottomStyle != null && bottomStyle != XlLineStyle.xlLineStyleNone) { cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).LineStyle = bottomStyle; }
            if (bottomWeight != null && bottomWeight != XlBorderWeight.xlHairline) { cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).Weight = bottomWeight; }
            if (leftStyle != null && leftStyle != XlLineStyle.xlLineStyleNone) { cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).LineStyle = leftStyle; }
            if (leftWeight != null && leftWeight != XlBorderWeight.xlHairline) { cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).Weight = leftWeight; }
            if (textColor != null && textColor != 1) { cell[1, 1].Font.Color = textColor; }
            if (backGroundColor != null && backGroundColor != 16777215) { cell[1, 1].Interior.Color = backGroundColor; }
            if (textFont != null && textFont != "") { cell[1, 1].Font.Name = textFont; }
            if (textSize != null && textSize != 10) { cell[1, 1].Font.Size = textSize; }
        }
        private void compareComponentCounts(int xmlCount, int ExcelCount)
        {
            if (xmlCount != ExcelCount)
            {
                MessageBox.Show("Warning, XML part count different from Excel part count.\n XML part count: " + xmlCount + "\nExcel part count: " + ExcelCount);
            }
        }
    }
}
