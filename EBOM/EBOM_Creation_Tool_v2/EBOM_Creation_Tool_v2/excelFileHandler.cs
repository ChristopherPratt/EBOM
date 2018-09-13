using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;



namespace EBOM_Creation_Tool_v2
{
    public class excelFileHandler
    {
        mainFrame mainFrame1;
        public Microsoft.Office.Interop.Excel.Application xlApp;
        public Workbooks xlWorkBooks;
        public Workbook xlWorkBook;
        public Worksheet xlWorkSheet;
        public int totalRows = 1000;
        public int totalColumns = 1000;



        public excelFileHandler(mainFrame m, ref excelSection template1, ref sort sort1, ref countParts countparts1)
        {
            mainFrame1 = m;
            openTemplateFile(ref xlApp, ref xlWorkBooks, ref xlWorkBook, ref xlWorkSheet);
            readTemplateFile(ref xlWorkSheet, totalRows, totalColumns, ref template1, ref sort1, ref countparts1, mainFrame1);
        }
        private void openTemplateFile(ref Microsoft.Office.Interop.Excel.Application xlApp1, ref Workbooks xlWorkBooks1, ref Workbook xlWorkBook1, ref Worksheet xlWorkSheet1)
        {
            string excelTemplate = System.AppDomain.CurrentDomain.BaseDirectory + "template.xlsx";

            if (!File.Exists(excelTemplate))
            {
                mainFrame.end = true;
                throw new Exception("Excel template not found in " + excelTemplate);                
            }
            xlApp1 = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBooks1 = xlApp.Workbooks;
            try
            {
                xlWorkBook1 = xlWorkBooks1.Open(excelTemplate, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false); //open the template file!
            }
            catch
            {
                MessageBox.Show("template file is not spelled correctly or is not in the same directory as EBOMCreationTool.exe");
                mainFrame.end = true;
                return;
            }
            xlWorkSheet1 = (Worksheet)xlWorkBook1.Worksheets.get_Item(1); //worksheet to write data to
        }
        
        private void readTemplateFile(ref Worksheet xlWorkSheet1, int totalRows1, int totalColumns1, ref excelSection template2, ref sort sort2, ref countParts countParts2, mainFrame mainFrame2)
        {
            excelFileParser excelFileParser1 = new excelFileParser();
            excelFileParser bodyRows = new excelFileParser();

            
            for (int row = 1; row < totalRows1; row++)
            {
                for (int column = 1; column < totalColumns1; column++)
                {
                    excelFileParser1.parseCell(xlWorkSheet.Cells[row, column], row, column, ref totalRows1, ref totalColumns1, ref template2, ref sort2, ref countParts2);
                }
                mainFrame2.writeToConsole("Row " + row + " finished.");
            }
            mainFrame2.writeToConsole("Finished reading template file.");
        }
    }
}
