﻿using System;
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
    class LoadTemplate
    {
        public int bodyRowStart { get; set; }
        public int quantity = 1000;
        public List<int> group;
        public class myCell
        {
            public int rowIndex { get; set; }
            public int columnIndex { get; set; }
            public string info { get; set; }
            public string text { get; set; }
            public int index  { get; set; }
            public double color { get; set; }

            public XlLineStyle topLineStyle { get; set; }
            public XlBorderWeight topWeight { get; set; }
            public XlLineStyle rightLineStyle { get; set; }
            public XlBorderWeight rightWeight { get; set; }
            public XlLineStyle bottomLineStyle { get; set; }
            public XlBorderWeight bottomWeight { get; set; }
            public XlLineStyle leftLineStyle { get; set; }
            public XlBorderWeight leftWeight { get; set; }

            public bool moreThanText { get; set; }
        }

        public myCell[,] allCells;
        public List<myCell> titleBlock;
        public List<myCell> headerRow;
        public List<List<myCell>> bodyRows;
        public List<double> bodyColors;
        public List<List<string>> sortOrder;
        public int columnEnd = 1000;
        public int rowEnd = 1000;

        public List<List<int>> mergedArea;
        public string time;
        public Microsoft.Office.Interop.Excel.Application xlAppOpen;
        public Workbooks xlWorkBooks2;
        public Workbook xlWorkBook2;
        public Worksheet xlWorkSheet2;
        public string templateFileName;
        public LoadTemplate()
        {
            time = getTime();
            titleBlock = new List<myCell>();
            headerRow = new List<myCell>();
            bodyRows = new List<List<myCell>>();
            bodyColors = new List<double>();
            group = new List<int>();
            try
            {
                //copyExcelFile();
                readExcelFile(1000, 1000); // choosing unthinkably huge number since i want to be able to cover any size template
            }                               // breaks in the loops that use those numbers prevent inefficiency.
            finally
            {
                //GC.Collect();
                //GC.WaitForPendingFinalizers();
            }
        }
        private string getTime()
        {
            string datePatt = @"hh.mm.ss.ff";
            DateTime saveUtcNow = DateTime.UtcNow;
            return saveUtcNow.ToString(datePatt);
        }

        public void readExcelFile(int totalRows, int totalColumns)
        {

            try
            {
                string excelTemplate = System.AppDomain.CurrentDomain.BaseDirectory + "template.xlsx";

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
                catch (Exception e)
                { return; }
                xlWorkSheet2 = (Worksheet)xlWorkBook2.Worksheets.get_Item(1); //worksheet to write data to

                getCells(totalRows, totalColumns);


            }
            finally
            {
                Console.WriteLine("Finished reading Excel File");
            }
        }
        public void getCells(int totalRows, int totalColumns)
        {
            List<List<myCell>> tempAllCells = new List<List<myCell>>();
            sortOrder = new List<List<string>>();
            mergedArea = new List<List<int>>();
            List<myCell> temp;
            for (int row = 1; row < totalRows + 1; row++)
            {
                if (row >= rowEnd) break;
                temp = new List<myCell>();
                for (int column = 1; column < totalColumns + 1; column++) // the plus one is because the excel columns and rows start at 1
                {
                    if (column >= columnEnd) break;
                    getAllCellProperties(xlWorkSheet2.Cells[row, column], row, column);
                }
                Console.WriteLine("Finished reading row " + row);
            }
            bodyRowStart = bodyRows[0][0].rowIndex;
        }
        public bool getallPropertiesOfRow = false;
        public int currentRow = 0;
        public myCell getAllCellProperties(Range cell, int row, int column)
        {
            myCell tempCell = new myCell();            
            bool matched = false;
            tempCell.rowIndex = row;
            tempCell.columnIndex = column;           
            tempCell.text = cell[1, 1].Text;
            tempCell.color = cell[1, 1].Interior.Color;
            tempCell.rightLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).LineStyle; // ignoring top border because it would overwrite header
            tempCell.rightWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).Weight;
            if (tempCell.text.Contains("[TextHere]")) { tempCell.text = tempCell.text.Split(':')[0]; titleBlock.Add(tempCell); return tempCell; }
            else if (tempCell.text.Contains("[HeaderHere]")) { tempCell.text = tempCell.text.Split('[')[0]; headerRow.Add(tempCell); tempCell.info = tempCell.text; return tempCell; }
            else if (tempCell.text.Contains("[BodyHere]"))
            {
                if (currentRow != row)
                {
                    bodyRows.Add(new List<myCell>());
                    bodyColors.Add(cell[1, 1].Interior.Color);
                }
                currentRow = row;
                if (tempCell.color != 16777215) tempCell.moreThanText = true;
                if (tempCell.rightLineStyle != XlLineStyle.xlLineStyleNone) // only add border style to cell class if it is other than the expected default.
                {
                    tempCell.moreThanText = true;
                    tempCell.rightLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).LineStyle; // ignoring top border because it would overwrite header
                    tempCell.rightWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).Weight;
                    tempCell.bottomLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).LineStyle;
                    tempCell.bottomWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).Weight;
                    tempCell.leftLineStyle = (XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).LineStyle;
                    tempCell.leftWeight = (XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).Weight;
                }
                bodyRows[bodyRows.Count - 1].Add(tempCell);
                return tempCell;
            }
            else if (tempCell.text.Contains("[Sort]"))
            {
                string[] sortInfo = tempCell.text.Split(']')[1].Split('('); // possible cell content [Sort](1)P,C,CN,L,R(4)incrementing
                int sortNum = sortInfo.Length;
                for (int a = 1; a < sortNum; a++)
                {
                    sortOrder.Add(new List<string>());
                    string[] tempDelimiter = sortInfo[a].Split(')')[1].Split(',');
                    sortOrder[sortOrder.Count - 1].Add(sortInfo[a].Split(')')[0]);
                    sortOrder[sortOrder.Count - 1].Add((column-1).ToString());
                    foreach (string tempString in tempDelimiter) sortOrder[sortOrder.Count - 1].Add(tempString); // should make above cell entry into { {1, *column* , *sortType* P,C,CN,L,R} , {4, *column*, incrementing} }
                }
                return tempCell;
            }
            else if (tempCell.text.Contains("[Group]"))
            {
                group.Add(column-1);
            }
            else if (tempCell.text.Contains("[Quantity]"))
            {
                quantity = column-1;
            }
            else if (tempCell.text.Contains("[EndRow]"))
            {
                rowEnd = row;
                return tempCell;
            }
            else if (tempCell.text.Contains("[EndColumn]"))
            {
                columnEnd = column;
                return tempCell;
            }
            return tempCell;
        }        
    }
}