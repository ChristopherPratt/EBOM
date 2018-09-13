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
    public class excelSection
    {

        public int currentRow;
        public List<string> TBtext; //title block text
        public List<int> TBrowIndex;
        public List<int> TBcolumnIndex;
        public List<string> Htext; // header text
        public List<int> HrowIndex;
        public List<int> HcolumnIndex;
        public List<List<string>> text; // body rows text
        public List<List<double>> textColor;
        public List<List<string>> textFont;
        public List<List<int>> textSize;
        public List<List<double>> backGroundColor;
        public List<List<int>> rowIndex;
        public List<List<int>> columnIndex;
        public List<List<XlLineStyle>> topLineStyle;
        public List<List<XlBorderWeight>> topWeight;
        public List<List<XlLineStyle>> rightLineStyle;
        public List<List<XlBorderWeight>> rightWeight;
        public List<List<XlLineStyle>> bottomLineStyle;
        public List<List<XlBorderWeight>> bottomWeight;
        public List<List<XlLineStyle>> leftLineStyle;
        public List<List<XlBorderWeight>> leftWeight;
        public int footerColor;
        public List<string> footerList;
        public int quantityColumn;

        public excelSection()
        {

            currentRow = 0;
            TBtext = new List<string>();
            TBrowIndex = new List<int>();
            TBcolumnIndex = new List<int>();
            Htext = new List<string>();
            HrowIndex = new List<int>();
            HcolumnIndex = new List<int>();
            text = new List<List<string>>();
            textColor = new List<List<double>>();
            textFont = new List<List<string>>();
            textSize = new List<List<int>>();
            backGroundColor = new List<List<double>>();
            rowIndex = new List<List<int>>();
            columnIndex = new List<List<int>>();
            topLineStyle = new List<List<XlLineStyle>>();
            topWeight = new List<List<XlBorderWeight>>();
            rightLineStyle = new List<List<XlLineStyle>>();
            rightWeight = new List<List<XlBorderWeight>>();
            bottomLineStyle = new List<List<XlLineStyle>>();
            bottomWeight = new List<List<XlBorderWeight>>();
            leftLineStyle = new List<List<XlLineStyle>>();
            leftWeight = new List<List<XlBorderWeight>>();
            footerList = new List<string>();

            addNewList(this);
        }
        public void addNewList(excelSection section)
        {
            section.text.Add(new List<string>());
            section.textColor.Add(new List<double>());
            section.textFont.Add(new List<string>());
            section.textSize.Add(new List<int>());
            section.backGroundColor.Add(new List<double>());
            section.rowIndex.Add(new List<int>());
            section.columnIndex.Add(new List<int>());
            section.topLineStyle.Add(new List<XlLineStyle>());
            section.topWeight.Add(new List<XlBorderWeight>());
            section.rightLineStyle.Add(new List<XlLineStyle>());
            section.rightWeight.Add(new List<XlBorderWeight>());
            section.bottomLineStyle.Add(new List<XlLineStyle>());
            section.bottomWeight.Add(new List<XlBorderWeight>());
            section.leftLineStyle.Add(new List<XlLineStyle>());
            section.leftWeight.Add(new List<XlBorderWeight>());
        }
        public void insertRow(excelSection source, int sourceIndex, excelSection sorted, int sortedIndex)
        {//  insert row from specified index of source list to specified index of received list.
            sorted.text.Insert(sortedIndex, source.text[sourceIndex]);
            sorted.textColor.Insert(sortedIndex, source.textColor[sourceIndex]);
            sorted.textFont.Insert(sortedIndex, source.textFont[sourceIndex]);
            sorted.textSize.Insert(sortedIndex, source.textSize[sourceIndex]);
            sorted.backGroundColor.Insert(sortedIndex, source.backGroundColor[sourceIndex]);
            sorted.rowIndex.Insert(sortedIndex, source.rowIndex[sourceIndex]);
            sorted.columnIndex.Insert(sortedIndex, source.columnIndex[sourceIndex]);
            sorted.topLineStyle.Insert(sortedIndex, source.topLineStyle[sourceIndex]);
            sorted.topWeight.Insert(sortedIndex, source.topWeight[sourceIndex]);
            sorted.rightLineStyle.Insert(sortedIndex, source.rightLineStyle[sourceIndex]);
            sorted.rightWeight.Insert(sortedIndex, source.rightWeight[sourceIndex]);
            sorted.bottomLineStyle.Insert(sortedIndex, source.bottomLineStyle[sourceIndex]);
            sorted.bottomWeight.Insert(sortedIndex, source.bottomWeight[sourceIndex]);
            sorted.leftLineStyle.Insert(sortedIndex, source.leftLineStyle[sourceIndex]);
            sorted.leftWeight.Insert(sortedIndex, source.leftWeight[sourceIndex]);
        }
        public void addRow(excelSection source, int insertIndex, excelSection sorted)
        { // add row from source list to the end of new received list
            sorted.text.Add(source.text[insertIndex]);
            sorted.textColor.Add(source.textColor[insertIndex]);
            sorted.textFont.Add(source.textFont[insertIndex]);
            sorted.textSize.Add(source.textSize[insertIndex]);
            sorted.backGroundColor.Add(source.backGroundColor[insertIndex]);
            sorted.rowIndex.Add(source.rowIndex[insertIndex]);
            sorted.columnIndex.Add(source.columnIndex[insertIndex]);
            sorted.topLineStyle.Add(source.topLineStyle[insertIndex]);
            sorted.topWeight.Add(source.topWeight[insertIndex]);
            sorted.rightLineStyle.Add(source.rightLineStyle[insertIndex]);
            sorted.rightWeight.Add(source.rightWeight[insertIndex]);
            sorted.bottomLineStyle.Add(source.bottomLineStyle[insertIndex]);
            sorted.bottomWeight.Add(source.bottomWeight[insertIndex]);
            sorted.leftLineStyle.Add(source.leftLineStyle[insertIndex]);
            sorted.leftWeight.Add(source.leftWeight[insertIndex]);
        }
        
       
    }
}
