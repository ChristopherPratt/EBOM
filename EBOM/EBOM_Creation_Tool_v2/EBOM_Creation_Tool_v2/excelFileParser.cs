using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace EBOM_Creation_Tool_v2
{
    class excelFileParser
    {
        public void parseCell(Range cell, int row, int column, ref int totalRows2, ref int totalColumns2, ref excelSection template3, ref sort sort3, ref countParts countParts2 )
        {
            /////////////////////// find and parse tag //////////////////////////
            string tag = parseTag(cell[1,1]);
            
            switch (tag)
            {
                case "TextHere":
                    {
                        populateTitleBlockCell(cell[1, 1], ref template3); cell[1, 1].Value = "";
                        break;
                    }
                case "HeaderHere":
                    {
                        populateHeaderList(cell[1, 1], ref template3); cell[1, 1].Value = "";
                        break;
                    }
                case "BodyHere":
                    {
                        populateBodySectionCell(cell[1, 1], ref template3);  cell[1, 1].Value = "";
                        break;
                    }
                case "Sort":
                    {
                        parseSort(cell[1, 1], ref sort3); cell[1, 1].Value = "";
                        break;
                    }
                case "Footer":
                    {
                        getFooterInfo(cell[1, 1], ref template3.footerColor, ref template3.footerList); cell[1, 1].Interior.Color = 16777215; cell[1, 1].Value = "";
                        break;
                    }
                case "Group":
                    {
                        getQuantityGrouping(column, ref countParts2.groupedColumns); cell[1, 1].Value = "";
                        break;
                    }
                case "Quantity":
                    {
                        getQuantityColumn(column, ref countParts2.quantityColumn); cell[1, 1].Value = "";
                        break;
                    }
                case "EndRow":
                    {
                        totalRows2 = row; cell[1, 1].Value = ""; // change loop iterator limit variable for rows if EndRow tag is found
                        break;
                    }
                case "EndColumn": 
                    {
                        totalColumns2 = column; cell[1, 1] = ""; // change loop iterator limit variable for columns if EndColumn tag is found
                        break;
                    }
            }
        }

        /////////////////////// find and parse tag //////////////////////////
        private string parseTag(Range tempText)
        {
            if (tempText.Text.Contains("["))
            {

                string[] tempSplit = tempText.Text.Split('[');
                for (int a = 0; a < tempSplit.Length; a++)
                    if (tempSplit[a].Contains("]"))
                        return tempSplit[a].Split(']')[0];
            }
            else return "";
            return "";
        }

        public void populateTitleBlockCell(Range cell, ref excelSection excel)
        {
            excel.TBtext.Add(cell[1, 1].Text.Split(':')[0]);
            excel.TBrowIndex.Add(cell[1, 1].Row);
            excel.TBcolumnIndex.Add(cell[1, 1].Column);
        }
        public void populateHeaderList(Range cell, ref excelSection excel)
        {
            excel.Htext.Add(cell[1, 1].Text.Split('[')[0]);
            excel.HrowIndex.Add(cell[1, 1].Row);
            excel.HcolumnIndex.Add(cell[1, 1].Column);
        }
        public void populateBodySectionCell(Range cell, ref excelSection excel)
        {
            int index = excel.rowIndex.Count - 1;

            if (excel.rowIndex[index].Count > 0)
                if (excel.rowIndex[index][0] != cell[1, 1].Row)
                {
                    excel.addNewList(excel);
                    index = excel.rowIndex.Count - 1;
                }

            excel.rowIndex[index].Add(cell[1, 1].Row);
            excel.columnIndex[index].Add(cell[1, 1].Column);
            excel.backGroundColor[index].Add(cell[1, 1].Interior.Color);

            excel.textSize[index].Add((int)cell[1, 1].Font.Size);
            excel.textFont[index].Add((string)cell[1, 1].Font.Name);
            excel.textColor[index].Add((double)cell[1, 1].Font.Color);

            //excel.rightLineStyle[index].Add((XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).LineStyle); // ignoring top border because it would overwrite header
            excel.rightLineStyle[index].Add(new XlLineStyle()); // ignoring top border because it would overwrite header

            excel.rightWeight[index].Add((XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).Weight);
            excel.bottomLineStyle[index].Add((XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).LineStyle);
            excel.bottomWeight[index].Add((XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).Weight);
            excel.leftLineStyle[index].Add((XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).LineStyle);
            excel.leftWeight[index].Add((XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).Weight);



            //if ((XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).LineStyle != XlLineStyle.xlLineStyleNone) // only add border style to cell class if it is other than the expected default.
            //{
            //    excel.rightLineStyle[index].Add((XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).LineStyle); // ignoring top border because it would overwrite header
            //    excel.rightWeight[index].Add((XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeRight).Weight);
            //    excel.bottomLineStyle[index].Add((XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).LineStyle);
            //    excel.bottomWeight[index].Add((XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeBottom).Weight);
            //    excel.leftLineStyle[index].Add((XlLineStyle)cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).LineStyle);
            //    excel.leftWeight[index].Add((XlBorderWeight)cell[1, 1].Borders(XlBordersIndex.xlEdgeLeft).Weight);
            //}
            //else
            //{

            //}
        }
        public void parseSort (Range cell, ref sort sort4)
        {
            string[] sortInfo = cell[1, 1].Text.Split(']')[1].Split('('); // possible cell content [Sort](1)P,C,CN,L,R(4)incrementing
            int sortNum = sortInfo.Length;
            for (int a = 1; a < sortNum; a++)
            {                
                string[] tempDelimiter = sortInfo[a].Split(')')[1].Split(',');
                sort4.priority.Add(Convert.ToInt32(sortInfo[a].Split(')')[0]));
                sort4.column.Add(cell[1, 1].Column - 1);
                //if (tempDelimiter[0].Equals("ascending")) sort4.type.Add("ascending");
                //else if (tempDelimiter[0].Equals("descending")) sort4.type.Add("descending");
                //else if ()
                //else sort4.type.Add("Custom");
                sort4.type.Add(tempDelimiter[0]);
                sort4.customSort.Add(new List<string>());
                for(int b = 1; b < tempDelimiter.Length; b++)
                    sort4.customSort[sort4.customSort.Count - 1].Add(tempDelimiter[b]); // should make above cell entry into { {1, *column* , *sortType* P,C,CN,L,R} , {4, *column*, incrementing} }
            }
        }
        public void getFooterInfo(Range cell, ref int footerColor, ref List<string> footerList)
        {
            footerColor = (int)cell[1, 1].Interior.Color;
            footerList.Add(cell[1, 1].Text.Split(']')[1]);
        }
        public void getQuantityGrouping(int column, ref List<int> columnGroup)
        {
            columnGroup.Add(column);
        }
        public void getQuantityColumn(int column, ref int quantityColumn)
        {
            quantityColumn = column;
        }
    }

}
