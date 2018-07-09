using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Text.RegularExpressions;

namespace WindowsFormsApp1
{
    class LoadXML
    {
        List<string> properties;
        List<string> headers;
        List<string> propertyValues;
        List<List<string>> unsorted;
        public List<List<LoadTemplate.myCell>> sorted;
        XmlDocument xmlRead;
        XmlNodeList nodeList;
        string apcbField = "";

        LoadTemplate template;

        public int totalPartCount;
        public LoadXML(LoadTemplate l)
        {
            template = l;
            properties = new List<string>();
            propertyValues = new List<string>();
            headers = new List<string>();
            unsorted = new List<List<string>>();
            sorted = new List<List<LoadTemplate.myCell>>();
            openXML();
            getHeaderTitles();
            getPartInfo();
            sort();
        }
        public void openXML()
        {
            xmlRead = new XmlDocument();
            xmlRead.Load(System.AppDomain.CurrentDomain.BaseDirectory + "altium.xml");
        }
        public string removeUnderscore(string text)
        {
            string newString = "";
            string[] temp = text.Split('_');
            if (temp.Length == 1) return text;
            else
            {
                newString = temp[0];
                for (int a = 1; a < temp.Length; a++)
                    newString += " " + temp[a];
            }
            return newString;
        }
        public void getHeaderTitles()
        {
            //Read the headers
            nodeList = xmlRead.SelectNodes("GRID")[0].SelectNodes("COLUMNS")[0].SelectNodes("COLUMN");
            if (nodeList.Count == 0)
            {
                throw new Exception("No nodes found in xml file when looking for headers, " + System.AppDomain.CurrentDomain.BaseDirectory + "altium.xml");
            }

            

            foreach (XmlNode node in nodeList) //for every header loaded
            {
                string nodeAttribute = node.Attributes.Item(1).Value;
                for ( int a = 0; a < template.titleBlock.Count; a++)//get indexes for all the headers and the title block
                {
                    if (nodeAttribute.Equals(template.titleBlock[a].text))
                    { template.titleBlock[a].index = Convert.ToInt32(node.Attributes.Item(2).Value); continue; }
                }
                for (int a = 0; a < template.headerRow.Count; a++)
                {
                    if (nodeAttribute.Equals(template.headerRow[a].text))
                    {
                        template.headerRow[a].index = Convert.ToInt32(node.Attributes.Item(2).Value);
                        continue;
                    }
                }

                //if (
                //        nodeAttribute.ToLower().Contains("prod") ||
                //        nodeAttribute.ToLower().Contains("edit") ||
                //        nodeAttribute.ToLower().Contains("addi") || //while not strictly a part of apcb when set, is still included in the export
                //        nodeAttribute.ToLower().Contains("appr") ||
                //        nodeAttribute.ToLower().Contains("modi") ||
                //        nodeAttribute.ToLower().Contains("titl") ||
                //        nodeAttribute.ToLower().Contains("cust") ||
                //        nodeAttribute.ToLower().Contains("rele")
                //  )
                //{
                //    properties.Add(nodeAttribute); //is a property value


                //}
                //else
                //{
                //    if (nodeAttribute.ToLower().Contains("desi"))
                //    {
                //        apcbField = nodeAttribute;
                //    }
                //    headers.Add(nodeAttribute);
                //}
            }
            
        }
        public void getPartInfo()
        {
            //Read the parts
            nodeList = xmlRead.SelectNodes("GRID")[0].SelectNodes("ROWS")[0].SelectNodes("ROW");
            if (nodeList.Count == 0)
            {
                throw new Exception("No nodes found in xml file when looking for parts, " + System.AppDomain.CurrentDomain.BaseDirectory + "altium.xml");
            }
            totalPartCount = nodeList.Count;

            //Use the first node to read the property info

            //Each node is a part, add the part to the list
            
            for (int a = 0; a < template.titleBlock.Count; a++)
            {
                 
                template.titleBlock[a].info = nodeList[0].Attributes[template.titleBlock[a].index].Value;
                if (template.titleBlock[a].info.Contains("ProjectAdditionalNote")) template.titleBlock[a].info = "";
            }
            int rowIndex = template.bodyRows[0][0].rowIndex;
            List<LoadTemplate.myCell> body1Default = template.bodyRows[0];

            template.bodyRows = new List<List<LoadTemplate.myCell>>();
            for (int row = 0; row < nodeList.Count; row++)
            {
                List<LoadTemplate.myCell> body1 = new List<LoadTemplate.myCell>();
                copyMyCellList(body1, body1Default);
                
                template.bodyRows.Add(body1);
                for (int column = 0; column < template.headerRow.Count; column++)
                {

                    template.bodyRows[row][column].info = nodeList[row].Attributes[template.headerRow[column].index].Value;
                    if (template.bodyRows[row][column].info.Contains("ProjectAdditionalNote") || template.bodyRows[row][column].info.Contains("Projectchangeindex")) template.bodyRows[row][column].info = "";
                    template.bodyRows[row][column].rowIndex = rowIndex;
                }
                rowIndex++;

                
            }






            //foreach (XmlNode node in nodeList) //for every part
            //{
            //    if (!string.IsNullOrEmpty(apcbField)) //nothing will be loaded if apcb is empty!
            //    {
            //        if (!propertiesLoaded && node.Attributes.GetNamedItem(apcbField).Value == "APCB")
            //        {
            //            //ignore exceptions in this area, as they are from non existent property values-
            //            //not an issue, the field will just be blank
            //            foreach (string property in properties)
            //            {

            //                propertyValues.Add(node.Attributes.GetNamedItem(property).Value);
            //            }
            //        }
            //    }
            //    propertiesLoaded = true;
            //    List<string> tempAttributes = new List<string>();
            //    //Add each part's attributes to the part object

            //    for (int i = 0; i < headers.Count; i++) //for every header attribute, including the properties
            //    {

            //        tempAttributes.Add(node.Attributes.GetNamedItem(headers[i]).Value);
            //    }
            //    unsorted.Add(tempAttributes);
            //    //tempAttributes.Clear(); //Memory management
            //}
            //Console.WriteLine("Unsorted parts = " + unsorted.Count);
        }

        public void copyMyCellList(List<LoadTemplate.myCell> cellList, List<LoadTemplate.myCell> OldCellList)
        {
            for (int a = 0; a < OldCellList.Count; a++)
            {
                cellList.Add(new LoadTemplate.myCell());
                cellList[a].color = OldCellList[a].color;
                cellList[a].columnIndex = OldCellList[a].columnIndex;
                cellList[a].rightLineStyle = OldCellList[a].rightLineStyle;
                cellList[a].rightWeight = OldCellList[a].rightWeight;
                cellList[a].bottomLineStyle = OldCellList[a].bottomLineStyle;
                cellList[a].bottomWeight = OldCellList[a].bottomWeight;
                cellList[a].leftLineStyle = OldCellList[a].leftLineStyle;
                cellList[a].leftWeight = OldCellList[a].leftWeight;
                cellList[a].moreThanText = OldCellList[a].moreThanText;
            }


        }

        //public void initialSorting(List<List<LoadTemplate.myCell>> bodyList, List<List<string>> sort)
        //{
        //    List<int> columnOrderIndexes = new List<int>();
        //    List<string> temp = new List<string>();
        //    for (int a = 0; a < sort.Count; a++)
        //    {
        //        columnOrderIndexes.Add(Convert.ToInt32(sort[a][1]));
        //    }





        //    Regex re = new Regex(@"([a-zA-Z]+)(\d+)");




            
        //    foreach (List<LoadTemplate.myCell> row in bodyList)
        //    {
        //        Match result = re.Match(row[designatorIndex].info);
        //        temp.Add(result.Groups[1].Value + "_" + row[packageIndex].info + "_" + row[valueIndex].info + "_" + result.Groups[2].Value);
        //    }
        //}

        //public void sorting(List<List<LoadTemplate.myCell>> list, int sortColumn, List<string> delimiter)
        //{
        //    List<List<LoadTemplate.myCell>> sortedList = new List<List<LoadTemplate.myCell>>();
        //    bool matched = false;

        //    sortedList.Add(list[0]);
        //    for (int a = 0; a < delimiter.Count; a++) // loop through the full list of parts for each delimiter and add it to the sortedList in order.
        //    {
        //        for (int b = 0; b < list.Count; b++) //looping through the list
        //        {
                    
        //            if (compareDelimiters(delimiter, a, list[b][sortColumn].info, list, sortedList)) sortedList.Add(list[b]);
        //            //if (list[b][sortColumn].info.Equals(delimiter[a])) sortedList.Add(list[b]);

        //        }
        //    }
        //}
        //public bool compareDelimiters(List<string> delimiter, int delimiterIndex, string info, List<List<LoadTemplate.myCell>> list, List<List<LoadTemplate.myCell>> sortedList)
        //{
        //    List<List<LoadTemplate.myCell>> tempSortedList = new List<List<LoadTemplate.myCell>>();

        //    List<string> similarDelimiters = new List<string>();
        //    int delimiterLength = delimiter[delimiterIndex].Length;
        //    bool matched = false;

        //    if (delimiter[delimiterIndex].Equals("incrementing"))
        //    {
                
        //    }
        //    else if (delimiter[delimiterIndex].Equals("descending"))
        //    {

        //    }
        //    else
        //    {
        //        for (int x = 0; x < delimiter[delimiterIndex].Length; x++) // loop through each character of a delimiter we want to test
        //        {
        //            if (delimiter[delimiterIndex][x].Equals(info[x])) matched = true; // if the character matches the cells text character then continue, if not, then break
        //            else { matched = false; break; }
        //        }
        //        if (matched) // now we need to test and make sure that there aren't other delimiters which are similar to the one we tested.
        //        {
        //            for (int y = 0; y < delimiter.Count; y++) // loop through all delimiters that we are sorting
        //            {
        //                if (y == delimiterIndex) continue; // we don't want to compare the delimeter we already compared.
        //                for (int z = 0; z < delimiter[y].Length; z++)
        //                {
        //                    if (delimiter[y][z].Equals(info[z])) matched = true; // see if there are any other delimiters which match the cell text that we were comparing earlier
        //                    else { matched = false; break; } 
        //                }
        //                if (matched && delimiter[y].Length > delimiterLength) { }//return false; // if another longer delimeter is a better match for the cell text, return false, but if we 
        //                //else return true;                                                   //were right the first time and there are no better matches, return false;
        //            }
        //        }
        //    }
        //    return false;
        //        //for (int a = 0; a < info.Length; a++) // for each character of text in question
        //        //{
        //        //    if (delimiter[delimiterIndex][a].Equals(info[a]))
        //        //    {
        //        //        for (int b = 0; b > delimiter.Length; b++)
        //        //        {
        //        //            if (b == delimiterIndex) continue; // skip this index since we already compared this one successfully
        //        //            if (delimiter[b][a].Equals(info[a]))
        //        //            {
        //        //            }
        //        //        }
        //        //    }
        //        //}

        //}

        public bool sorting()
        {
            return false;
        }

        public void sort()
        {
            int quantityIndex = 0;
            int valueIndex = 0;
            int packageIndex = 0;
            int designatorIndex = 0;

            bool matched = false;
            for (int a = 0; a < template.headerRow.Count; a++)
            {
                if (template.headerRow[a].text.ToLower().Contains("quan") || template.headerRow[a].text.ToLower().Contains("qty")) quantityIndex = a; // get the indexes or certain header columns
                //if (template.headerRow[a].text.ToLower().Contains("val")) valueIndex = a;
                //if (template.headerRow[a].text.ToLower().Contains("pack")) packageIndex = a;
                //if (template.headerRow[a].text.ToLower().Contains("desi")) designatorIndex = a;
            }

            List<List<string>> newSortOrder = new List<List<string>>();
            for (int a = 0; a < template.sortOrder.Count; a++)
            {
                for (int b = 0; b < template.sortOrder.Count; b++)
                {
                    if (template.sortOrder[b][0].Equals((a + 1).ToString())) { newSortOrder.Add(template.sortOrder[b]); break; }
                }
            }
            template.sortOrder = newSortOrder;

            int matchedCount = 0;
            sorted.Add(template.bodyRows[0]);
            for (int a = 1; a < template.bodyRows.Count; a++) // for each component
            {
                matchedCount = 0;
                for (int b = 0; b < sorted.Count; b++) // for each sorted component
                {
                    int tempMatchedCount = 0;
                    for (int c = 0; c < template.sortOrder.Count; c++) // for each sort method
                    {
                        if (sorting(template.bodyRows[a],  ))
                    }
                }
            }



            //sorting(template.bodyRows);

            Regex re = new Regex(@"([a-zA-Z]+)(\d+)");
            

            //List<string> temp = new List<string>();
            //foreach (List<LoadTemplate.myCell> row in template.bodyRows)
            //{
            //    Match result = re.Match(row[designatorIndex].info);
            //    temp.Add(result.Groups[1].Value + "_" + row[packageIndex].info + "_" + row[valueIndex].info + "_" + result.Groups[2].Value);
            //}
            //temp.Sort();

            //initialSorting(template.bodyRows, template.sortOrder);

            

            //initialSorting(template.bodyRows, template.sortOrder);

            //foreach (List<string> mySort in newSortOrder) sorting(template.bodyRows, Convert.ToInt32(mySort[1]), mySort);
            
            ////////////////////////////// create unique lists of all designators ///////////////////////////////// BOM<Designator<Rows<Cells>>>
            List<List<List<LoadTemplate.myCell>>> bodyDesignatorList = new List<List<List<LoadTemplate.myCell>>>();
            char[] currentDesignator = new char[] { ' ' }; // placeholder for first example
            List<char[]> charDesignatorList = new List<char[]>();
            bool match = false;
            int charLength = 0;
            for (int a = 1; a < template.bodyRows.Count; a++) //make a  list which is a list of all rows that have the same first letter designator
            {
                if (template.bodyRows[a][designatorIndex].info.Contains("PCB")) template.bodyRows[a][designatorIndex].info = "PCB"; // old way EBOM was make had PCB designator be APCB. we are changing that to PCB like it ought.
                if (!Char.IsLetter(template.bodyRows[a][designatorIndex].info[1])) charLength = 1; else charLength = 2; // get first letter of new designator 
                char[] newDesignator = new char[charLength];
                for (int d = 0; d < charLength; d++) newDesignator[d] = template.bodyRows[a][designatorIndex].info[d]; // certain designators have 2 letters, like CN1 so we need to differentiate them from capacitors and the like
                if (!currentDesignator.SequenceEqual(newDesignator))
                {
                    for (int b = 0; b < charDesignatorList.Count; b++) if (charDesignatorList[b].SequenceEqual(newDesignator)) { bodyDesignatorList[b].Add(template.bodyRows[a]); match = true; break; }
                    if (match){ match = false; continue; }
                    bodyDesignatorList.Add(new List<List<LoadTemplate.myCell>>());
                    bodyDesignatorList[bodyDesignatorList.Count - 1].Add(template.bodyRows[a]);
                    charDesignatorList.Add(newDesignator);
                }
                else bodyDesignatorList[bodyDesignatorList.Count - 1].Add(template.bodyRows[a]);
            }
            ////////////////////////////// create unique lists of all designators /////////////////////////////////

            ////////////////////////////// Sort designator by parts with same value///////////////////////////////// BOM<Designator<model<parts<Cells>>>>
            List<List<List<List<LoadTemplate.myCell>>>> valueDesignatorList = new List<List<List<List<LoadTemplate.myCell>>>>();
            int count = 1;
            for (int a = 0; a < bodyDesignatorList.Count; a++) // loop through all designator lists
            {
                valueDesignatorList.Add(new List<List<List<LoadTemplate.myCell>>>());// add new designator list
                valueDesignatorList[a].Add(new List<List<LoadTemplate.myCell>>());// add new model to list

                valueDesignatorList[a][valueDesignatorList[a].Count - 1].Add(bodyDesignatorList[a][0]); // add first row of new model 
                for (int b = 1; b < bodyDesignatorList[a].Count; b++) // loop through the rows of a specific designator list
                {
                    //bodyDesignatorList[a][b][quantityIndex].info = "";
                    matched = false;
                    if (bodyDesignatorList[a][b][valueIndex].info == bodyDesignatorList[a][b - 1][valueIndex].info)
                        //if (bodyDesignatorList[a][b][packageIndex].info == bodyDesignatorList[a][b - 1][packageIndex].info) matched = true; // make sure that the two components are identical
                        matched = true; // make sure that the two components are identical
                    else matched = false;
                    

                    if (matched)
                    {
                        valueDesignatorList[a][valueDesignatorList[a].Count - 1].Add(bodyDesignatorList[a][b]); //if both components are the same then add it to the component list
                        count++;
                    }
                    else
                    {
                        //modelDesignatorList[a][modelDesignatorList.Count - 1][0][quantityIndex].info = count.ToString();
                        //count = 1;
                        valueDesignatorList[a].Add(new List<List<LoadTemplate.myCell>>());// add new model to list
                        valueDesignatorList[a][valueDesignatorList.Count - 1].Add(bodyDesignatorList[a][b]);

                    }
                }
            }
            ////////////////////////////// Sort designator by parts with same value /////////////////////////////////

            ////////////////////////////// Sort value by parts with same package ///////////////////////////////// BOM<Designator<value<package<parts<Cells>>>>>

            matched = false;
            string previousPackage = "";
            List<List<List<List<List<LoadTemplate.myCell>>>>> packageDesignatorList = new List<List<List<List<List<LoadTemplate.myCell>>>>>();
            for (int a = 0; a < valueDesignatorList.Count; a++)
            {
                packageDesignatorList.Add(new List<List<List<List<LoadTemplate.myCell>>>>());// add new designator list
                for (int b = 0; b < valueDesignatorList[a].Count; b++)
                {
                    packageDesignatorList[a].Add(new List<List<List<LoadTemplate.myCell>>>()); // add new value to list
                    packageDesignatorList[a][b].Add(new List<List<LoadTemplate.myCell>>()); // add new package to list
                    if (packageDesignatorList[a][b][packageDesignatorList[a][b].Count-1][packageIndex].Equals(previousPackage))
                    {
                        packageDesignatorList[]
                    }
                    
                }
            }
            ////////////////////////////// Sort value by parts with same package /////////////////////////////////
            //sort the original 2d array list multiple times instead of creating a 5 dimensional array. it will save memory and be faster and less prone to error hopefully.

            //    int count = 0;
            ////temp.Add(template.bodyRows[0]);
            //for (int a = 1; a < template.bodyRows.Count; a++)
            //{
            //    template.bodyRows[a][quantityIndex].text = "";
            //    matched = false;
            //    if (template.bodyRows[a][valueIndex].text == template.bodyRows[a - 1][valueIndex].text)
            //        if (template.bodyRows[a][packageIndex].text == template.bodyRows[a - 1][packageIndex].text) matched = true; // make sure that the two components are identical
            //        else matched = false;
            //    else matched = false;

            //    if (matched)
            //        //temp.Add(unsorted[a]); //if both components are the same then add it to the component list
            //        count++;
            //    else
            //    {
            //        //temp[0][quantityIndex] = temp.Count.ToString(); // if both components aren't the same then change the quantity of the top component to the amount of all components
            //        //sorted.Add(temp);
            //        //temp = new List<List<string>> { unsorted[a] };
            //        template.bodyRows[a-count-1][quantityIndex].text = count.ToString();
            //        count = 0;
            //    }
            //}

        }
    }  
}

