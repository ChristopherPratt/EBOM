using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop.Excel;


namespace EBOMCreationTool
{
    class LoadXML
    {
        bool end = false;
        List<string> properties;
        List<string> headers;
        List<string> propertyValues;
        List<List<string>> unsorted;
        public List<List<LoadTemplate.myCell>> sorted;
        XmlDocument xmlRead;
        XmlNodeList nodeList;
        string apcbField = "";

        LoadTemplate template;
        MainFrame mainframe;

        public int totalPartCount;
        public string xmlFile;
        public LoadXML(MainFrame m, LoadTemplate l, string XMLfile)
        {
            mainframe = m;
            xmlFile = XMLfile;
            template = l;
            if (mainframe.end) return;

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
            if (mainframe.end) return;
            xmlRead = new XmlDocument();
            //xmlRead.Load(System.AppDomain.CurrentDomain.BaseDirectory + "altium.xml");

            try { xmlRead.Load(@xmlFile); }
            catch (Exception e)
            {
                MessageBox.Show(".XML file is incorrect");
                mainframe.end = true;

            }
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
            if (mainframe.end) return;
            //Read the headers
            nodeList = xmlRead.SelectNodes("GRID")[0].SelectNodes("COLUMNS")[0].SelectNodes("COLUMN");
            if (nodeList.Count == 0)
            {
                throw new Exception("No nodes found in xml file when looking for headers, " + System.AppDomain.CurrentDomain.BaseDirectory + "altium.xml");
            }

            

            foreach (XmlNode node in nodeList) //for every header loaded
            {
                // looking in the caption attribute of node
                string nodeAttribute = Regex.Replace(node.Attributes.Item(1).Value, "[^a-zA-Z]", "").ToUpper();
                for ( int a = 0; a < template.titleBlock.Count; a++)//get indexes for all the headers and the title block
                {
                    string titleCell = Regex.Replace(template.titleBlock[a].text, "[^a-zA-Z]", "").ToUpper(); // no spaces, symbols, or letters, all uppers
                    if (nodeAttribute.Equals(titleCell))
                    { template.titleBlock[a].index = Convert.ToInt32(node.Attributes.Item(5).Value); continue; }
                }
                for (int a = 0; a < template.headerRow.Count; a++)
                {
                    string headerCell = Regex.Replace(template.headerRow[a].text, "[^a-zA-Z]", "").ToUpper(); // no spaces, symbols, or letters, all uppers

                    if (nodeAttribute.Equals(headerCell))
                    {
                        template.headerRow[a].index = Convert.ToInt32(node.Attributes.Item(5).Value);
                        continue;
                    }
                }

            }
            
        }
        public void getPartInfo()
        {
            if (mainframe.end) return;
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
                try
                {
                    template.titleBlock[a].info = nodeList[0].Attributes[template.titleBlock[a].index].Value;
                    if (template.titleBlock[a].info.Contains("ProjectAdditionalNote")) template.titleBlock[a].info = "";
                }
                catch (Exception e)
                {
                    MessageBox.Show("The Index of an attribute in the .XML is greater than the total amount of attributes.\n The .XML file must be remade.\nThe program will now close.");
                    mainframe.end = true;
                    return;
                }                
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
                    try
                    {
                        template.bodyRows[row][column].info = nodeList[row].Attributes[template.headerRow[column].index].Value;
                        if (template.bodyRows[row][column].info.Contains("ProjectAdditionalNote") || template.bodyRows[row][column].info.Contains("Projectchangeindex")) template.bodyRows[row][column].info = "";
                        if (template.bodyRows[row][column].info.Contains("APCB")) template.bodyRows[row][column].info = "PCB";
                        template.bodyRows[row][column].rowIndex = rowIndex;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("The Index of an attribute in the .XML is greater than the total amount of attributes.\n The .XML file must be remade.\nThe program will now close.");                        
                        mainframe.end = true;
                        return;
                    }

                }
                rowIndex++;

                
            }


        }

        public void copyMyCellList(List<LoadTemplate.myCell> cellList, List<LoadTemplate.myCell> OldCellList)
        {
            if (mainframe.end) return;
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
        // sort all body columns by a list of priority soring
        //contents of sortOrder {1, *column*, setorder, P,C,CN,L,R}
        public bool sorting(List<LoadTemplate.myCell> currentRow, List<LoadTemplate.myCell> comparedRow)
        {
            if (end) return false;
            for (int a = 0; a < template.sortOrder.Count; a++)// for each sort found
            {
                string currentCell = getCellText(currentRow, Convert.ToInt32(template.sortOrder[a][1])); // get actual text data from cell
                string comparedCell = getCellText(comparedRow, Convert.ToInt32(template.sortOrder[a][1]));

                if (template.sortOrder[a][2].Equals("setorder.beginning")) // if the user requests a specific order to the cells in a column
                {
                    int currentIndex = getIndexBeginning(template.sortOrder[a], currentCell);
                    int comparedIndex = getIndexBeginning(template.sortOrder[a], comparedCell);// get the indexes of the specific sort order which match the 2 cells we are comparing

                   // if (currentIndex == -1) return false; // -1 means that the current cell text does not fit into the sort algorithm and should be placed at the end
                    if (currentIndex > comparedIndex) // if current index belongs after the compared index, then return false and move on to the next part
                        return false;
                    else if (currentIndex == comparedIndex) //if  they are equal then compare the next sort option
                        continue;
                    else if (currentIndex < comparedIndex) // if current index belongs before the compared index, then return true and insert the current part where compared part is and push compared part down the list
                        return true;                    
                }
                if (template.sortOrder[a][2].Equals("setorder.end")) // if the user requests a specific order to the cells in a column
                {
                    int currentIndex = getIndexEnd(template.sortOrder[a], currentCell);
                    int comparedIndex = getIndexEnd(template.sortOrder[a], comparedCell);// get the indexes of the specific sort order which match the 2 cells we are comparing

                    // if (currentIndex == -1) return false; // -1 means that the current cell text does not fit into the sort algorithm and should be placed at the end
                    if (currentIndex > comparedIndex) // if current index belongs after the compared index, then return false and move on to the next part
                        return false;
                    else if (currentIndex == comparedIndex) //if  they are equal then compare the next sort option
                        continue;
                    else if (currentIndex < comparedIndex) // if current index belongs before the compared index, then return true and insert the current part where compared part is and push compared part down the list
                        return true;
                }
                else if (template.sortOrder[a][2].Equals("ascending"))
                {
                    double sortString = sortAlphanumeric(currentCell, comparedCell);
                    if (sortString > 0)  // if current index belongs after the compared index, then return false and move on to the next part
                        return false;
                    if (sortString == 0) //if  they are equal then compare the next sort option
                        continue;
                    if (sortString < 0)// if current index belongs before the compared index, then return true and insert the current part where compared part is and push compared part down the list
                        return true;
                }
                else if (template.sortOrder[a][2].Equals("descending"))
                {
                    double sortString = sortAlphanumeric(currentCell, comparedCell);
                    if (sortString < 0)// if current index belongs after the compared index, then return false and move on to the next part
                        return false;
                    if (sortString == 0) //if  they are equal then compare the next sort option
                        continue;
                    if (sortString > 0)// if current index belongs before the compared index, then return true and insert the current part where compared part is and push compared part down the list
                        return true;
                }
            }
            return false;
        }
        public string getCellText (List<LoadTemplate.myCell> cell, int index)
        {            
            return cell[index].info;
        }
        // return index of which designator it is.
        public int getIndexBeginning(List<string> sort, string item)
        {
            int currentMaxElement = 0;
            int currentMax = 0;
            int currentElement = 0;
            int currentCompare = 0;
            for (int a = 3; a < sort.Count; a++) // for each sort level
            {
                currentCompare = 0;
                for (int b = 0; b < sort[a].Length; b++) // for each character of a sort
                {
                    if (item.Length >= sort[a].Length) // make sure the item is long enough to compare
                    {
                        if (sort[a][b] == item[b]) // if both items have the same character at the same index
                        {
                            //if (currentElement == a) currentCompare++; // if we are still on the same element increment counter
                            currentCompare++;
                            if (b == sort[a].Length - 1 && currentCompare > currentMax) // if its at the end of the compare item and there are more characters alike then any other sort
                            {
                                currentMaxElement = a; // save over new sort counter
                                currentMax = currentCompare;
                                currentCompare = 0; // reset currentcompare
                            }                           
                        }
                        else { currentCompare = 0; break; }
                    }                    
                }                
            }
            //if (currentCompare > currentMax) currentMaxElement = currentElement; // for last elements
            return currentMaxElement;
        }
        public int getIndexEnd(List<string> sort, string item)
        {
            int currentMaxElement = 0;
            int currentMax = 0;
            int currentElement = 0;
            int currentCompare = 0;
            for (int a = 3; a < sort.Count; a++) // for each sort level
            {
                for (int b = 0; b < sort[a].Length; b++) // for each character of a sort
                {
                    if (item.Length >= sort[a].Length) // make sure the item is long enough to compare
                    {
                        if (sort[a][b] == item[b + (item.Length - sort[a].Length)]) // if both items have the same character at the same index
                        {
                            //if (currentElement == a) currentCompare++; // if we are still on the same element increment counter
                            currentCompare++;
                            if (b == sort[a].Length - 1 && currentCompare > currentMax) // if its at the end of the compare item and there are more characters alike then any other sort
                            {
                                currentMaxElement = a; // save over new sort counter
                                currentMax = currentCompare;
                                currentCompare = 0; // reset currentcompare
                            }
                        }
                        else { currentCompare = 0; break; }
                        //{
                        //    if (currentElement == a) currentCompare++; // if we are still on the same element increment counter
                        //    else
                        //    {
                        //        if (currentCompare > currentMax) // if current sort counter is greater than max counter for another sort
                        //        {
                        //            currentMaxElement = currentElement; // save over new sort counter
                        //            currentMax = currentCompare;
                        //        }
                        //        currentElement = a; currentCompare = 1; // reset counter and save new element sort
                        //    }
                        //}
                    }

                }

            }
            //if (currentCompare > currentMax) currentMaxElement = currentElement; // for last element
            return currentMaxElement;
        }
        //p n m
        // alphanumeric sorting system which prioritizes numeric values
        // for example, under normal string.compare(,) 100nf would go before 20nf. we know this to be wrong because 20 is less than 100 - except string.compare(,) orders character by character, and 1 is less than 2.
        public double sortAlphanumeric(string current, string compare)
        {
            //double fakenum; // need an output for the tryParse function
            //takes groups of characters that are either letters or numbers and compares them to see which belongs before the other for an ascending sort order
            double sortbyGroupsCurrent(string[] shortString, string[] longString) // sub function here is the opposite of bottom function
            {
                for (int a = 0; a < longString.Length; a++) // for each group of either letters or numbers 
                {                    
                    //if (shortString[a].All(c => c >= '0' && c <= '9') && longString[a].All(c => c >= '0' && c <= '9')) // special case where both sets of characters happen to be numbers
                    if (double.TryParse(shortString[a], out double n) && double.TryParse(longString[a], out double m)) // special case where both sets of characters happen to be numbers
                    {
                        double final = Convert.ToDouble(shortString[a]) - Convert.ToDouble(longString[a]); // subract both numbers to determine order. this is the opposite on bottom function
                    if (final != 0) return final; // if both numbers are the same then we must continue sorting the next group of characters in the loop.
                    }
                    else
                    {
                        if (a < shortString.Length) // make sure that we aren't going beyond the bounds of the array
                        {
                            double final = string.Compare(shortString[a], longString[a]); // compares both sets of characters which still priorites numbers
                            if (final != 0) return final;// returns -1 if shortString[a] belongs before longString[a] and 1 if after, 0 if equals. if it is 0, continue in the loop.
                        }
                        else return -1; // return -1 if up to the length of shortString all characters were the same but longString has more characters. if so, shortString belongs before longString
                    }
                }
                return 0;// both groups of characters are identical
            }
            //takes groups of characters that are either letters or numbers and compares them to see which belongs before the other for an ascending sort order
            double sortbyGroupsCompare(string[] shortString, string[] longString)
            {
                for (int a = 0; a < longString.Length; a++) // for each group of either letters or numbers 
                {
                    
                    if (double.TryParse(shortString[a], out double n) && double.TryParse(longString[a], out double m)) // special case where both sets of characters happen to be numbers
                    {
                        double final = Convert.ToDouble(longString[a]) - Convert.ToDouble(shortString[a]);// subract both numbers to determine order. this is the opposite on bottom function
                        if (final != 0) return final;// if both numbers are the same then we must continue sorting the next group of characters in the loop.
                    }
                    else
                    {
                        if (a < shortString.Length) // make sure that we aren't going beyond the bounds of the array
                            {
                                double final = string.Compare(longString[a], shortString[a]); // compares both sets of characters which still priorites numbers
                                if (final != 0) return final;// returns -1 if shortString[a] belongs before longString[a] and 1 if after, 0 if equals. if it is 0, continue in the loop.
                            }
                        else return 1; // return 11 if up to the length of shortString all characters were the same but longString has more characters. if so, shortString belongs before longString
                    }
                }
                return 0; // both groups of characters are identical
            }

            string[] resultCurrent;
            if (current.Split('.').Length > 2) // if the string contains more than 1 decimal then it isn't likely meant to be a floating point number, so split string like there aren't decimals
                resultCurrent = Regex.Matches(current, @"\D+|\d+").Cast<Match>().Select(m => m.Value).ToArray();// splits the string into an array of groups of characters that have elements of which are only sets of numbers OR letters (allows decimals)_
            else resultCurrent = Regex.Matches(current, @"\D+|[0-9\.]+").Cast<Match>().Select(m => m.Value).ToArray();// splits the string into an array of groups of characters that have elements of which are only sets of numbers OR letters

            string[] resultCompare;
            if (current.Split('.').Length > 2)
                resultCompare = Regex.Matches(compare, @"\D+|\d+").Cast<Match>().Select(m => m.Value).ToArray();
            else resultCompare = Regex.Matches(compare, @"\D+|[0-9\.]+").Cast<Match>().Select(m => m.Value).ToArray();


            if (resultCurrent.Length >= resultCompare.Length) return sortbyGroupsCurrent(resultCurrent, resultCompare); // have to use 2 different functions depending on which array has more elements.
            else return sortbyGroupsCompare(resultCompare, resultCurrent);
        }

        public void sort()
        {
            if (mainframe.end) return;
            bool matched = false;
            List<List<string>> newSortOrder = new List<List<string>>();
            for (int a = 0; a < template.sortOrder.Count; a++)
            {
                for (int b = 0; b < template.sortOrder.Count; b++)
                {
                    if (template.sortOrder[b][0].Equals((a + 1).ToString())) { newSortOrder.Add(template.sortOrder[b]); break; }
                }
            }
            template.sortOrder = newSortOrder;
            sorted.Add(template.bodyRows[0]);
            for (int a = 1; a < template.bodyRows.Count; a++)
            {
                for (int b = 0; b < sorted.Count; b++)
                {
                    if (sorting(template.bodyRows[a], sorted[b]))
                    {
                        sorted.Insert(b, template.bodyRows[a]);
                        break;
                    }
                    else
                    {
                        if (sorted.Count - 1 == b)
                        {
                            sorted.Add(template.bodyRows[a]);
                            break;
                        }
                    }


                }
            }
            bool updateQuantity = true;
            if (template.quantity == 1000 || template.group.Count == 0) updateQuantity = false; // we want to make sure that the template allows for updating a quantity by like components.
            // this section changes the row index to what it should be after sorting because we use the row  index to write to the excel file
            // we also set the quantity of similar parts for the top part based on how many of them there are and leave the other quantity cells blank.
            int bodyStart = template.bodyRowStart;
            int count = 1;
            foreach (LoadTemplate.myCell tempMycell in sorted[0]) tempMycell.rowIndex = bodyStart; // loop through all cells in a row and change each cells row index to what it should be
            bodyStart++;// because we are starting with the second element in the body rows we need to increment the start of the body row index by 1
            for (int a = 1; a < sorted.Count; a++) // loop through all rows
            {
                foreach (LoadTemplate.myCell tempMycell in sorted[a]) tempMycell.rowIndex = bodyStart; // loop through all cells in a row and change each cells row index to what it should be
                bodyStart++; // incredment the row index after all cells are updated
                if (updateQuantity)
                {
                    sorted[a][template.quantity].info = ""; // change all rows quantity cell to blank
                    matched = false;
                    for (int b = 0; b < template.group.Count; b++)
                    {
                        if (sorted[a][template.group[b]].info == sorted[a - 1][template.group[b]].info)  // check to see if this row and the previous rows values are identical
                        {
                            if (b == template.group.Count - 1) matched = true;
                            continue;
                        }
                        else { matched = false;  break; }    
                    }
                    if (matched) // if both rows matched
                    {
                        count++; // increment count counter describing how many parts in a row are the same part
                        if (a == sorted.Count - 1) sorted[a - count][template.quantity].info = count.ToString(); // add last element since there is nothing to compare it to.
                    }
                    else
                    {
                        sorted[a - count][template.quantity].info = count.ToString(); // if both components aren't the same then change the quantity of the top component to the amount of all components
                        count = 1;
                        if (a == sorted.Count - 1)
                        {
                            sorted[a - count][template.quantity].info = count.ToString(); // since we are in the "else" section we know that we are dealing with a part with no similar parts so we change the row above
                            sorted[a][template.quantity].info = count.ToString();//          this parts quantity and then set the bottom part quantity to 1 because we know its unitque
                        }
                    }
                }
                
            }
        }
    }  
}

