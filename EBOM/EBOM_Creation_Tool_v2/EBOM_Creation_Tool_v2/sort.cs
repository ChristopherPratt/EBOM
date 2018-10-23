using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;


namespace EBOM_Creation_Tool_v2
{
    public class sort
    {
        public List<int> priority;
        public List<int> column;
        public List<string> type;
        public List<List<string>> customSort;
        public List<List<string>> sorted;
        public sort()
        {
            priority = new List<int>();
            column = new List<int>();
            type = new List<string>();
            customSort = new List<List<string>>();

        }
        public void start(mainFrame mainFrame1, sort sort1, List<List<string>> unsorted, List<string> headers, List<int> headerIndex, countParts countParts1 )
        {
            sortTheSortList();
            sorted = new List<List<string>>();
            sorted.Add(unsorted[0]);
            bool insertHere = false;
            for (int a = 1; a < unsorted.Count; a++)
            {
                for (int b = 0; b < sorted.Count; b++)
                {
                    insertHere = getSort(sort1, unsorted[a], sorted[b]);
                    if (insertHere) { insertRow(unsorted[a], b, ref sorted); break; }
                    //else sorted.Add(unsorted[a]);
                }
                if (!insertHere) sorted.Add(unsorted[a]);
            }
            mainFrame1.writeToConsole("finished sorting components.");
            countParts1.updateQuantity(countParts1.quantityColumn, countParts1.groupedColumns, headerIndex[0], ref sorted);
            mainFrame1.writeToConsole("finished updating quantities.");
        }
        private void insertRow(List<string> newRow, int index, ref List<List<string>> sorted)
        {
            sorted.Insert(index, newRow);
        }

        private bool getSort(sort sort1, List<string> unsortedRow, List<string> sortedRow)
        {
            for (int a = 0; a < sort1.type.Count; a++)
            {
                double result = 0;
                switch (sort1.type[a])
                {
                    case "ascending":
                        {
                            result = alphaNumericSort(unsortedRow[sort1.column[a]], sortedRow[sort1.column[a]]);
                            break;
                        }
                    case "descending":
                        {
                            result = alphaNumericSort(unsortedRow[sort1.column[a]], sortedRow[sort1.column[a]]) * -1;
                            break;
                        }
                    case "setorder.beginning":
                        {
                            result = setOrderBeginning(sort1.customSort[a], unsortedRow[sort1.column[a]], sortedRow[sort1.column[a]]);
                            break;
                        }
                    case "setorder.end":
                        {
                            result = setOrderEnd(sort1.customSort[a], unsortedRow[sort1.column[a]], sortedRow[sort1.column[a]]);
                            break;
                        }
                }

                if (result < 0) //unsorted row belongs before sorted row
                    {
                        return true;
                    }
                else if(result == 0) //unsorted cell matched sorted cell - continue testing same row
                    {
                        continue;
                    }
                else if(result > 0) //unsorted row belongs after sorted row
                    {
                        return false;
                    }
                
            }
            return false;
        }
     
        private double alphaNumericSort(string current, string compare)
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
        
        private double setOrderBeginning(List<string> sortOrder, string unsortedCell, string sortedCell)
        {
            int currentIndex = getIndexBeginning(sortOrder, unsortedCell);
            int comparedIndex = getIndexBeginning(sortOrder, sortedCell);// get the indexes of the specific sort order which match the 2 cells we are comparing

            // if (currentIndex == -1) return false; // -1 means that the current cell text does not fit into the sort algorithm and should be placed at the end
            if (currentIndex > comparedIndex) // if current index belongs after the compared index, then return false and move on to the next part
                return 1;
            else if (currentIndex == comparedIndex) //if  they are equal then compare the next sort option
                return 0;
            else if (currentIndex < comparedIndex) // if current index belongs before the compared index, then return true and insert the current part where compared part is and push compared part down the list
                return -1;
            return 0;
        }
        private double setOrderEnd(List<string> sortOrder, string unsortedCell, string sortedCell)
        {
            int currentIndex = getIndexEnd(sortOrder, unsortedCell);
            int comparedIndex = getIndexEnd(sortOrder, sortedCell);// get the indexes of the specific sort order which match the 2 cells we are comparing

            // if (currentIndex == -1) return false; // -1 means that the current cell text does not fit into the sort algorithm and should be placed at the end
            if (currentIndex > comparedIndex) // if current index belongs after the compared index, then return false and move on to the next part
                return 1;
            else if (currentIndex == comparedIndex) //if  they are equal then compare the next sort option
                return 0;
            else if (currentIndex < comparedIndex) // if current index belongs before the compared index, then return true and insert the current part where compared part is and push compared part down the list
                return -1;
            return 0;
        }
        public int getIndexBeginning(List<string> sort, string item)
        {
            int currentMaxElement = 0;
            int currentMax = 0;
            int currentCompare = 0;
            for (int a = 0; a < sort.Count; a++) // for each sort level
            {
                currentCompare = 0;
                for (int b = 0; b < sort[a].Length; b++) // for each character of a sort
                {
                    if (item.Length >= sort[a].Length) // make sure the item is long enough to compare
                    {
                        if (sort[a][b] == item[b]) // if both items have the same character at the same index
                        {
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
            return currentMaxElement;
        }
        public int getIndexEnd(List<string> sort, string item)
        {
            int currentMaxElement = 0;
            int currentMax = 0;
            int currentCompare = 0;
            for (int a = 0; a < sort.Count; a++) // for each sort level
            {
                for (int b = 0; b < sort[a].Length; b++) // for each character of a sort
                {
                    if (item.Length >= sort[a].Length) // make sure the item is long enough to compare
                    {
                        if (sort[a][b] == item[b + (item.Length - sort[a].Length)]) // if both items have the same character at the same index
                        {
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
            return currentMaxElement;
        }
        private void sortTheSortList()
        {
            List<int> tempPriority = new List<int>();
            List<int>  tempColumn = new List<int>();
            List<string>  tempType = new List<string>();
            List<List<string>>  tempCustomSort = new List<List<string>>();

            for (int a = 1; a < priority.Count+1; a++)
            {
                for (int b = 0; b < priority.Count; b++)
                {
                    if (priority[b].Equals(a))
                    {
                        tempPriority.Add(priority[b]);
                        tempColumn.Add(column[b]);
                        tempType.Add(type[b]);
                        tempCustomSort.Add(customSort[b]);
                    }
                }
            }
            priority = tempPriority;
            column = tempColumn;
            type = tempType;
            customSort = tempCustomSort;
        }
    }
}
