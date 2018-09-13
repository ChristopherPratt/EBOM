using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EBOM_Creation_Tool_v2
{
    public class countParts
    {
        public int quantityColumn;
        public List<int> groupedColumns;
        public countParts()
        {
            groupedColumns = new List<int>();
        }
        public void updateQuantity(int quantityIndex,List<int> groupedColumns1, ref List<List<string>> sorted)
        {
            // this section changes the row index to what it should be after sorting because we use the row  index to write to the excel file
            // we also set the quantity of similar parts for the top part based on how many of them there are and leave the other quantity cells blank.
            bool matched = false;
            int count = 1;
            for (int a = 1; a < sorted.Count; a++) // loop through all rows
            {
                if (!(quantityIndex == 0 || groupedColumns1.Count == 0))  // we want to make sure that the template allows for updating a quantity by like components.
                {
                    sorted[a][quantityIndex] = ""; // change all rows quantity cell to blank
                    matched = false;
                    for (int b = 0; b < groupedColumns1.Count; b++)
                    {
                        if (sorted[a][groupedColumns1[b]] == sorted[a - 1][groupedColumns1[b]])  // check to see if this row and the previous rows values are identical
                        {
                            if (b == groupedColumns1.Count - 1) matched = true;
                            continue;
                        }
                        else { matched = false; break; }
                    }
                    if (matched) // if both rows matched
                    {
                        count++; // increment count counter describing how many parts in a row are the same part
                        if (a == sorted.Count - 1) sorted[a - count][quantityIndex] = count.ToString(); // add last element since there is nothing to compare it to.
                    }
                    else
                    {
                        sorted[a - count][quantityIndex] = count.ToString(); // if both components aren't the same then change the quantity of the top component to the amount of all components
                        count = 1;
                        if (a == sorted.Count - 1)
                        {
                            sorted[a - count][quantityIndex] = count.ToString(); // since we are in the "else" section we know that we are dealing with a part with no similar parts so we change the row above
                            sorted[a][quantityIndex] = count.ToString();//          this parts quantity and then set the bottom part quantity to 1 because we know its unitque
                        }
                    }
                }

            }
        }
    }
}
