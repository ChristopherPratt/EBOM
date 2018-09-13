using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace EBOM_Creation_Tool_v2
{
    class xmlFileParser
    {
        // get certain node attrubutes and save it to a string list.
        public List<string> readNode(XmlNode node, int[] indexList)
        {
            List<string> attributeList = new List<string>();
            try
            {
                foreach (int index in indexList)
                {
                    if (node.Attributes[index].Value.Contains("ProjectAdditionalNote") || node.Attributes[index].Value.Contains("Projectchangeindex")) attributeList.Add("");
                    else if (node.Attributes[index].Value.Contains("APCB")) attributeList.Add("PCB");
                    else attributeList.Add(node.Attributes[index].Value);
                }
                return attributeList;
            }
            catch
            {
                throw new Exception("The Index of an attribute in the .XML is greater than the total amount of attributes");
            }            
        }        
        // collect the index of all the headers  and titleblocks we need so we can access them in a component
        public void getColumnHeaderIndex(XmlNode node, List<string> cellNameList, ref int[] index)
        {
            // looking in the caption attribute of node
            try
            {
                string nodeAttribute = Regex.Replace(node.Attributes.Item(1).Value, "[^a-zA-Z]", "").ToUpper();
                for (int a = 0; a < cellNameList.Count; a++)//get indexes for the title block
                {
                    string titleCell = Regex.Replace(cellNameList[a], "[^a-zA-Z]", "").ToUpper(); // no spaces, symbols, or letters, all uppers
                    if (nodeAttribute.Equals(titleCell))
                    { index[a] = Convert.ToInt32(node.Attributes.Item(5).Value); continue; }
                }
            }
            catch
            {
                throw new Exception("The Index of an attribute in the .XML is greater than the total amount of attributes");
            }
        }
        //get all the info needed for the titleblock. we only need the top row for this.
        public void getTitleBlockInfo (XmlNode node, int[] index, ref List<string> titleBlock)
        {
             titleBlock = readNode(node, index);
        }
        // get all the component info according to the index list and add it to a 2 dimesional string list to represent the body of the EBOM
        public int getComponentInfo (XmlNodeList nodeList, int[]index, ref List<List<string>> componentInfo)
        {
            int totalParts = 0;
            foreach (XmlNode node in nodeList)
            {
                componentInfo.Add(readNode(node, index));
                totalParts++;
            }
            return totalParts;
        }
    }
}
