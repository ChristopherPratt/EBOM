using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace EBOMgui
{
    class xmlFileHandler
    {
        public int[] TBindex; // title block index
        public int[] Hindex; // header index
        public int totalPartCount;
        public List<string> titleBlockInfo;
        public List<List<string>> componentAttributes;
        public string exportFileName;
        public List<int> attributeIndexes;
        public List<string> attributeNames;


        public void run(MainFrame mainFrame1, string filePath)
        {
            
            try
            {
                //instantiate all the necessary lists and objects for handling the XML
                componentAttributes = new List<List<string>>();
                attributeIndexes = new List<int>();
                attributeNames = new List<string>();
                titleBlockInfo = new List<string>();
                XmlDocument xmlRead = new XmlDocument();


                openXML(xmlRead, filePath); // load the xml file into the xmlRead object
                XmlNodeList partAttributesNodeList = xmlRead.SelectNodes("GRID")[0].SelectNodes("COLUMNS")[0].SelectNodes("COLUMN");
                XmlNodeList componentNodeList = xmlRead.SelectNodes("GRID")[0].SelectNodes("ROWS")[0].SelectNodes("ROW");

                getColumnNamesandIndexes(partAttributesNodeList,ref attributeNames,ref attributeIndexes);
                totalPartCount = getComponentInfo(componentNodeList, ref componentAttributes, attributeIndexes);
                mainFrame1.writeToConsole("Finished reading xml file.");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                throw e;
            }
        }


        public void openXML(XmlDocument xmlRead1, string filePath1)
        {
            try { xmlRead1.Load(filePath1); }
            catch
            {
                throw new Exception("Source file path incorrect");
            }
        }
        // get certain node attrubutes and save it to a string list.
        public List<string> readNode(XmlNode node, List<int> indexList)
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
        
        //get all the info needed for the titleblock. we only need the top row for this.
        public void getColumnNamesandIndexes(XmlNodeList nodes, ref List<string> attributeNames1, ref List<int> attributeIndexes1)
        {
            List<string> temp = new List<string>();
            List<int> indexList = new List<int>{ 1, 5 };
            foreach (XmlNode node in nodes)
            {
                temp = readNode(node, indexList);
                attributeNames1.Add(temp[0]);
                attributeIndexes1.Add(Convert.ToInt32(temp[1]));
            }

        }
        // get all the component info according to the index list and add it to a 2 dimesional string list to represent the body of the EBOM
        public int getComponentInfo(XmlNodeList nodeList, ref List<List<string>> componentInfo, List<int> attributeIndexes2)
        {
            int totalParts = 0;
            foreach (XmlNode node in nodeList)
            {
                //int indexList = node.Attributes.Count;
                componentInfo.Add(readNode(node, attributeIndexes2));
                totalParts++;
            }
            return totalParts;
        }
    }
}
