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

namespace WindowsFormsApp1
{
    class LoadXML
    {
        List<string> properties;
        List<string> headers;
        List<string> propertyValues;
        List<List<string>> unsorted;
        public List<List<List<LoadTemplate.myCell>>> sorted;
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
            sorted = new List<List<List<LoadTemplate.myCell>>>();
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
                if (template.headerRow[a].text.ToLower().Contains("val")) valueIndex = a;
                if (template.headerRow[a].text.ToLower().Contains("pack")) packageIndex = a;
                if (template.headerRow[a].text.ToLower().Contains("desi")) designatorIndex = a;
            }
            
            List<List<List<LoadTemplate.myCell>>> bodyDesignatorList = new List<List<List<LoadTemplate.myCell>>>();
            string currentDesignator = "";
            List<char[]> charDesignatorList = new List<char[]>();
            bool match = false;
            int charLength = 0;
            for (int a = 1; a < template.bodyRows.Count; a++) //make a  list which is a list of all rows that have the same first letter designator
            {
                if (template.bodyRows[a][designatorIndex].info.Contains("PCB")) template.bodyRows[a][designatorIndex].info = "PCB"; // old way EBOM was make had PCB designator be APCB. we are changing that to PCB like it ought.
                if (!Char.IsLetter(template.bodyRows[a][designatorIndex].info[1])) charLength = 1; else charLength = 2; // get first letter of new designator 
                char[] newDesignator = new char[charLength];
                for (int d = 0; d < charLength; d++) newDesignator[d] = template.bodyRows[a][designatorIndex].info[d]; // certain designators have 2 letters, like CN1 so we need to differentiate them from capacitors and the like
                if (!currentDesignator.Equals(template.bodyRows[a][designatorIndex].info[0]))
                {
                    for (int b = 0; b < charDesignatorList.Count; b++) if (charDesignatorList[b].Equals(newDesignator)) { bodyDesignatorList[b].Add(template.bodyRows[a]); match = true; break; }
                    if (match){ match = false; continue; }
                    bodyDesignatorList.Add(new List<List<LoadTemplate.myCell>>());
                    bodyDesignatorList[bodyDesignatorList.Count - 1].Add(template.bodyRows[a]);
                    charDesignatorList.Add(newDesignator);
                }
                else bodyDesignatorList[bodyDesignatorList.Count - 1].Add(template.bodyRows[a]);
            }

            int count = 1;
            for (int a = 0; a < bodyDesignatorList.Count; a++) // loop through all designator lists
            {
                sorted.Add(new List<List<LoadTemplate.myCell>>());
                sorted[sorted.Count - 1].Add(bodyDesignatorList[a][0]);
                for (int b = 1; b < bodyDesignatorList[a].Count; b++) // loop through the rows of a specific designator list
                {
                    bodyDesignatorList[a][b][quantityIndex].info = "";
                    matched = false;
                    if (bodyDesignatorList[a][b][valueIndex].info == bodyDesignatorList[a][b - 1][valueIndex].info)
                        if (bodyDesignatorList[a][b][packageIndex].info == bodyDesignatorList[a][b - 1][packageIndex].info) matched = true; // make sure that the two components are identical
                        else matched = false;
                    else matched = false;

                    if (matched)
                    {
                        sorted[sorted.Count - 1].Add(bodyDesignatorList[a][b]); //if both components are the same then add it to the component list
                        count++;
                    }
                    else
                    {
                        sorted[sorted.Count - 1][0][quantityIndex].info = count.ToString();
                        count = 1;
                        sorted.Add(new List<List<LoadTemplate.myCell>>());
                        sorted[sorted.Count - 1].Add(bodyDesignatorList[a][b]);

                    }
                }
            }


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

