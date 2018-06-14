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
        List<List<List<string>>> sorted;
        XmlDocument xmlRead;
        XmlNodeList nodeList;
        string apcbField = "";

        public LoadXML()
        {
            properties = new List<string>();
            propertyValues = new List<string>();
            headers = new List<string>();
            unsorted = new List<List<string>>();
            sorted = new List<List<List<string>>>();
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
                string nodeAttribute = node.Attributes.Item(0).Value;
                if (
                        nodeAttribute.ToLower().Contains("prod") ||
                        nodeAttribute.ToLower().Contains("edit") ||
                        nodeAttribute.ToLower().Contains("addi") || //while not strictly a part of apcb when set, is still included in the export
                        nodeAttribute.ToLower().Contains("appr") ||
                        nodeAttribute.ToLower().Contains("modi") ||
                        nodeAttribute.ToLower().Contains("titl") ||
                        nodeAttribute.ToLower().Contains("cust") ||
                        nodeAttribute.ToLower().Contains("rele")
                  )
                {
                    properties.Add(nodeAttribute); //is a property value


                }
                else
                {
                    if (nodeAttribute.ToLower().Contains("desi"))
                    {
                        apcbField = nodeAttribute;
                    }
                    headers.Add(nodeAttribute);
                }
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

            //Use the first node to read the property info

            //Each node is a part, add the part to the list
            bool propertiesLoaded = false;
            foreach (XmlNode node in nodeList) //for every part
            {
                if (!string.IsNullOrEmpty(apcbField)) //nothing will be loaded if apcb is empty!
                {
                    if (!propertiesLoaded && node.Attributes.GetNamedItem(apcbField).Value == "APCB")
                    {
                        //ignore exceptions in this area, as they are from non existent property values-
                        //not an issue, the field will just be blank
                        foreach (string property in properties)
                        {

                            propertyValues.Add(node.Attributes.GetNamedItem(property).Value);
                        }
                    }
                }
                propertiesLoaded = true;
                List<string> tempAttributes = new List<string>();
                //Add each part's attributes to the part object

                for (int i = 0; i < headers.Count; i++) //for every header attribute, including the properties
                {

                    tempAttributes.Add(node.Attributes.GetNamedItem(headers[i]).Value);
                }
                unsorted.Add(tempAttributes);
                //tempAttributes.Clear(); //Memory management
            }
            Console.WriteLine("Unsorted parts = " + unsorted.Count);
        }
        public void sort()
        {
            int quantityIndex = 0;
            int valueIndex = 0;
            int packageIndex = 0;
            bool matched = false;
            for (int a = 0; a < headers.Count; a++)
            {
                if (headers[a].ToLower().Contains("quan") || headers[a].ToLower().Contains("qty")) quantityIndex = a;
                if (headers[a].ToLower().Contains("val")) valueIndex = a;
                if (headers[a].ToLower().Contains("pack")) packageIndex = a;


            }
            List<List<string>> temp = new List<List<string>>();
            temp.Add(unsorted[0]);
            for (int a = 1; a < unsorted.Count; a++)
            {
                unsorted[a][quantityIndex] = "";
                matched = false;
                if (unsorted[a][valueIndex] == unsorted[a - 1][valueIndex])
                    if (unsorted[a][packageIndex] == unsorted[a - 1][packageIndex]) matched = true; // make sure that the two components are identical
                    else matched = false;
                else matched = false;

                if (matched)
                    temp.Add(unsorted[a]); //if both components are the same then add it to the component list
                else
                {
                    temp[0][quantityIndex] = temp.Count.ToString(); // if both components aren't the same then change the quantity of the top component to the amount of all components
                    sorted.Add(temp);
                    temp = new List<List<string>> { unsorted[a] };
                }
            }

        }
    }  
}

