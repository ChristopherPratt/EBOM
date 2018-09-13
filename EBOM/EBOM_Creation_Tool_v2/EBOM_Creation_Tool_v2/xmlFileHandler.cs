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
    class xmlFileHandler
    {
        public int[] TBindex; // title block index
        public int[] Hindex; // header index
        public int totalPartCount;
        public List<string> titleBlockInfo;
        public List<List<string>> componentAttributes;
        public string exportFileName;


        public xmlFileHandler(mainFrame mainFrame1, excelSection excelSection1, string filePath)
        {
            try
            {                
                //instantiate all the necessary lists and objects for handling the XML
                componentAttributes = new List<List<string>>();
                titleBlockInfo = new List<string>();
                TBindex = new int[excelSection1.TBtext.Count];
                Hindex = new int[excelSection1.Htext.Count];
                xmlFileParser xmlFileParser1 = new xmlFileParser(); 
                XmlDocument xmlRead = new XmlDocument();


                openXML(xmlRead, filePath); // load the xml file into the xmlRead object
                XmlNodeList partAttributesNodeList = xmlRead.SelectNodes("GRID")[0].SelectNodes("COLUMNS")[0].SelectNodes("COLUMN");
                XmlNodeList componentNodeList = xmlRead.SelectNodes("GRID")[0].SelectNodes("ROWS")[0].SelectNodes("ROW");
                readHeaderAttributes(partAttributesNodeList, xmlFileParser1, excelSection1.TBtext, ref TBindex); 
                readHeaderAttributes(partAttributesNodeList, xmlFileParser1, excelSection1.Htext, ref Hindex); 
                readComponents(componentNodeList, xmlFileParser1, TBindex, Hindex, ref titleBlockInfo, ref componentAttributes, ref totalPartCount);
                setupExcel(filePath);
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

        // function responsible for looping through all nodes of selected nodelist and parsing nodes to look for index of specific attribute name in component.
        public void readHeaderAttributes(XmlNodeList nodeList, xmlFileParser xmlFileParser2, List<string> cellNameList, ref int[] index)
        {
            if (nodeList.Count == 0)
            {
                throw new Exception("No nodes found in xml file when looking for headers");
            }
            foreach (XmlNode node in nodeList) //for every header loaded
            {
                xmlFileParser2.getColumnHeaderIndex(node, cellNameList, ref index);
            }
        }
        public void readComponents(XmlNodeList nodeList, xmlFileParser xmlFileParser2, int[] TBindex, int[] Hindex, ref List<string> titleBlockInfo1, ref List<List<string>> componentsAttributes1, ref int totalPartCount1)
        {
            //Read the parts
            // XmlNodeList nodeList;
            if (nodeList.Count == 0)
            {
                throw new Exception("No nodes found in xml file when looking for headers");
            }
            xmlFileParser2.getTitleBlockInfo(nodeList[0], TBindex, ref titleBlockInfo1);
            totalPartCount1 =  xmlFileParser2.getComponentInfo(nodeList, Hindex, ref componentsAttributes1);
        }
        private void setupExcel(string xmlFile)
        {
            exportFileName = System.IO.Path.ChangeExtension(xmlFile, null) + ".xlsx";
        }
    }
}
