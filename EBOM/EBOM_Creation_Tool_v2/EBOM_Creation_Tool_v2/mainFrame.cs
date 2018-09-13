using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;



namespace EBOM_Creation_Tool_v2
{
    public class mainFrame
    {
        public string filename;
        mainFrameScreen mainFrameScreen1;
        string[] fileNames;


        public static bool end = false;

        public mainFrame(mainFrameScreen m, string[] a)
        { 
            mainFrameScreen1 = m;
            fileNames = a;
        }
        public void start() // main order of the program.
        {
            Thread runParser = new Thread(delegate ()
            {
                try
                {
                    excelSection template = new excelSection();
                    sort sort1 = new sort();
                    countParts countParts1 = new countParts();
                    excelFileHandler excelFileHandler1 = new excelFileHandler(this, ref template, ref sort1, ref countParts1);
                    xmlFileHandler xmlFileHandler1 = new xmlFileHandler(this, template, filename);
                    sort1.start(this, sort1, xmlFileHandler1.componentAttributes, template.Htext, countParts1);
                    EBOMexcelFile eBOMexcelFile1 = new EBOMexcelFile(excelFileHandler1, xmlFileHandler1.componentAttributes, template, filename, this, xmlFileHandler1.totalPartCount);
                    writeToConsole("Complete");
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            });
            runParser.Name = "CreateEBOM";
            runParser.Start();

        }
        public void writeToConsole(string myMessage)
        {
            mainFrameScreen1.writeToConsole(myMessage);    
        }
        public void chooseSourceXML()
        {
            if (fileNames.Length > 0)
            {
                filename = fileNames[0];
                if (verifyXML(filename))
                    mainFrameScreen1.updateFileNameLabel(filename);
                else
                    System.Windows.Forms.Application.Exit();
            }
            else
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    filename = openFileDialog1.FileName;
                    if (verifyXML(filename))
                    {
                        mainFrameScreen1.updateFileNameLabel(filename);
                        mainFrameScreen1.enableStartButton(true);
                    }
                    else mainFrameScreen1.enableStartButton(false);
                }
            }            
        }
        public bool verifyXML(string file)
        {
            string[] fileTemp = file.Split('.');
            if (fileTemp.Length > 1)
            {
                if (!fileTemp[fileTemp.Length - 1].ToLower().Equals("xml"))
                {
                    MessageBox.Show("File selected is not a .XML");
                    //System.Windows.Forms.Application.Exit(); 
                    return false;
                }
                else return true;
            }
            else
            {
                MessageBox.Show("File selected is not a .XML");
                //System.Windows.Forms.Application.Exit();
                return false;
            }
        }
    }
    
}
