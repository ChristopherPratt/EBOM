using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;




namespace EBOM_Creation_Tool_v2
{
    public class mainFrame
    {
        public string filename;
        mainFrameScreen mainFrameScreen1;
        string[] fileNames;
        Thread runParser, run;

        excelFileHandler excelFileHandler1;


        public static bool end = false;

        public mainFrame(mainFrameScreen m, string[] a)
        {
            mainFrameScreen1 = m;
            fileNames = a;
            if (fileNames.Length > 0)
            {
                run = new Thread(delegate ()
                {
                    foreach (string file in fileNames)
                    {
                        filename = file;
                        start();
                        runParser.Join();
                    }
                    Thread.Sleep(1500);
                    Application.Exit();
                });
                run.Name = "run";
                run.Start();
            }
        }
        public void start() // main order of the program.
        {
            runParser = new Thread(delegate ()
            {                
                try
                {                    
                    try
                    {
                        excelSection template = new excelSection(); //initiating necessary objects
                        sort sort1 = new sort();
                        countParts countParts1 = new countParts();
                        excelFileHandler1 = new excelFileHandler(this); 
                        excelFileHandler1.run(ref template, ref sort1, ref countParts1); // read template excel file
                        xmlFileHandler xmlFileHandler1 = new xmlFileHandler(this, template, filename); // read source xml file
                        sort1.start(this, sort1, xmlFileHandler1.componentAttributes, template.Htext, template.HcolumnIndex, countParts1); // sort xml file info
                        EBOMexcelFile eBOMexcelFile1 = new EBOMexcelFile(excelFileHandler1, sort1.sorted, template, xmlFileHandler1.exportFileName, this, xmlFileHandler1.totalPartCount, xmlFileHandler1.titleBlockInfo);// create new EBOM from xml file using template file info
                        writeToConsole("Complete");
                        mainFrameScreen1.enableStartButton(true);
                        mainFrameScreen1.enableSourceButton(true);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.ToString());
                    }
                    finally
                    {
                        try
                        {
                            Marshal.FinalReleaseComObject(excelFileHandler1.xlWorkSheet);
                            //template.xlWorkBook.Close(false, misValue, misValue); 
                            excelFileHandler1.xlWorkBook.Close(SaveChanges: false);
                            //template.xlWorkBook.Close();
                            Marshal.FinalReleaseComObject(excelFileHandler1.xlWorkBook);
                            excelFileHandler1.xlWorkBooks.Close();
                            Marshal.FinalReleaseComObject(excelFileHandler1.xlWorkBooks);
                            excelFileHandler1.xlApp.Quit();
                            Marshal.FinalReleaseComObject(excelFileHandler1.xlApp); // excel objects don't releast comObjects to excel so you have to force it
                        }
                        catch {}                        
                    }
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
        public void handleExcelPorts()
        {
            Marshal.FinalReleaseComObject(excelFileHandler1.xlWorkSheet);
            //template.xlWorkBook.Close(false, misValue, misValue); 
            excelFileHandler1.xlWorkBook.Close(SaveChanges: false);
            //template.xlWorkBook.Close();
            Marshal.FinalReleaseComObject(excelFileHandler1.xlWorkBook);
            excelFileHandler1.xlWorkBooks.Close();
            Marshal.FinalReleaseComObject(excelFileHandler1.xlWorkBooks);
            excelFileHandler1.xlApp.Quit();
            Marshal.FinalReleaseComObject(excelFileHandler1.xlApp); // excel objects don't releast comObjects to excel so you have to force it
        }

        public void missingParts(bool result)
        {
            mainFrameScreen1.changeFormColor(result);
        }
        public void terminateThread()
        {
            mainFrame.end = true;
            //runParser.Join();
            //Application.Exit();
        }
         
    }
    
}
