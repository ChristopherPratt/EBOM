using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;



namespace EBOMgui
{
    public class MainFrame
    {
        List<int> attributeRows, attributeColumn;
        List<string> attributeNames;
        List<List<string>> componentContent;

        bool rightMouseDown = false, leftMouseDown = false;

        public bool headerType = true;
        xmlFileHandler xmlFileHandler1;
        mainFrameScreen mainFrameScreen1;
        public MainFrame(mainFrameScreen m)
        {
            mainFrameScreen1 = m;
        }
        public void insertAttribute()
        {

        }
        public void getFile()
        {
            string filename = "";
            OpenFileDialog openFileDialog1 = new OpenFileDialog(); // opens file explorer menu
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialog1.FileName; // saves chosen file name to filename variable
                if (verifyXML(filename)) // verifies the file extension is .xml
                {
                    xmlFileHandler1 = new xmlFileHandler();
                    xmlFileHandler1.run(this, filename); // opens the xml file and extracts all attributes, indexes and component info.
                    foreach (string element in xmlFileHandler1.attributeNames)
                        mainFrameScreen1.addElementToAttributeListBox(element); // loop through all attribute names and input them into the attribute listbox
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

        public void paintCells(bool header, int row, int column, string text)
        {
            mainFrameScreen1.dgvEBOM_ChangeColor(Color.Green, row, column);
            mainFrameScreen1.dgvEBOM_setText(text, row, column);
            if (header)
            {
                for (int a = row; a < row + 10; a++)
                {
                    mainFrameScreen1.dgvEBOM_ChangeColor(Color.Yellow, a, column);
                }
            }
            else
            {
                mainFrameScreen1.dgvEBOM_ChangeColor(Color.Orange, row, column+1);
            }
        }

        public void populateAttributeListBox(List<string> attributes)
        {

        }

        public bool determineDrag()
        {
            int count = 0;
            while (mainFrameScreen1.lbMouseDown)
            {
                Thread.Sleep(75);
                if (count++ == 2) return true;
            }
            return false;
        }

        public void mouseDownEvent(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right) rightMouseDown = true;
            else if (e.Button == MouseButtons.Left) leftMouseDown = true;
        }

        public void dgvSelectionChanged()
        {
            //if (rightMouseDown) mainFrameScreen1.dgvEBOMClearSelected();
        }
        public void writeToConsole(string text)
        {
            mainFrameScreen1.writeToConsole(text);
        }
    }
}
