using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EBOM_Creation_Tool_v2
{
    public partial class mainFrameScreen : Form
    {
        mainFrame mainFrame1;
        delegate void dgetpMainFrame(Action job);

        public mainFrameScreen(string[] args)
        {
            mainFrame1 = new mainFrame(this, args);
            InitializeComponent();

            if (args.Length > 0)
            {
                bStart.Enabled = false;
                bChooseSource.Enabled = false;
            }
        }

        public void getScreen(Action job) // set the gui console to enabled depending on some conditions
        {
            try
            {
                if (this.rtbConsole.InvokeRequired)
                {
                    dgetpMainFrame d = new dgetpMainFrame(getScreen);
                    this.Invoke(d, new object[] { job });
                }
                else
                {
                    job();
                }
            }
            catch { };
        }

        public void enableStartButton(bool set)
        {
            bStart.Enabled = set;
        }

        private void bChooseSource_Click(object sender, EventArgs e)
        {
            mainFrame1.chooseSourceXML();
        }

        public void updateFileNameLabel(string filename)
        {
            Action myACtion = () =>
            {
                lblFileName.Text = filename;
            };
            getScreen(myACtion);            
        }

        private void bStart_Click(object sender, EventArgs e)
        {
            mainFrame1.start();
        }

        public void writeToConsole(string text)
        {
            Action myACtion = () =>
            {
                rtbConsole.AppendText(text + "\n");
            };
            getScreen(myACtion);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                test();
                MessageBox.Show("test");
            }
            catch
            {
                MessageBox.Show(".XML file is incorrect");
            }
            
        }
        private void test()
        {
            try
            {
                test2();
                MessageBox.Show("test2");
            }
            catch
            {
                MessageBox.Show("test2 catch");
                throw new System.InvalidOperationException("The Index of an attribute in the .XML is greater than the total amount of attributes");


            }

        }
        private void test2()
        {
            test3();
            MessageBox.Show("test3");
        }
        private void test3()
        {
            throw new System.InvalidOperationException("The Index of an attribute in the .XML is greater than the total amount of attributes"); 

        }
    }
}
