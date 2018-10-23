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

            bStart.Enabled = false;
            if (args.Length > 0)
            {                    
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
            Action myACtion = () =>
            {
                bStart.Enabled = set;
            };
            getScreen(myACtion);
        }
        public void enableSourceButton(bool set)
        {
            Action myACtion = () =>
            {
                bChooseSource.Enabled = set;
            };
            getScreen(myACtion);
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
            enableStartButton(false);
            enableSourceButton(false);
        }

        public void writeToConsole(string text)
        {
            Action myACtion = () =>
            {
                rtbConsole.AppendText(text + "\n");
            };
            getScreen(myACtion);
        }

        public void changeFormColor(bool result)
        {
            Action myACtion = () =>
            {
                if (result)
                {
                    ActiveForm.BackColor = Color.Green;
                }
                else
                {
                    ActiveForm.BackColor = Color.Red;
                }
            };
            getScreen(myACtion);           
        }       

        private void mainFrameScreen_FormClosing(object sender, FormClosingEventArgs e)
        {
            mainFrame1.terminateThread();
        }
    }
}
