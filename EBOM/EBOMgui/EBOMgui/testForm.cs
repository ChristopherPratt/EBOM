using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;


namespace EBOMgui
{
    public partial class testForm : Form
    {
        delegate void dgetpMainFrame(Action job);

        public testForm()
        {
            InitializeComponent();
            //listBox1.AllowDrop = true;
        }
        public void getScreen(Action job) // set the gui console to enabled depending on some conditions
        {
            try
            {
                if (this.dataGridView1.InvokeRequired)
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
        private void dataGridView1_MouseEnter(object sender, EventArgs e)
        {
            richTextBox1.AppendText("dtg Entered" + "\n");
        }

        private void dataGridView1_MouseLeave(object sender, EventArgs e)
        {
            richTextBox1.AppendText("dtg left" + "\n");
        }

        private void listBox1_MouseEnter(object sender, EventArgs e)
        {
            richTextBox1.AppendText("lb entered" + "\n");
        }

        private void listBox1_MouseLeave(object sender, EventArgs e)
        {
            richTextBox1.AppendText("lb left" + "\n");
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_MouseDown(object sender, MouseEventArgs e)
        {
            Control control = (Control)sender;
            if (control.Capture)
            {
                control.Capture = false;
            }
            //if (e.Button == MouseButtons.Right)
            //{
            //    int index = this.listBox1.IndexFromPoint(e.Location);
            //    if (index != ListBox.NoMatches)
            //    {
            //        listBox1.SelectedIndex = index;
            //    }
            //}
            //richTextBox1.AppendText("lb mouse down" + "\n");
            //Thread run = new Thread(delegate ()
            //{
            //    Action myACtion = () =>
            //    {
            //        //dataGridView1.Focus();
            //    };
            //    getScreen(myACtion);
            //});
            //run.Name = "drag";
            //run.Start();
        }

        private void listBox1_MouseUp(object sender, MouseEventArgs e)
        {
            richTextBox1.AppendText("mouse up" + "\n");
        }

        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            richTextBox1.AppendText("dgv mouse down" + "\n");
        }

        private void dataGridView1_MouseUp(object sender, MouseEventArgs e)
        {
            richTextBox1.AppendText("dgv mouse up" + "\n");
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.HideSelection = false;
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
        }

        private void listBox1_MouseMove(object sender, MouseEventArgs e)
        {
            //richTextBox1.AppendText("lb mouse move" + "\n");

        }

        private void listBox1_DragLeave(object sender, EventArgs e)
        {
            richTextBox1.AppendText("lb drag leave" + "\n");
        }
    }
}
