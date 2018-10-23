﻿using System;
using System.Reflection;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;


namespace EBOMgui
{
    public partial class mainFrameScreen : Form
    {


        public bool lbMouseDown = false, dgvMouseDown = false;
        int mouseDownIndex = 0;

        delegate void dgetpMainFrame(Action job);
        public MainFrame mainFrame1;

        int currentCellRow;
        int currentCellColumn;


        public mainFrameScreen()
        {
            mainFrame1 = new MainFrame(this);
            InitializeComponent();
            createDataTable();
        }

        public void getScreen(Action job) // set the gui console to enabled depending on some conditions
        {
            try
            {
                if (this.dgvEBOM.InvokeRequired)
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

        private void createDataTable()
        {
            dgvEBOM.ColumnCount = 50;
            string[] row = new string[dgvEBOM.ColumnCount];
            for (int a = 0; a < dgvEBOM.ColumnCount; a++) //populate string[] row with all values of 1 row to fill all columns
                row[a] = (a * 1 + 1).ToString();
            for (int a = 0; a < dgvEBOM.ColumnCount; a++) // create all columns
                dgvEBOM.Columns[a].Name = (a+1).ToString();
            for (int a = 0; a < dgvEBOM.ColumnCount; a++) // create all rows
                dgvEBOM.Rows.Add(row);
            for (int a = 0; a < dgvEBOM.ColumnCount; a++) // populate all rows
                dgvEBOM.Rows[a].HeaderCell.Value = (a+1).ToString();

            dgvEBOM.DoubleBuffered(true);

        }
        
        public void changeFocusLB()
        {
            //Action myACtion = () =>
            //{
            //    lbAttributes.Focus();
            //    writeToConsole("LB focus");
            //};
            //getScreen(myACtion);
        }
        public void changeFocusDGV()
        {
            //Action myACtion = () =>
            //{
            //    dgvEBOM.Focus();
            //    writeToConsole("DGV focus");
            //};
            //getScreen(myACtion);
        }

        public void writeToConsole(string text)
        {
            //Action myACtion = () =>
            //{
            //    rtbConsole.AppendText(text + "\n");
            //};
            //getScreen(myACtion);
        }


        

        private void lbAttributes_DragDrop(object sender, DragEventArgs e)
        {
            //rtbConsole.AppendText(lbAttributes.SelectedItem.ToString() + "\n");
           // rtbConsole.AppendText("smokes" + "\n");

        }

        
       
        

     
     

        private void rtbConsole_TextChanged(object sender, EventArgs e)
        {
            rtbConsole.HideSelection = false;
            rtbConsole.SelectionStart = rtbConsole.Text.Length;
            rtbConsole.ScrollToCaret();
        }

        private void lbAttributes_MouseLeave(object sender, EventArgs e)
        {
            //dgvEBOM.Focus();
            rtbConsole.AppendText("lb mouse leave" + "\n");
        }
        private void lbAttributes_MouseDown(object sender, MouseEventArgs e)
        {
            Control control = (Control)sender;
            if (control.Capture)
            {
                control.Capture = false; // this prevents the control the mouse is hovering over from capturing all mouse events so other controls cant
            }
            if (e.Button == MouseButtons.Right)
            {
                lbMouseDown = true;
                rtbConsole.AppendText("lb mouse right down" + "\n");
                dgvEBOM.ClearSelection();
                int index = this.lbAttributes.IndexFromPoint(e.Location);
                if (index != ListBox.NoMatches)
                {
                    lbAttributes.SelectedIndex = index;
                }
                //Thread run = new Thread(delegate ()
                //{
                //    if (mainFrame1.determineDrag())
                //        changeFocusDGV();
                //});
                //run.Name = "drag";
                //run.Start();
                dgvEBOM.Capture = true;
            }
            else
                //{
                //    dgvEBOM.Focus();
                //    //Thread run = new Thread(delegate ()
                //    //{
                //    //    if (mainFrame1.determineDrag())
                //    //        changeFocusDGV();
                //    //});
                //    //run.Name = "drag";
                //    //run.Start();
                rtbConsole.AppendText("lb mouse down" + "\n");
                //}
                dgvEBOM.Focus();


            //rtbConsole.AppendText("smokes" + "\n");

            //changeFocusDGV();

            //mouseDownIndex = lbAttributes.SelectedIndex;
        }

        private void dgvEBOM_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            lbAttributes.ClearSelected();
            rtbConsole.AppendText("dgv mouse down" + "\n");

            //Thread.Sleep(1000);
            //dgvEBOM.CurrentCell.Selected = false;
            //dgvEBOM.ClearSelection();
            lbMouseDown = true;
            if (e.Button == MouseButtons.Right)
                dgvEBOM.CurrentCell = dgvEBOM.Rows[e.RowIndex].Cells[e.ColumnIndex];


        }


        private void dgvEBOM_MouseUp(object sender, MouseEventArgs e)
        {
            if (lbMouseDown)
            {
                if (lbAttributes.SelectedItem != null)
                {
                    //if (currentCellColumn >= 0 && currentCellRow >=0)
                    //    dgvEBOM[currentCellColumn, currentCellRow].Value = lbAttributes.SelectedItem.ToString();
                }
            }
            lbMouseDown = false;
            rtbConsole.AppendText("dgv mouse up" + "\n");
        }
        private void dgvEBOM_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            rtbConsole.AppendText("enter " + e.RowIndex + " " + e.ColumnIndex + "\n");
            currentCellRow = e.RowIndex;
            currentCellColumn = e.ColumnIndex;
            if (lbMouseDown)
            {
                if (lbAttributes.SelectedItem == null)

                mainFrame1.paintCells(mainFrame1.headerType, e.RowIndex, e.ColumnIndex, lbAttributes.SelectedItem.ToString());
            }
        }

        private void dgvEBOM_MouseDown(object sender, MouseEventArgs e)
        {



        }
        
        public void dgvEBOM_ChangeColor(Color color, int row, int column)
        {
            Action myACtion = () =>
            {
                dgvEBOM[column, row].Style.ForeColor = color;
            };
            getScreen(myACtion);
        }
        public void dgvEBOM_setText(string text, int row, int column)
        {
            Action myACtion = () =>
            {
                dgvEBOM[column, row].Value = text;
            };
            getScreen(myACtion);
        }

        private void dgvEBOM_SelectionChanged(object sender, EventArgs e)
        {
            //rtbConsole.AppendText("selection changed" + "\n");
            //mainFrame1.dgvSelectionChanged();
        }

        public void dgvEBOMClearSelected()
        {
            dgvEBOM.ClearSelection();
        }

        private void bOpenFile_Click(object sender, EventArgs e)
        {
            mainFrame1.getFile();
        }
        public void addElementToAttributeListBox(string element)
        {
            lbAttributes.Items.Add(element);
        }

        private void rbHeaderColumn_CheckedChanged(object sender, EventArgs e)
        {
            if (rbHeaderColumn.Checked) mainFrame1.headerType = true;
            else mainFrame1.headerType = false;
        }

       
    }

    public static class ExtensionMethods
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }
    }
}
