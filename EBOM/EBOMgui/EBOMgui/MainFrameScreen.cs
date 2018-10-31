using System;
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
                dgvEBOM.Columns[a].Name = (a + 1).ToString();
            for (int a = 0; a < dgvEBOM.ColumnCount; a++) // create all rows
            {
                dgvEBOM.Rows.Add(row);
                for (int b = 0; b < row.Length; b++)
                {
                    dgvEBOM[b, a].Value = row[b];
                }
            }
            for (int a = 0; a < dgvEBOM.ColumnCount; a++) // populate all rows
                dgvEBOM.Rows[a].HeaderCell.Value = (a + 1).ToString();

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
        public void writeLineToConsole(string text)
        {
            Action myACtion = () =>
            {
                rtbConsole.AppendText(text + "\n");
            };
            getScreen(myACtion);
        }







        private void rtbConsole_TextChanged(object sender, EventArgs e)
        {
            rtbConsole.HideSelection = false;
            rtbConsole.SelectionStart = rtbConsole.Text.Length;
            rtbConsole.ScrollToCaret();
        }
        ////////////////////////////////////////// lbAttributes Events /////////////////////////////////////////////
        public string getlbAttributeSelectedString()
        {
            return lbAttributes.SelectedItem.ToString();
        }
        public void clearlbAttribute()
        {
            lbAttributes.ClearSelected();
        }
        private void lbAttributes_MouseLeave(object sender, EventArgs e)
        {
            //dgvEBOM.Focus();
            rtbConsole.AppendText("lb mouse leave" + "\n");
        }
        private void lbAttributes_MouseUp(object sender, MouseEventArgs e)
        {
            lbMouseDown = false;
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
                rtbConsole.AppendText("lb mouse right down" + "\n");
                dgvEBOM.ClearSelection();
                int index = this.lbAttributes.IndexFromPoint(e.Location);
                if (index != ListBox.NoMatches)
                {
                    lbAttributes.SelectedIndex = index;
                }
                mainFrame1.lbMouseDownEvent(e, index);                
                dgvEBOM.Capture = true;
            }
            
                //{
                //    dgvEBOM.Focus();
                //    //Thread run = new Thread(delegate ()
                //    //{
                //    //    if (mainFrame1.determineDrag())
                //    //        changeFocusDGV();
                //    //});
                //    //run.Name = "drag";
                //    //run.Start();
                //rtbConsole.AppendText("lb mouse down" + "\n");
                //}
                //dgvEBOM.Focus();


            //rtbConsole.AppendText("smokes" + "\n");

            //changeFocusDGV();

            //mouseDownIndex = lbAttributes.SelectedIndex;
        }

        ////////////////////////////////////////// lbAttributes Events /////////////////////////////////////////////

        //////////////////////////////////////////// dgvEBOM Events ////////////////////////////////////////////////

        private void dgvEBOM_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                //dgvEBOM.CurrentCell = dgvEBOM.Rows[e.RowIndex].Cells[e.ColumnIndex];
                mainFrame1.dgvMouseDownEvent(e);                    
            }
            rtbConsole.AppendText("dgv mouse down" + "\n");
        }
        private void dgvEBOM_MouseUp(object sender, MouseEventArgs e)
        {
            //if (lbMouseDown)
            //{
            //    lbMouseDown = false;
            //    if (lbAttributes.SelectedItem != null)
            //    {
            //        //if (currentCellColumn >= 0 && currentCellRow >=0)
            //        //    dgvEBOM[currentCellColumn, currentCellRow].Value = lbAttributes.SelectedItem.ToString();
            //    }
            //}
            //else if (dgvMouseDown)
            //{
            //    dgvMouseDown = false;
            //}
            mainFrame1.dgvMouseUpEvent(e);
            rtbConsole.AppendText("dgv mouse up" + "\n");
        }
        private void dgvEBOM_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.ColumnIndex > 1 && e.RowIndex > 1)
            //    dgvEBOM.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            mainFrame1.dgvEBOM_CellMouseEnter(e.RowIndex, e.ColumnIndex);
            rtbConsole.AppendText("enter " + e.RowIndex + " " + e.ColumnIndex + "\n");
            
            //currentCellRow = e.RowIndex;
            //currentCellColumn = e.ColumnIndex;

        }


        
        public void dgvEBOM_ChangeColor(Color color, int row, int column)
        {
            Action myACtion = () =>
            {
                dgvEBOM[column, row].Style.BackColor = color;
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

        public void getdgvEBOMCellInfo(ref string text, ref Color color, int row, int column)
        {
            
            //text = dgvEBOM[column, row].Value.ToString();
            color = dgvEBOM.Rows[row].Cells[column].Style.BackColor;
            text = dgvEBOM.Rows[row].Cells[column].Value.ToString();
        }

        ////////////////////////////////////////// dgvEBOM Events /////////////////////////////////////////////

        private void bOpenFile_Click(object sender, EventArgs e)
        {
            mainFrame1.getFile();
        }
        public void addElementToAttributeListBox(string element)
        {
            lbAttributes.Items.Add(element);
        }

        private void lbAttributes_SelectedIndexChanged(object sender, EventArgs e)
        {

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
