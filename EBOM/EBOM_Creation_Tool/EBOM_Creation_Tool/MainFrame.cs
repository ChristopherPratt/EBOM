using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;



namespace EBOMCreationTool
{
    
    public partial class MainFrame : Form
    {
        DataSet dataSet;
        private List<string> MergedRowsInFirstColumn = new List<string>();

        public MainFrame()
        {
            InitializeComponent();
            //dataGrid.Paint += new PaintEventHandler(dataGrid_Paint);
        }
        //private void createDataTable()
        //{
        //    dataGrid.ColumnCount = 26;
        //    string[] row = new string[dataGrid.ColumnCount];
        //    for (int a = 0; a < dataGrid.ColumnCount; a++)
        //        row[a] = (a * 11111).ToString();
        //    for (int a = 0; a < dataGrid.ColumnCount; a++)
        //        dataGrid.Columns[a].Name = a.ToString();
        //    for (int a = 0; a < dataGrid.ColumnCount; a++)
        //        dataGrid.Rows.Add(row);
        //    for (int a = 0; a < dataGrid.ColumnCount; a++)
        //        dataGrid.Rows[a].HeaderCell.Value = a.ToString() ;
           
        //    dataGrid.DoubleBuffered(true);
            
        //}

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            Console.WriteLine(e.ColumnIndex + " " + e.RowIndex);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //createDataTable();
            //dataGrid.Dock = DockStyle.Top;
            //this.dataGrid.Paint += new PaintEventHandler(dataGrid_Paint);
            //this.dataGridView1.Scroll += new ScrollEventHandler(dataGridView1_Scroll);

            //this.dataGridView1.CellPainting +=

            //    new DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);
        }
        
       

        //private void MergeCells(int RowId1, int RowId2, int Column, bool isSelected)
        //{
        //    int rectHeight = 0;
        //    string MergedRows = String.Empty;
        //    int rowBeginning;
        //    int rowEnd;
            
        //    Graphics g = dataGrid.CreateGraphics();
        //    Pen gridPen = new Pen(dataGrid.GridColor);

        //    if (RowId1 > RowId2) { rowBeginning = RowId2; rowEnd = RowId1; }
        //    else                 { rowBeginning = RowId1; rowEnd = RowId2; }

        //    //Cells Rectangles
        //    Rectangle CellRectangle1 = dataGrid.GetCellDisplayRectangle(Column, rowBeginning, true);
        //    Rectangle CellRectangle2 = dataGrid.GetCellDisplayRectangle(Column, rowEnd, true);

            
            

        //    for (int i = rowBeginning; i <= rowEnd; i++)
        //    {
        //        rectHeight += dataGrid.GetCellDisplayRectangle(Column, i, false).Height;
        //    }

        //    Rectangle newCell = new Rectangle(CellRectangle1.X, CellRectangle1.Y, CellRectangle1.Width, rectHeight);

        //    g.FillRectangle(new SolidBrush(isSelected ? dataGrid.DefaultCellStyle.SelectionBackColor : dataGrid.DefaultCellStyle.BackColor), newCell);

        //    g.DrawRectangle(gridPen, newCell);

        //    g.DrawString(dataGrid.Rows[rowBeginning].Cells[Column].Value.ToString(), dataGrid.DefaultCellStyle.Font, new SolidBrush(isSelected ? dataGrid.DefaultCellStyle.SelectionForeColor : dataGrid.DefaultCellStyle.ForeColor), newCell.X + newCell.Width / 3, newCell.Y - 6 + newCell.Height / 2);
        //}

        
        private void bChooseXML_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbXML.Text = openFileDialog1.FileName;
            }
        }

        private void bChooseTemplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbTemplate.Text = openFileDialog1.FileName;
            }
        }

        private void ChooseExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel files (*.xlsx)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbExport.Text = saveFileDialog1.FileName + ".xlsx";
            }
        }
        private void bStart_Click(object sender, EventArgs e)
        {
            CreateExcelFile c;
            LoadXML l;
            LoadTemplate t;



            t = new LoadTemplate(tbTemplate.Text);
            l = new LoadXML(t, tbXML.Text);
            c = new CreateExcelFile(l, t, tbExport.Text);
        }

        private void tbExport_TextChanged(object sender, EventArgs e)
        {

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
