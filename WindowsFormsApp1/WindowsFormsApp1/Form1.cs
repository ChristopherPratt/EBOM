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



namespace WindowsFormsApp1
{
    
    public partial class Form1 : Form
    {
        DataSet dataSet;
        private List<string> MergedRowsInFirstColumn = new List<string>();

        public Form1()
        {
            InitializeComponent();
            dataGrid.Paint += new PaintEventHandler(dataGrid_Paint);
        }
        private void createDataTable()
        {
            dataGrid.ColumnCount = 26;
            string[] row = new string[dataGrid.ColumnCount];
            for (int a = 0; a < dataGrid.ColumnCount; a++)
                row[a] = (a * 11111).ToString();
            for (int a = 0; a < dataGrid.ColumnCount; a++)
                dataGrid.Columns[a].Name = a.ToString();
            for (int a = 0; a < dataGrid.ColumnCount; a++)
                dataGrid.Rows.Add(row);
            for (int a = 0; a < dataGrid.ColumnCount; a++)
                dataGrid.Rows[a].HeaderCell.Value = a.ToString() ;
           
            dataGrid.DoubleBuffered(true);
            
        }

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            Console.WriteLine(e.ColumnIndex + " " + e.RowIndex);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            createDataTable();
            dataGrid.Dock = DockStyle.Top;
            this.dataGrid.Paint += new PaintEventHandler(dataGrid_Paint);
            //this.dataGridView1.Scroll += new ScrollEventHandler(dataGridView1_Scroll);

            //this.dataGridView1.CellPainting +=

            //    new DataGridViewCellPaintingEventHandler(dataGridView1_CellPainting);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGrid.SelectedCells != null)
                if (dataGrid.SelectedCells.Count > 1)
                    MergeCells(Convert.ToInt32(dataGrid.SelectedCells[0].RowIndex), Convert.ToInt32(dataGrid.SelectedCells[dataGrid.SelectedCells.Count-1].RowIndex), Convert.ToInt32(dataGrid.SelectedCells[0].ColumnIndex), false);
        }
        private void dataGrid_Paint(object sender, PaintEventArgs e)
        {
            if (dataGrid.SelectedCells != null)
                if (dataGrid.SelectedCells.Count > 1)
                    MergeCells(Convert.ToInt32(dataGrid.SelectedCells[0].RowIndex), Convert.ToInt32(dataGrid.SelectedCells[dataGrid.SelectedCells.Count - 1].RowIndex), Convert.ToInt32(dataGrid.SelectedCells[0].ColumnIndex), true);

        }
        /*
        private void Merge()
        {
            int[] RowsToMerge = new int[2];
            RowsToMerge[0] = -1;

            //Merge first column at first
            for (int i = 0; i < dataSet.Tables["tbl_main"].Rows.Count - 1; i++)
            {
                if (dataSet.Tables["tbl_main"].Rows[i]["Manufacture"] == dataSet.Tables["tbl_main"].Rows[i + 1]["Manufacture"])
                {
                    if (RowsToMerge[0] == -1)
                    {
                        RowsToMerge[0] = i;
                        RowsToMerge[1] = i + 1;
                    }
                    else
                    {
                        RowsToMerge[1] = i + 1;
                    }
                }
                else
                {
                    MergeCells(RowsToMerge[0], RowsToMerge[1], dataGrid.Columns["Manufacture"].Index, isSelectedCell(RowsToMerge, dataGrid.Columns["Manufacture"].Index) ? true : false);
                    CollectMergedRowsInFirstColumn(RowsToMerge[0], RowsToMerge[1]);
                    RowsToMerge[0] = -1;
                }
                if (i == dataSet.Tables["tbl_main"].Rows.Count - 2)
                {
                    MergeCells(RowsToMerge[0], RowsToMerge[1], dataGrid.Columns["Manufacture"].Index, isSelectedCell(RowsToMerge, dataGrid.Columns["Manufacture"].Index) ? true : false);
                    CollectMergedRowsInFirstColumn(RowsToMerge[0], RowsToMerge[1]);
                    RowsToMerge[0] = -1;
                }
            }
            if (RowsToMerge[0] != -1)
            {
                MergeCells(RowsToMerge[0], RowsToMerge[1], dataGrid.Columns["Manufacture"].Index, isSelectedCell(RowsToMerge, dataGrid.Columns["Manufacture"].Index) ? true : false);
                RowsToMerge[0] = -1;
            }

            //merge all other columns
            for (int iColumn = 1; iColumn < dataSet.Tables["tbl_main"].Columns.Count - 1; iColumn++)
            {
                for (int iRow = 0; iRow < dataSet.Tables["tbl_main"].Rows.Count - 1; iRow++)
                {
                    if ((dataSet.Tables["tbl_main"].Rows[iRow][iColumn] == dataSet.Tables["tbl_main"].Rows[iRow + 1][iColumn]) &&
                         (isRowsHaveOneCellInFirstColumn(iRow, iRow + 1)))
                    {
                        if (RowsToMerge[0] == -1)
                        {
                            RowsToMerge[0] = iRow;
                            RowsToMerge[1] = iRow + 1;
                        }
                        else
                        {
                            RowsToMerge[1] = iRow + 1;
                        }
                    }
                    else
                    {
                        if (RowsToMerge[0] != -1)
                        {
                            MergeCells(RowsToMerge[0], RowsToMerge[1], iColumn, isSelectedCell(RowsToMerge, iColumn) ? true : false);
                            RowsToMerge[0] = -1;
                        }
                    }
                }
                if (RowsToMerge[0] != -1)
                {
                    MergeCells(RowsToMerge[0], RowsToMerge[1], iColumn, isSelectedCell(RowsToMerge, iColumn) ? true : false);
                    RowsToMerge[0] = -1;
                }
            }
        }
        */

        private void MergeCells(int RowId1, int RowId2, int Column, bool isSelected)
        {
            int rectHeight = 0;
            string MergedRows = String.Empty;
            int rowBeginning;
            int rowEnd;
            
            Graphics g = dataGrid.CreateGraphics();
            Pen gridPen = new Pen(dataGrid.GridColor);

            if (RowId1 > RowId2) { rowBeginning = RowId2; rowEnd = RowId1; }
            else                 { rowBeginning = RowId1; rowEnd = RowId2; }

            //Cells Rectangles
            Rectangle CellRectangle1 = dataGrid.GetCellDisplayRectangle(Column, rowBeginning, true);
            Rectangle CellRectangle2 = dataGrid.GetCellDisplayRectangle(Column, rowEnd, true);

            
            

            for (int i = rowBeginning; i <= rowEnd; i++)
            {
                rectHeight += dataGrid.GetCellDisplayRectangle(Column, i, false).Height;
            }

            Rectangle newCell = new Rectangle(CellRectangle1.X, CellRectangle1.Y, CellRectangle1.Width, rectHeight);

            g.FillRectangle(new SolidBrush(isSelected ? dataGrid.DefaultCellStyle.SelectionBackColor : dataGrid.DefaultCellStyle.BackColor), newCell);

            g.DrawRectangle(gridPen, newCell);

            g.DrawString(dataGrid.Rows[rowBeginning].Cells[Column].Value.ToString(), dataGrid.DefaultCellStyle.Font, new SolidBrush(isSelected ? dataGrid.DefaultCellStyle.SelectionForeColor : dataGrid.DefaultCellStyle.ForeColor), newCell.X + newCell.Width / 3, newCell.Y - 6 + newCell.Height / 2);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CreateExcelFile c;
           // c = new CreateExcelFile();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            LoadXML l;
            l = new LoadXML();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            LoadTemplate t;
            t = new LoadTemplate();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            CreateExcelFile c;
            LoadXML l;
            LoadTemplate t;

            
            l = new LoadXML();
            t = new LoadTemplate();
            c = new CreateExcelFile(l,t);
        }
        //private bool isSelectedCell(int[] Rows, int ColumnIndex)
        //{
        //    if (dataGrid.SelectedCells.Count > 0)
        //    {
        //        for (int iCell = Rows[0]; iCell <= Rows[1]; iCell++)
        //        {
        //            for (int iSelCell = 0; iSelCell < dataGrid.SelectedCells.Count; iSelCell++)
        //            {
        //                if (dataGrid.Rows[iCell].Cells[ColumnIndex] == dataGrid.SelectedCells[iSelCell])
        //                {
        //                    return true;
        //                }
        //            }
        //        }
        //        return false;
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //}
        //private bool isRowsHaveOneCellInFirstColumn(int RowId1, int RowId2)
        //{

        //    foreach (string rowsCollection in MergedRowsInFirstColumn)
        //    {
        //        string[] RowsNumber = rowsCollection.Split(';');

        //        if ((isStringInArray(RowsNumber, RowId1.ToString())) &&
        //            (isStringInArray(RowsNumber, RowId2.ToString())))
        //        {
        //            return true;
        //        }
        //    }
        //    return false;
        //}       

        //private void CollectMergedRowsInFirstColumn(int RowId1, int RowId2)
        //{
        //    string MergedRows = String.Empty;

        //    for (int i = RowId1; i <= RowId2; i++)
        //    {
        //        MergedRows += i.ToString() + ";";
        //    }
        //    MergedRowsInFirstColumn.Add(MergedRows.Remove(MergedRows.Length - 1, 1));
        //}

        //private bool isStringInArray(string[] Array, string value)
        //{
        //    foreach (string item in Array)
        //    {
        //        if (item == value)
        //        {
        //            return true;
        //        }

        //    }
        //    return false;
        //}



        //void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)

        //{

        //    //merge the cell[1,1] and cell[2,1]

        //    if (e.RowIndex == 1)

        //    {

        //        if (e.ColumnIndex == 1)

        //        {

        //            e.PaintBackground(e.ClipBounds, true);

        //            Rectangle r = e.CellBounds;

        //            Rectangle r1 = this.dataGridView1.GetCellDisplayRectangle(2, 1, true);

        //            r.Width += r1.Width - 1;

        //            r.Height -= 1;

        //            using (SolidBrush brBk = new SolidBrush(e.CellStyle.BackColor))

        //            using (SolidBrush brFr = new SolidBrush(e.CellStyle.ForeColor))

        //            {

        //                e.Graphics.FillRectangle(brBk, r);

        //                StringFormat sf = new StringFormat();

        //                sf.Alignment = StringAlignment.Center;

        //                sf.LineAlignment = StringAlignment.Center;

        //                e.Graphics.DrawString("cell merged", e.CellStyle.Font, brFr, r, sf);

        //            }

        //            e.Handled = true;

        //        }

        //        if (e.ColumnIndex == 2)

        //        {

        //            using (Pen p = new Pen(this.dataGridView1.GridColor))

        //            {

        //                e.Graphics.DrawLine(p, e.CellBounds.Left, e.CellBounds.Bottom - 1,

        //                    e.CellBounds.Right, e.CellBounds.Bottom - 1);

        //                e.Graphics.DrawLine(p, e.CellBounds.Right - 1, e.CellBounds.Top,

        //                    e.CellBounds.Right - 1, e.CellBounds.Bottom);

        //            }

        //            e.Handled = true;

        //        }

        //    }

        //}
        //void dataGridView1_Scroll(object sender, ScrollEventArgs e)

        //{

        //    //only redraw the cell[1,1] and cell[2,1] when scrolling

        //    this.dataGridView1.InvalidateCell(1, 1);

        //    this.dataGridView1.InvalidateCell(2, 1);

        //}
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
