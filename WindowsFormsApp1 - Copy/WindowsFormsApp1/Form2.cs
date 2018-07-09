using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;


namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        DataGridView dataGrid;
        DataSet dataSet;
        Button button;
        Thread gui;

        private List<string> MergedRowsInFirstColumn = new List<string>();
        private bool readyToMerge = false;


        public Form2()
        {
            InitializeComponent();
            CreateGridAndButton();
        }

        private void CreateGridAndButton()
        {
            dataGrid = new DataGridView();
            dataSet = new DataSet();

            dataGrid.Height = this.Height - 100;
            dataGrid.Dock = DockStyle.Top;
            dataGrid.ReadOnly = true;
            dataGrid.AllowUserToAddRows = false;
            dataGrid.AllowUserToResizeRows = false;
            dataGrid.RowHeadersVisible = false;
            this.dataGrid.Paint += new PaintEventHandler(dataGrid_Paint);

            this.Controls.Add(this.dataGrid);

            button = new Button();
            button.Text = "Merge";
            button.Dock = DockStyle.Bottom;
            button.Click += new System.EventHandler(this.button_Click);

            this.Controls.Add(this.button);

           
            DataTable tbl_main = new DataTable("tbl_main");

            tbl_main.Columns.Add("Manufacture");
            tbl_main.Columns.Add("Name");
            tbl_main.Columns.Add("CPU");
            tbl_main.Columns.Add("RAM");
            tbl_main.Columns.Add("Price");

            DataRow row;

            row = tbl_main.NewRow();
            row["Manufacture"] = "Dell";
            row["Name"] = "Inspiron 1525";
            row["CPU"] = "T7250";
            row["RAM"] = "2048 MB";
            row["Price"] = "$654";
            tbl_main.Rows.Add(row);

            row = tbl_main.NewRow();
            row["Manufacture"] = "Dell";
            row["Name"] = "Inspiron 1525";
            row["CPU"] = "T5750";
            row["RAM"] = "2048 MB";
            row["Price"] = "$690";
            tbl_main.Rows.Add(row);

            row = tbl_main.NewRow();
            row["Manufacture"] = "Dell";
            row["Name"] = "Studio 1535";
            row["CPU"] = "T5750";
            row["RAM"] = "2048 MB";
            row["Price"] = "$735";
            tbl_main.Rows.Add(row);

            row = tbl_main.NewRow();
            row["Manufacture"] = "Dell";
            row["Name"] = "Vostro 1510";
            row["CPU"] = "T5870";
            row["RAM"] = "2048 MB";
            row["Price"] = "$724";
            tbl_main.Rows.Add(row);

            row = tbl_main.NewRow();
            row["Manufacture"] = "HP";
            row["Name"] = "530";
            row["CPU"] = "T5200";
            row["RAM"] = "1024 MB";
            row["Price"] = "$545";
            tbl_main.Rows.Add(row);

            row = tbl_main.NewRow();
            row["Manufacture"] = "HP";
            row["Name"] = "6720s";
            row["CPU"] = "T5670";
            row["RAM"] = "1024 MB";
            row["Price"] = "$637";
            tbl_main.Rows.Add(row);

            row = tbl_main.NewRow();
            row["Manufacture"] = "HP";
            row["Name"] = "Pavilion dv9233";
            row["CPU"] = "T5670";
            row["RAM"] = "1024 MB";
            row["Price"] = "$780";
            tbl_main.Rows.Add(row);

            dataSet.Tables.Add(tbl_main);
            dataGrid.DataSource = dataSet;
            dataGrid.DataMember = "tbl_main";

        }

        private void dataGrid_Paint(object sender, PaintEventArgs e)
        {
            if (readyToMerge)
            {
                Merge();
            }
        }

        private void button_Click(object sender, EventArgs e)
        {
            gui = new Thread(delegate ()
            {
                Merge();
                readyToMerge = true;
            });
            gui.Name = "gui";
            gui.Start();
           
        }

        private bool isSelectedCell(int[] Rows, int ColumnIndex)
        {

            if (dataGrid.SelectedCells.Count > 0)
            {
                for (int iCell = Rows[0]; iCell <= Rows[1]; iCell++)
                {
                    for (int iSelCell = 0; iSelCell < dataGrid.SelectedCells.Count; iSelCell++)
                    {
                        if (dataGrid.Rows[iCell].Cells[ColumnIndex] == dataGrid.SelectedCells[iSelCell])
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
            else
            {
                return false;
            }
        }

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

        private bool isRowsHaveOneCellInFirstColumn(int RowId1, int RowId2)
        {

            foreach (string rowsCollection in MergedRowsInFirstColumn)
            {
                string[] RowsNumber = rowsCollection.Split(';');

                if ((isStringInArray(RowsNumber, RowId1.ToString())) &&
                    (isStringInArray(RowsNumber, RowId2.ToString())))
                {
                    return true;
                }
            }
            return false;
        }

        private bool isStringInArray(string[] Array, string value)
        {
            foreach (string item in Array)
            {
                if (item == value)
                {
                    return true;
                }

            }
            return false;
        }

        private void CollectMergedRowsInFirstColumn(int RowId1, int RowId2)
        {
            string MergedRows = String.Empty;

            for (int i = RowId1; i <= RowId2; i++)
            {
                MergedRows += i.ToString() + ";";
            }
            MergedRowsInFirstColumn.Add(MergedRows.Remove(MergedRows.Length - 1, 1));
        }

        private void MergeCells(int RowId1, int RowId2, int Column, bool isSelected)
        {
            Graphics g = dataGrid.CreateGraphics();
            Pen gridPen = new Pen(dataGrid.GridColor);

            //Cells Rectangles
            Rectangle CellRectangle1 = dataGrid.GetCellDisplayRectangle(Column, RowId1, true);
            Rectangle CellRectangle2 = dataGrid.GetCellDisplayRectangle(Column, RowId2, true);

            int rectHeight = 0;
            string MergedRows = String.Empty;

            for (int i = RowId1; i <= RowId2; i++)
            {
                rectHeight += dataGrid.GetCellDisplayRectangle(Column, i, false).Height;
            }

            Rectangle newCell = new Rectangle(CellRectangle1.X, CellRectangle1.Y, CellRectangle1.Width, rectHeight);

            g.FillRectangle(new SolidBrush(isSelected ? dataGrid.DefaultCellStyle.SelectionBackColor : dataGrid.DefaultCellStyle.BackColor), newCell);

            g.DrawRectangle(gridPen, newCell);

            g.DrawString(dataGrid.Rows[RowId1].Cells[Column].Value.ToString(), dataGrid.DefaultCellStyle.Font, new SolidBrush(isSelected ? dataGrid.DefaultCellStyle.SelectionForeColor : dataGrid.DefaultCellStyle.ForeColor), newCell.X + newCell.Width / 3, newCell.Y + newCell.Height / 3);
        }
    }
}