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
        int currentRow, currentColumn;

        List<int> attributeRows, attributeColumn, attributeIndex;
        List<string> attributeNames;
        List<List<string>> componentContent;
        

        int prevRow, prevColumn;
        string prevText;
        bool prevHeader;
        Color prevColorAttribute, prevColorComponent, prevColorTitle;

        string dragName;
        int prevSetIndex, currentAttributeIndex;

        dgvInfo prevCellInfo,dragInfo; // for use in dragging a cell set around and having the cells your are dragging over return to their previous settings.

        public bool dgvMouseDown = false, leftMouseDown = false, moveAttribute = false, lbMouseDown = false, newAttribute = false, newPaint = false;

        public bool headerType = false;
        xmlFileHandler xmlFileHandler1;
        mainFrameScreen mainFrameScreen1;
        public MainFrame(mainFrameScreen m)
        {
            mainFrameScreen1 = m;
            prevCellInfo = new dgvInfo();
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

        public bool determineDrag(ref bool mouseDown)
        {
            int count = 0;
            while (mouseDown)
            {
                Thread.Sleep(10);
                if (count++ == 20)
                {
                    return true;
                }
            }
            return false;
        }




        public void writeToConsole(string text)
        {
            mainFrameScreen1.writeToConsole(text);
        }


        ///////////////////////////////////lbAttributes Methods //////////////////////////////////////////
        public void lbMouseUpEvent(MouseEventArgs e)
        {
            lbMouseDown = false;
        }
        public void lbMouseDownEvent(MouseEventArgs e, int index)
        {
            if (index > -1 && xmlFileHandler1 != null) // if index is -1 the the user didn't actually click on an item in the list box
            {
                currentAttributeIndex = index; // held int for determining name of new attribute and all other components
                lbMouseDown = true;
            }
            
            //Thread run = new Thread(delegate ()
            //{
            //    if (determineDrag(ref lbMouseDown))
            //    {
            //        lbDragTrue = true;
            //        generateDragList(1, 1, attributeNames[index]);
            //    }
                    
            //});
            //run.Name = "drag";
            //run.Start();
        }


        ///////////////////////////////////lbAttributes Methods //////////////////////////////////////////

        ///////////////////////////////////dgvEBOM Methods //////////////////////////////////////////
        public void dgvMouseUpEvent(MouseEventArgs e)
        {
            if (moveAttribute)
            {
                moveAttribute = false;
                if (currentRow > -1 && currentColumn > -1)
                {
                    attributeRows[prevSetIndex] = dragInfo.rows[0];
                    attributeColumn[prevSetIndex] = dragInfo.columns[0];                }
                else
                {
                    attributeRows.RemoveAt(prevSetIndex);
                    attributeColumn.RemoveAt(prevSetIndex);
                    attributeIndex.RemoveAt(prevSetIndex);
                    attributeNames.RemoveAt(prevSetIndex);
                    componentContent.RemoveAt(prevSetIndex);
                }
            }
            if (newAttribute)
            {
                newAttribute = false;
                if (currentRow > -1 && currentColumn > -1)
                {
                    attributeRows.Add(dragInfo.rows[0]);
                    attributeColumn.Add(dragInfo.columns[0]);
                    attributeIndex.Add(dragInfo.[0]);
                    attributeNames.Add(dragInfo.text[0]);
                    componentContent.Add(dragInfo.rows[0]);
                }
                else
                {

                }
            }
            //newAttribute = false;
            //moveAttribute = false;

        }

        public void dgvMouseDownEvent(DataGridViewCellMouseEventArgs e)
        {
            dgvMouseDown = true;

            prevSetIndex = checkIfCellIsOccupied(e.RowIndex, e.ColumnIndex);
            //run.Join();
            if (prevSetIndex < 0 || xmlFileHandler1 == null)
            {
                dgvMouseDown = false;
                mainFrameScreen1.writeLineToConsole("drag = false");
            }
            else
            {
                mainFrameScreen1.clearlbAttribute(); 
                currentAttributeIndex = attributeIndex[prevSetIndex]; // held int for determining name of new attribute and all other components
            }


        }

        public void dgvEBOM_CellMouseEnter(int row, int column)
        {
            currentRow = row;
            currentColumn = column;
            if (lbMouseDown)
            {
                lbMouseDown = false;
                newAttribute = true;
                 generateDragList(row, column, xmlFileHandler1.attributeNames[currentAttributeIndex]);
                newPaint = true;
            }
            if (dgvMouseDown)
            {
                dgvMouseDown = false;
                moveAttribute = true;
                generateDragList(row, column, xmlFileHandler1.attributeNames[currentAttributeIndex]);
                newPaint = true;
            }

            if (row > -1 && column > -1)
            {
                if (moveAttribute || newAttribute)
                {
                    generateDragList(row, column, xmlFileHandler1.attributeNames[currentAttributeIndex]);
                    paintCells(headerType, row, column, prevText, newPaint);
                    if (newPaint) newPaint = false;
                }
            }

        }

        public void paintCells(bool header1, int row1, int column1, string text1, bool newPaint1)
        {
            if (!newPaint1)
            {
                setCellInfo(prevCellInfo.rows, prevCellInfo.columns, prevCellInfo.colors, prevCellInfo.text);
            }
            //paint(header1, row1, column1, text1, Color.Green, Color.Yellow, Color.Orange);

            if (header1)
            {
                getCellInfo(row1, row1 + 10, column1, column1);
                setCellInfo(dragInfo.rows, dragInfo.columns, dragInfo.colors, dragInfo.text);
            }
            else
            {
                getCellInfo(row1, row1, column1, column1 + 1);
                setCellInfo(dragInfo.rows, dragInfo.columns, dragInfo.colors, dragInfo.text);
            }
            //void paint(bool header, int row, int column, string text, Color attribute, Color title, Color component)
            //{

            //    mainFrameScreen1.dgvEBOM_ChangeColor(attribute, row, column);
            //    mainFrameScreen1.dgvEBOM_setText(text, row, column);
            //    if (header)
            //    {
            //        getCellInfo(row, row + 10, column, column);
            //        for (int a = row + 1; a < row + 11; a++)
            //        {
            //            mainFrameScreen1.dgvEBOM_ChangeColor(component, a, column);
            //        }
            //    }
            //    else
            //    {
            //        getCellInfo(row, row, column, column + 1);
            //        mainFrameScreen1.dgvEBOM_ChangeColor(title, row, column + 1);
            //    }
            //}

        }

        public void setCellsInDGVEBOM()
        {

        }

        // gets all cell info for a range of rows and columns and saves it to the dgvInfo object
        //usually used for when a user is dragging a value over the grid and needs to replace the contents of cells as the mouse passes over a cell section
        public void setCellInfo(List<int> rows, List<int> columns, List<Color> colors, List<string> text)
        {
            for (int a = 0; a < rows.Count; a++)
            {
                mainFrameScreen1.dgvEBOM_ChangeColor(colors[a], rows[a], columns[a]);
                mainFrameScreen1.dgvEBOM_setText(text[a], rows[a], columns[a]);
            }
        }
        public void getCellInfo(int begRow, int endRow, int begColumn, int endColumn)
        {
            prevCellInfo = new dgvInfo();
            string tempText = ""; Color tempColor = Color.Red;
            for (int a = begRow; a < endRow + 1; a++)
            {
                for (int b = begColumn; b < endColumn + 1; b++)
                {
                    mainFrameScreen1.getdgvEBOMCellInfo(ref tempText, ref tempColor, a, b);
                    prevCellInfo.colors.Add(tempColor);
                    prevCellInfo.text.Add(tempText);
                    prevCellInfo.rows.Add(a);
                    prevCellInfo.columns.Add(b);
                }
            }
        }

        //creates a list of all the cells we need for dragging around a section of cells to decide where we want something.
        public void generateDragList(int row, int column, string name)
        {
            dragInfo = new dgvInfo();
            dragInfo.text.Add(name);
            dragInfo.rows.Add(row);
            dragInfo.columns.Add(column);
            dragInfo.colors.Add(Color.Green);
            if (headerType)
            {
                for (int a = row + 1; a < row + 11; a++)
                {
                    dragInfo.rows.Add(a);
                    dragInfo.columns.Add(column);
                    dragInfo.text.Add("");
                    dragInfo.colors.Add(Color.Yellow);
                }
            }
            else
            {
                for (int a = column + 1; a < column + 2; a++)
                {
                    dragInfo.rows.Add(row);
                    dragInfo.columns.Add(a);
                    dragInfo.text.Add("");
                    dragInfo.colors.Add(Color.Orange);
                }
            }
        }

        // return index of existing attribute populated into excel template
        public int checkIfCellIsOccupied(int currentRow, int currentColumn)
        {
            if (attributeRows != null)
                for (int a = 0; a < attributeRows.Count; a++)
                    if (attributeRows[a] == currentRow)
                        if (attributeColumn[a] == currentColumn)
                            return a;
            return -1; // return -1 if it doesn't match existing column
        }

       

        public void dgvDragEvent(int row, int column)
        {
            dgvMouseDown = true;
            //Thread run = new Thread(delegate ()
            //{
            //    if (determineDrag(ref dgvMouseDown))
            //    {
            //        dgvDragTrue = true;
            //    }
            //})
            //{
            //    Name = "drag"
            //};
            //run.Start();
            

        }


        public void dgvSelectionChanged()
        {
            //if (rightMouseDown) mainFrameScreen1.dgvEBOMClearSelected();
        }
        ///////////////////////////////////dgvEBOM Methods //////////////////////////////////////////






        

       
    }
}
