namespace EBOMgui
{
    partial class mainFrameScreen
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.rtbConsole = new System.Windows.Forms.RichTextBox();
            this.panel5 = new System.Windows.Forms.Panel();
            this.bDeleteNote = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.bAddNote = new System.Windows.Forms.Button();
            this.lbNotes = new System.Windows.Forms.ListBox();
            this.tbNote = new System.Windows.Forms.TextBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.cbSortColumn = new System.Windows.Forms.ComboBox();
            this.cbSortPriority = new System.Windows.Forms.ComboBox();
            this.bDeleteCustomSort = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.bDeleteSort = new System.Windows.Forms.Button();
            this.bAddSort = new System.Windows.Forms.Button();
            this.lbSort = new System.Windows.Forms.ListBox();
            this.tbSelectedCusomSortOption = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cbSortType = new System.Windows.Forms.ComboBox();
            this.lbSelectedCustomSort = new System.Windows.Forms.ListBox();
            this.bAddCustomSort = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lbCustomSortOptions = new System.Windows.Forms.ListBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.rbHeaderColumn = new System.Windows.Forms.RadioButton();
            this.rbTitleBlock = new System.Windows.Forms.RadioButton();
            this.panel2 = new System.Windows.Forms.Panel();
            this.bDeleteQuantityColumn = new System.Windows.Forms.Button();
            this.bAddQuantityColumn = new System.Windows.Forms.Button();
            this.cbQuantityColumn = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.cbQuantityOptions = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.lbCounting = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dgvEBOM = new System.Windows.Forms.DataGridView();
            this.bOpenFile = new System.Windows.Forms.Button();
            this.lbAttributes = new System.Windows.Forms.ListBox();
            this.panel1.SuspendLayout();
            this.panel5.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvEBOM)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlDark;
            this.panel1.Controls.Add(this.rtbConsole);
            this.panel1.Controls.Add(this.panel5);
            this.panel1.Controls.Add(this.panel4);
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.dgvEBOM);
            this.panel1.Controls.Add(this.bOpenFile);
            this.panel1.Controls.Add(this.lbAttributes);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1105, 795);
            this.panel1.TabIndex = 0;
            // 
            // rtbConsole
            // 
            this.rtbConsole.Location = new System.Drawing.Point(7, 671);
            this.rtbConsole.Name = "rtbConsole";
            this.rtbConsole.Size = new System.Drawing.Size(1087, 114);
            this.rtbConsole.TabIndex = 34;
            this.rtbConsole.Text = "";
            this.rtbConsole.TextChanged += new System.EventHandler(this.rtbConsole_TextChanged);
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.SystemColors.Control;
            this.panel5.Controls.Add(this.bDeleteNote);
            this.panel5.Controls.Add(this.label11);
            this.panel5.Controls.Add(this.bAddNote);
            this.panel5.Controls.Add(this.lbNotes);
            this.panel5.Controls.Add(this.tbNote);
            this.panel5.Location = new System.Drawing.Point(919, 488);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(174, 180);
            this.panel5.TabIndex = 33;
            // 
            // bDeleteNote
            // 
            this.bDeleteNote.Location = new System.Drawing.Point(87, 156);
            this.bDeleteNote.Name = "bDeleteNote";
            this.bDeleteNote.Size = new System.Drawing.Size(72, 21);
            this.bDeleteNote.TabIndex = 39;
            this.bDeleteNote.Text = "Delete Note";
            this.bDeleteNote.UseVisualStyleBackColor = true;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(3, 5);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(35, 13);
            this.label11.TabIndex = 38;
            this.label11.Text = "Notes";
            // 
            // bAddNote
            // 
            this.bAddNote.Location = new System.Drawing.Point(14, 156);
            this.bAddNote.Name = "bAddNote";
            this.bAddNote.Size = new System.Drawing.Size(71, 21);
            this.bAddNote.TabIndex = 38;
            this.bAddNote.Text = "Add Note";
            this.bAddNote.UseVisualStyleBackColor = true;
            // 
            // lbNotes
            // 
            this.lbNotes.FormattingEnabled = true;
            this.lbNotes.Location = new System.Drawing.Point(3, 26);
            this.lbNotes.Name = "lbNotes";
            this.lbNotes.Size = new System.Drawing.Size(168, 95);
            this.lbNotes.TabIndex = 38;
            // 
            // tbNote
            // 
            this.tbNote.Location = new System.Drawing.Point(3, 131);
            this.tbNote.Name = "tbNote";
            this.tbNote.Size = new System.Drawing.Size(168, 20);
            this.tbNote.TabIndex = 0;
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.SystemColors.Control;
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.Controls.Add(this.cbSortColumn);
            this.panel4.Controls.Add(this.cbSortPriority);
            this.panel4.Controls.Add(this.bDeleteCustomSort);
            this.panel4.Controls.Add(this.label12);
            this.panel4.Controls.Add(this.bDeleteSort);
            this.panel4.Controls.Add(this.bAddSort);
            this.panel4.Controls.Add(this.lbSort);
            this.panel4.Controls.Add(this.tbSelectedCusomSortOption);
            this.panel4.Controls.Add(this.label7);
            this.panel4.Controls.Add(this.label1);
            this.panel4.Controls.Add(this.label6);
            this.panel4.Controls.Add(this.label2);
            this.panel4.Controls.Add(this.cbSortType);
            this.panel4.Controls.Add(this.lbSelectedCustomSort);
            this.panel4.Controls.Add(this.bAddCustomSort);
            this.panel4.Controls.Add(this.label4);
            this.panel4.Controls.Add(this.label5);
            this.panel4.Controls.Add(this.lbCustomSortOptions);
            this.panel4.Location = new System.Drawing.Point(7, 488);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(569, 181);
            this.panel4.TabIndex = 32;
            // 
            // cbSortColumn
            // 
            this.cbSortColumn.FormattingEnabled = true;
            this.cbSortColumn.Items.AddRange(new object[] {
            "Ascending",
            "Descending",
            "Custom (Starts With)",
            "Custom (Ends With)"});
            this.cbSortColumn.Location = new System.Drawing.Point(247, 25);
            this.cbSortColumn.Name = "cbSortColumn";
            this.cbSortColumn.Size = new System.Drawing.Size(129, 21);
            this.cbSortColumn.TabIndex = 38;
            // 
            // cbSortPriority
            // 
            this.cbSortPriority.FormattingEnabled = true;
            this.cbSortPriority.Items.AddRange(new object[] {
            "Ascending",
            "Descending",
            "Custom (Starts With)",
            "Custom (Ends With)"});
            this.cbSortPriority.Location = new System.Drawing.Point(247, 64);
            this.cbSortPriority.Name = "cbSortPriority";
            this.cbSortPriority.Size = new System.Drawing.Size(129, 21);
            this.cbSortPriority.TabIndex = 37;
            // 
            // bDeleteCustomSort
            // 
            this.bDeleteCustomSort.Location = new System.Drawing.Point(436, 152);
            this.bDeleteCustomSort.Name = "bDeleteCustomSort";
            this.bDeleteCustomSort.Size = new System.Drawing.Size(60, 20);
            this.bDeleteCustomSort.TabIndex = 35;
            this.bDeleteCustomSort.Text = "Delete";
            this.bDeleteCustomSort.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label12.Location = new System.Drawing.Point(382, 4);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(1, 170);
            this.label12.TabIndex = 34;
            // 
            // bDeleteSort
            // 
            this.bDeleteSort.Location = new System.Drawing.Point(285, 138);
            this.bDeleteSort.Name = "bDeleteSort";
            this.bDeleteSort.Size = new System.Drawing.Size(72, 30);
            this.bDeleteSort.TabIndex = 22;
            this.bDeleteSort.Text = "Delete Sort";
            this.bDeleteSort.UseVisualStyleBackColor = true;
            // 
            // bAddSort
            // 
            this.bAddSort.Location = new System.Drawing.Point(211, 138);
            this.bAddSort.Name = "bAddSort";
            this.bAddSort.Size = new System.Drawing.Size(57, 30);
            this.bAddSort.TabIndex = 21;
            this.bAddSort.Text = "Add Sort";
            this.bAddSort.UseVisualStyleBackColor = true;
            // 
            // lbSort
            // 
            this.lbSort.FormattingEnabled = true;
            this.lbSort.Location = new System.Drawing.Point(3, 14);
            this.lbSort.Name = "lbSort";
            this.lbSort.Size = new System.Drawing.Size(180, 160);
            this.lbSort.TabIndex = 3;
            // 
            // tbSelectedCusomSortOption
            // 
            this.tbSelectedCusomSortOption.Location = new System.Drawing.Point(390, 130);
            this.tbSelectedCusomSortOption.Name = "tbSelectedCusomSortOption";
            this.tbSelectedCusomSortOption.Size = new System.Drawing.Size(106, 20);
            this.tbSelectedCusomSortOption.TabIndex = 6;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(0, 1);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(40, 13);
            this.label7.TabIndex = 20;
            this.label7.Text = "Sorting";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(193, 67);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Priority";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(193, 29);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(42, 13);
            this.label6.TabIndex = 19;
            this.label6.Text = "Column";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(193, 103);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(31, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Type";
            // 
            // cbSortType
            // 
            this.cbSortType.FormattingEnabled = true;
            this.cbSortType.Items.AddRange(new object[] {
            "Ascending",
            "Descending",
            "Custom (Starts With)",
            "Custom (Ends With)"});
            this.cbSortType.Location = new System.Drawing.Point(248, 102);
            this.cbSortType.Name = "cbSortType";
            this.cbSortType.Size = new System.Drawing.Size(129, 21);
            this.cbSortType.TabIndex = 9;
            // 
            // lbSelectedCustomSort
            // 
            this.lbSelectedCustomSort.FormattingEnabled = true;
            this.lbSelectedCustomSort.Location = new System.Drawing.Point(390, 29);
            this.lbSelectedCustomSort.Name = "lbSelectedCustomSort";
            this.lbSelectedCustomSort.Size = new System.Drawing.Size(106, 95);
            this.lbSelectedCustomSort.TabIndex = 11;
            // 
            // bAddCustomSort
            // 
            this.bAddCustomSort.Location = new System.Drawing.Point(390, 152);
            this.bAddCustomSort.Name = "bAddCustomSort";
            this.bAddCustomSort.Size = new System.Drawing.Size(40, 20);
            this.bAddCustomSort.TabIndex = 15;
            this.bAddCustomSort.Text = "Add";
            this.bAddCustomSort.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(499, 13);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Sort Options";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(387, 13);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(109, 13);
            this.label5.TabIndex = 14;
            this.label5.Text = "Selected Custom Sort";
            // 
            // lbCustomSortOptions
            // 
            this.lbCustomSortOptions.FormattingEnabled = true;
            this.lbCustomSortOptions.Location = new System.Drawing.Point(502, 29);
            this.lbCustomSortOptions.Name = "lbCustomSortOptions";
            this.lbCustomSortOptions.Size = new System.Drawing.Size(55, 147);
            this.lbCustomSortOptions.TabIndex = 13;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.SystemColors.Control;
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.rbHeaderColumn);
            this.panel3.Controls.Add(this.rbTitleBlock);
            this.panel3.Location = new System.Drawing.Point(195, 5);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(104, 50);
            this.panel3.TabIndex = 31;
            // 
            // rbHeaderColumn
            // 
            this.rbHeaderColumn.AutoSize = true;
            this.rbHeaderColumn.Location = new System.Drawing.Point(5, 28);
            this.rbHeaderColumn.Name = "rbHeaderColumn";
            this.rbHeaderColumn.Size = new System.Drawing.Size(98, 17);
            this.rbHeaderColumn.TabIndex = 1;
            this.rbHeaderColumn.Text = "Header Column";
            this.rbHeaderColumn.UseVisualStyleBackColor = true;
            this.rbHeaderColumn.CheckedChanged += new System.EventHandler(this.rbHeaderColumn_CheckedChanged);
            // 
            // rbTitleBlock
            // 
            this.rbTitleBlock.AutoSize = true;
            this.rbTitleBlock.Checked = true;
            this.rbTitleBlock.Location = new System.Drawing.Point(5, 5);
            this.rbTitleBlock.Name = "rbTitleBlock";
            this.rbTitleBlock.Size = new System.Drawing.Size(75, 17);
            this.rbTitleBlock.TabIndex = 0;
            this.rbTitleBlock.TabStop = true;
            this.rbTitleBlock.Text = "Title Block";
            this.rbTitleBlock.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.Control;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.bDeleteQuantityColumn);
            this.panel2.Controls.Add(this.bAddQuantityColumn);
            this.panel2.Controls.Add(this.cbQuantityColumn);
            this.panel2.Controls.Add(this.label9);
            this.panel2.Controls.Add(this.cbQuantityOptions);
            this.panel2.Controls.Add(this.label8);
            this.panel2.Controls.Add(this.label10);
            this.panel2.Controls.Add(this.lbCounting);
            this.panel2.Location = new System.Drawing.Point(582, 489);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(331, 180);
            this.panel2.TabIndex = 30;
            // 
            // bDeleteQuantityColumn
            // 
            this.bDeleteQuantityColumn.Location = new System.Drawing.Point(242, 93);
            this.bDeleteQuantityColumn.Name = "bDeleteQuantityColumn";
            this.bDeleteQuantityColumn.Size = new System.Drawing.Size(87, 22);
            this.bDeleteQuantityColumn.TabIndex = 36;
            this.bDeleteQuantityColumn.Text = "Delete Column";
            this.bDeleteQuantityColumn.UseVisualStyleBackColor = true;
            // 
            // bAddQuantityColumn
            // 
            this.bAddQuantityColumn.Location = new System.Drawing.Point(162, 93);
            this.bAddQuantityColumn.Name = "bAddQuantityColumn";
            this.bAddQuantityColumn.Size = new System.Drawing.Size(80, 22);
            this.bAddQuantityColumn.TabIndex = 35;
            this.bAddQuantityColumn.Text = "Add Column";
            this.bAddQuantityColumn.UseVisualStyleBackColor = true;
            // 
            // cbQuantityColumn
            // 
            this.cbQuantityColumn.FormattingEnabled = true;
            this.cbQuantityColumn.Items.AddRange(new object[] {
            "Ascending",
            "Descending",
            "Custom (Starts With)",
            "Custom (Ends With)"});
            this.cbQuantityColumn.Location = new System.Drawing.Point(170, 26);
            this.cbQuantityColumn.Name = "cbQuantityColumn";
            this.cbQuantityColumn.Size = new System.Drawing.Size(147, 21);
            this.cbQuantityColumn.TabIndex = 37;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(167, 13);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(84, 13);
            this.label9.TabIndex = 36;
            this.label9.Text = "Quantity Column";
            // 
            // cbQuantityOptions
            // 
            this.cbQuantityOptions.FormattingEnabled = true;
            this.cbQuantityOptions.Items.AddRange(new object[] {
            "Ascending",
            "Descending",
            "Custom (Starts With)",
            "Custom (Ends With)"});
            this.cbQuantityOptions.Location = new System.Drawing.Point(170, 66);
            this.cbQuantityOptions.Name = "cbQuantityOptions";
            this.cbQuantityOptions.Size = new System.Drawing.Size(147, 21);
            this.cbQuantityOptions.TabIndex = 35;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(3, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(106, 13);
            this.label8.TabIndex = 21;
            this.label8.Text = "Component Counting";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(167, 50);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(85, 13);
            this.label10.TabIndex = 34;
            this.label10.Text = "Quantity Options";
            // 
            // lbCounting
            // 
            this.lbCounting.FormattingEnabled = true;
            this.lbCounting.Location = new System.Drawing.Point(3, 15);
            this.lbCounting.Name = "lbCounting";
            this.lbCounting.Size = new System.Drawing.Size(158, 160);
            this.lbCounting.TabIndex = 33;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(108, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "Component Attributes";
            // 
            // dgvEBOM
            // 
            this.dgvEBOM.AllowDrop = true;
            this.dgvEBOM.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvEBOM.Location = new System.Drawing.Point(197, 61);
            this.dgvEBOM.Name = "dgvEBOM";
            this.dgvEBOM.ReadOnly = true;
            this.dgvEBOM.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvEBOM.Size = new System.Drawing.Size(896, 420);
            this.dgvEBOM.TabIndex = 2;
            this.dgvEBOM.CellMouseDown += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvEBOM_CellMouseDown);
            this.dgvEBOM.CellMouseEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvEBOM_CellMouseEnter);
            this.dgvEBOM.SelectionChanged += new System.EventHandler(this.dgvEBOM_SelectionChanged);
            this.dgvEBOM.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dgvEBOM_MouseUp);
            // 
            // bOpenFile
            // 
            this.bOpenFile.Location = new System.Drawing.Point(15, 5);
            this.bOpenFile.Name = "bOpenFile";
            this.bOpenFile.Size = new System.Drawing.Size(75, 23);
            this.bOpenFile.TabIndex = 1;
            this.bOpenFile.Text = "Open File";
            this.bOpenFile.UseVisualStyleBackColor = true;
            this.bOpenFile.Click += new System.EventHandler(this.bOpenFile_Click);
            // 
            // lbAttributes
            // 
            this.lbAttributes.FormattingEnabled = true;
            this.lbAttributes.Location = new System.Drawing.Point(7, 61);
            this.lbAttributes.Name = "lbAttributes";
            this.lbAttributes.Size = new System.Drawing.Size(182, 420);
            this.lbAttributes.TabIndex = 0;
            this.lbAttributes.SelectedIndexChanged += new System.EventHandler(this.lbAttributes_SelectedIndexChanged);
            this.lbAttributes.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lbAttributes_MouseDown);
            this.lbAttributes.MouseLeave += new System.EventHandler(this.lbAttributes_MouseLeave);
            this.lbAttributes.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lbAttributes_MouseUp);
            // 
            // mainFrameScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1101, 791);
            this.Controls.Add(this.panel1);
            this.Name = "mainFrameScreen";
            this.Text = "EBOMgui";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel5.ResumeLayout(false);
            this.panel5.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvEBOM)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button bOpenFile;
        private System.Windows.Forms.ListBox lbAttributes;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbSelectedCusomSortOption;
        private System.Windows.Forms.ListBox lbSort;
        private System.Windows.Forms.ComboBox cbSortType;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button bAddCustomSort;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ListBox lbCustomSortOptions;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListBox lbSelectedCustomSort;
        private System.Windows.Forms.RadioButton rbHeaderColumn;
        private System.Windows.Forms.RadioButton rbTitleBlock;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ListBox lbCounting;
        private System.Windows.Forms.Button bDeleteSort;
        private System.Windows.Forms.Button bAddSort;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cbQuantityOptions;
        private System.Windows.Forms.ComboBox cbQuantityColumn;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Button bDeleteNote;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button bAddNote;
        private System.Windows.Forms.ListBox lbNotes;
        private System.Windows.Forms.TextBox tbNote;
        private System.Windows.Forms.Button bDeleteQuantityColumn;
        private System.Windows.Forms.Button bAddQuantityColumn;
        private System.Windows.Forms.DataGridView dgvEBOM;
        private System.Windows.Forms.Button bDeleteCustomSort;
        private System.Windows.Forms.RichTextBox rtbConsole;
        private System.Windows.Forms.ComboBox cbSortColumn;
        private System.Windows.Forms.ComboBox cbSortPriority;
    }
}

