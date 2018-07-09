namespace EBOMCreationTool
{
    partial class MainFrame
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
            this.bChooseTemplate = new System.Windows.Forms.Button();
            this.bStart = new System.Windows.Forms.Button();
            this.bChooseXML = new System.Windows.Forms.Button();
            this.tbTemplate = new System.Windows.Forms.TextBox();
            this.tbXML = new System.Windows.Forms.TextBox();
            this.ChooseExport = new System.Windows.Forms.Button();
            this.tbExport = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // bChooseTemplate
            // 
            this.bChooseTemplate.Location = new System.Drawing.Point(12, 12);
            this.bChooseTemplate.Name = "bChooseTemplate";
            this.bChooseTemplate.Size = new System.Drawing.Size(218, 46);
            this.bChooseTemplate.TabIndex = 5;
            this.bChooseTemplate.Text = "Choose Template";
            this.bChooseTemplate.UseVisualStyleBackColor = true;
            this.bChooseTemplate.Click += new System.EventHandler(this.bChooseTemplate_Click);
            // 
            // bStart
            // 
            this.bStart.Location = new System.Drawing.Point(195, 186);
            this.bStart.Name = "bStart";
            this.bStart.Size = new System.Drawing.Size(218, 46);
            this.bStart.TabIndex = 6;
            this.bStart.Text = "Start";
            this.bStart.UseVisualStyleBackColor = true;
            this.bStart.Click += new System.EventHandler(this.bStart_Click);
            // 
            // bChooseXML
            // 
            this.bChooseXML.Location = new System.Drawing.Point(12, 64);
            this.bChooseXML.Name = "bChooseXML";
            this.bChooseXML.Size = new System.Drawing.Size(218, 46);
            this.bChooseXML.TabIndex = 7;
            this.bChooseXML.Text = "Choose XML Source";
            this.bChooseXML.UseVisualStyleBackColor = true;
            this.bChooseXML.Click += new System.EventHandler(this.bChooseXML_Click);
            // 
            // tbTemplate
            // 
            this.tbTemplate.Location = new System.Drawing.Point(249, 26);
            this.tbTemplate.Name = "tbTemplate";
            this.tbTemplate.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.tbTemplate.Size = new System.Drawing.Size(355, 20);
            this.tbTemplate.TabIndex = 8;
            this.tbTemplate.Text = "C:\\Users\\christopher.pratt\\Documents\\000EBOM\\EBOM\\EBOM\\EBOM_Creation_Tool\\EBOM_Cr" +
    "eation_Tool\\bin\\Debug\\template.xlsx";
            this.tbTemplate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // tbXML
            // 
            this.tbXML.Location = new System.Drawing.Point(249, 78);
            this.tbXML.Name = "tbXML";
            this.tbXML.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.tbXML.Size = new System.Drawing.Size(355, 20);
            this.tbXML.TabIndex = 9;
            this.tbXML.Text = "C:\\Users\\christopher.pratt\\Documents\\000EBOM\\EBOM\\EBOM\\EBOM_Creation_Tool\\EBOM_Cr" +
    "eation_Tool\\bin\\Debug\\altium.xml";
            this.tbXML.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // ChooseExport
            // 
            this.ChooseExport.Location = new System.Drawing.Point(12, 116);
            this.ChooseExport.Name = "ChooseExport";
            this.ChooseExport.Size = new System.Drawing.Size(218, 46);
            this.ChooseExport.TabIndex = 10;
            this.ChooseExport.Text = "Choose Export Location";
            this.ChooseExport.UseVisualStyleBackColor = true;
            this.ChooseExport.Click += new System.EventHandler(this.ChooseExport_Click);
            // 
            // tbExport
            // 
            this.tbExport.Location = new System.Drawing.Point(249, 130);
            this.tbExport.Name = "tbExport";
            this.tbExport.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.tbExport.Size = new System.Drawing.Size(355, 20);
            this.tbExport.TabIndex = 11;
            this.tbExport.Text = "C:\\Users\\christopher.pratt\\Documents\\000EBOM\\EBOM\\EBOM\\EBOM_Creation_Tool\\EBOM_Cr" +
    "eation_Tool\\bin\\Debug\\New_EBOM.xlsx";
            this.tbExport.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.tbExport.TextChanged += new System.EventHandler(this.tbExport_TextChanged);
            // 
            // MainFrame
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(619, 263);
            this.Controls.Add(this.tbExport);
            this.Controls.Add(this.ChooseExport);
            this.Controls.Add(this.tbXML);
            this.Controls.Add(this.tbTemplate);
            this.Controls.Add(this.bChooseXML);
            this.Controls.Add(this.bStart);
            this.Controls.Add(this.bChooseTemplate);
            this.DoubleBuffered = true;
            this.Name = "MainFrame";
            this.Text = "EBOM_Creation_Tool";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button bChooseTemplate;
        private System.Windows.Forms.Button bStart;
        private System.Windows.Forms.Button bChooseXML;
        private System.Windows.Forms.TextBox tbTemplate;
        private System.Windows.Forms.TextBox tbXML;
        private System.Windows.Forms.Button ChooseExport;
        private System.Windows.Forms.TextBox tbExport;
    }
}

