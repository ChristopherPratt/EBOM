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
            this.bStart = new System.Windows.Forms.Button();
            this.bChooseXML = new System.Windows.Forms.Button();
            this.tbXML = new System.Windows.Forms.TextBox();
            this.rtbConsole = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // bStart
            // 
            this.bStart.Location = new System.Drawing.Point(202, 64);
            this.bStart.Name = "bStart";
            this.bStart.Size = new System.Drawing.Size(218, 46);
            this.bStart.TabIndex = 6;
            this.bStart.Text = "Start";
            this.bStart.UseVisualStyleBackColor = true;
            this.bStart.Click += new System.EventHandler(this.bStart_Click);
            // 
            // bChooseXML
            // 
            this.bChooseXML.Location = new System.Drawing.Point(15, 12);
            this.bChooseXML.Name = "bChooseXML";
            this.bChooseXML.Size = new System.Drawing.Size(218, 46);
            this.bChooseXML.TabIndex = 7;
            this.bChooseXML.Text = "Choose XML Source";
            this.bChooseXML.UseVisualStyleBackColor = true;
            this.bChooseXML.Click += new System.EventHandler(this.bChooseXML_Click);
            // 
            // tbXML
            // 
            this.tbXML.Location = new System.Drawing.Point(252, 26);
            this.tbXML.Name = "tbXML";
            this.tbXML.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.tbXML.Size = new System.Drawing.Size(355, 20);
            this.tbXML.TabIndex = 9;
            this.tbXML.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.tbXML.TextChanged += new System.EventHandler(this.tbXML_TextChanged);
            // 
            // rtbConsole
            // 
            this.rtbConsole.Location = new System.Drawing.Point(15, 129);
            this.rtbConsole.Name = "rtbConsole";
            this.rtbConsole.Size = new System.Drawing.Size(592, 224);
            this.rtbConsole.TabIndex = 10;
            this.rtbConsole.Text = "";
            this.rtbConsole.TextChanged += new System.EventHandler(this.rtbConsole_TextChanged);
            // 
            // MainFrame
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(619, 365);
            this.Controls.Add(this.rtbConsole);
            this.Controls.Add(this.tbXML);
            this.Controls.Add(this.bChooseXML);
            this.Controls.Add(this.bStart);
            this.DoubleBuffered = true;
            this.Name = "MainFrame";
            this.Text = "EBOM_Creation_Tool v1.3";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainFrame_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button bStart;
        private System.Windows.Forms.Button bChooseXML;
        private System.Windows.Forms.TextBox tbXML;
        private System.Windows.Forms.RichTextBox rtbConsole;
    }
}

