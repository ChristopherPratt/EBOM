namespace EBOM_Creation_Tool_v2
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
            this.bChooseSource = new System.Windows.Forms.Button();
            this.bStart = new System.Windows.Forms.Button();
            this.lblFileName = new System.Windows.Forms.Label();
            this.rtbConsole = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // bChooseSource
            // 
            this.bChooseSource.Location = new System.Drawing.Point(12, 10);
            this.bChooseSource.Name = "bChooseSource";
            this.bChooseSource.Size = new System.Drawing.Size(145, 40);
            this.bChooseSource.TabIndex = 0;
            this.bChooseSource.Text = "Choose source XML file";
            this.bChooseSource.UseVisualStyleBackColor = true;
            this.bChooseSource.Click += new System.EventHandler(this.bChooseSource_Click);
            // 
            // bStart
            // 
            this.bStart.Location = new System.Drawing.Point(240, 55);
            this.bStart.Name = "bStart";
            this.bStart.Size = new System.Drawing.Size(145, 40);
            this.bStart.TabIndex = 1;
            this.bStart.Text = "Start";
            this.bStart.UseVisualStyleBackColor = true;
            this.bStart.Click += new System.EventHandler(this.bStart_Click);
            // 
            // lblFileName
            // 
            this.lblFileName.AutoSize = true;
            this.lblFileName.Location = new System.Drawing.Point(163, 24);
            this.lblFileName.Name = "lblFileName";
            this.lblFileName.Size = new System.Drawing.Size(0, 13);
            this.lblFileName.TabIndex = 2;
            this.lblFileName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // rtbConsole
            // 
            this.rtbConsole.Location = new System.Drawing.Point(14, 114);
            this.rtbConsole.Name = "rtbConsole";
            this.rtbConsole.Size = new System.Drawing.Size(622, 249);
            this.rtbConsole.TabIndex = 3;
            this.rtbConsole.Text = "";
            // 
            // mainFrameScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(652, 376);
            this.Controls.Add(this.rtbConsole);
            this.Controls.Add(this.lblFileName);
            this.Controls.Add(this.bStart);
            this.Controls.Add(this.bChooseSource);
            this.Name = "mainFrameScreen";
            this.Text = "EBOM Creation Tool v2.0";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bChooseSource;
        private System.Windows.Forms.Button bStart;
        private System.Windows.Forms.Label lblFileName;
        private System.Windows.Forms.RichTextBox rtbConsole;
    }
}

