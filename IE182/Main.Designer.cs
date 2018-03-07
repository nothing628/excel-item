namespace IE182
{
    partial class Main
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
            this.gridMain = new unvell.ReoGrid.ReoGridControl();
            this.textBoxFileName = new System.Windows.Forms.TextBox();
            this.buttonChooseFile = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.buttonStoreToDB = new System.Windows.Forms.Button();
            this.buttonProcess = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // gridMain
            // 
            this.gridMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gridMain.BackColor = System.Drawing.Color.White;
            this.gridMain.ColumnHeaderContextMenuStrip = null;
            this.gridMain.LeadHeaderContextMenuStrip = null;
            this.gridMain.Location = new System.Drawing.Point(1, 62);
            this.gridMain.Name = "gridMain";
            this.gridMain.RowHeaderContextMenuStrip = null;
            this.gridMain.Script = null;
            this.gridMain.SheetTabContextMenuStrip = null;
            this.gridMain.SheetTabNewButtonVisible = true;
            this.gridMain.SheetTabVisible = true;
            this.gridMain.SheetTabWidth = 300;
            this.gridMain.ShowScrollEndSpacing = true;
            this.gridMain.Size = new System.Drawing.Size(983, 397);
            this.gridMain.TabIndex = 0;
            // 
            // textBoxFileName
            // 
            this.textBoxFileName.Enabled = false;
            this.textBoxFileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxFileName.Location = new System.Drawing.Point(12, 12);
            this.textBoxFileName.Multiline = true;
            this.textBoxFileName.Name = "textBoxFileName";
            this.textBoxFileName.Size = new System.Drawing.Size(517, 37);
            this.textBoxFileName.TabIndex = 1;
            // 
            // buttonChooseFile
            // 
            this.buttonChooseFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonChooseFile.Location = new System.Drawing.Point(546, 12);
            this.buttonChooseFile.Name = "buttonChooseFile";
            this.buttonChooseFile.Size = new System.Drawing.Size(57, 37);
            this.buttonChooseFile.TabIndex = 2;
            this.buttonChooseFile.Text = "...";
            this.buttonChooseFile.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.buttonChooseFile.UseVisualStyleBackColor = true;
            this.buttonChooseFile.Click += new System.EventHandler(this.buttonChooseFile_Click);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(876, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(96, 37);
            this.button1.TabIndex = 3;
            this.button1.Text = "Setting";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // buttonStoreToDB
            // 
            this.buttonStoreToDB.Enabled = false;
            this.buttonStoreToDB.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonStoreToDB.Location = new System.Drawing.Point(623, 12);
            this.buttonStoreToDB.Name = "buttonStoreToDB";
            this.buttonStoreToDB.Size = new System.Drawing.Size(123, 37);
            this.buttonStoreToDB.TabIndex = 4;
            this.buttonStoreToDB.Text = "Store To DB";
            this.buttonStoreToDB.UseVisualStyleBackColor = true;
            this.buttonStoreToDB.Click += new System.EventHandler(this.buttonStoreToDB_Click);
            // 
            // buttonProcess
            // 
            this.buttonProcess.Enabled = false;
            this.buttonProcess.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonProcess.Location = new System.Drawing.Point(752, 12);
            this.buttonProcess.Name = "buttonProcess";
            this.buttonProcess.Size = new System.Drawing.Size(89, 37);
            this.buttonProcess.TabIndex = 5;
            this.buttonProcess.Text = "Process";
            this.buttonProcess.UseVisualStyleBackColor = true;
            this.buttonProcess.Click += new System.EventHandler(this.buttonProcess_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 461);
            this.Controls.Add(this.buttonProcess);
            this.Controls.Add(this.buttonStoreToDB);
            this.Controls.Add(this.buttonChooseFile);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBoxFileName);
            this.Controls.Add(this.gridMain);
            this.Name = "Main";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Main";
            this.Load += new System.EventHandler(this.Main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private unvell.ReoGrid.ReoGridControl gridMain;
        private System.Windows.Forms.TextBox textBoxFileName;
        private System.Windows.Forms.Button buttonChooseFile;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button buttonStoreToDB;
        private System.Windows.Forms.Button buttonProcess;
    }
}

