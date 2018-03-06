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
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.buttonExport = new System.Windows.Forms.Button();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
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
            this.gridMain.Size = new System.Drawing.Size(1349, 665);
            this.gridMain.TabIndex = 0;
            // 
            // textBoxFileName
            // 
            this.textBoxFileName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxFileName.Enabled = false;
            this.textBoxFileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxFileName.Location = new System.Drawing.Point(404, 12);
            this.textBoxFileName.Multiline = true;
            this.textBoxFileName.Name = "textBoxFileName";
            this.textBoxFileName.Size = new System.Drawing.Size(517, 37);
            this.textBoxFileName.TabIndex = 1;
            // 
            // buttonChooseFile
            // 
            this.buttonChooseFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonChooseFile.Location = new System.Drawing.Point(937, 12);
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
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(1242, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(96, 37);
            this.button1.TabIndex = 3;
            this.button1.Text = "Setting";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "Excel files|*.xlsx;*.xls";
            // 
            // buttonExport
            // 
            this.buttonExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonExport.Location = new System.Drawing.Point(1020, 12);
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.Size = new System.Drawing.Size(87, 37);
            this.buttonExport.TabIndex = 4;
            this.buttonExport.Text = "Export";
            this.buttonExport.UseVisualStyleBackColor = true;
            this.buttonExport.Click += new System.EventHandler(this.buttonExport_Click);
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.DefaultExt = "xlsx";
            this.saveFileDialog.FileName = "IE128_generated.xlsx";
            this.saveFileDialog.Filter = "Excel files|*.xlsx;*.xls";
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1350, 729);
            this.Controls.Add(this.buttonExport);
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
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button buttonExport;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
    }
}

