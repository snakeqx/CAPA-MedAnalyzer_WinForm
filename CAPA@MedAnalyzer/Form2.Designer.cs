namespace CAPA_MedAnalyzer
{
    partial class Form2
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
            this.DgvExcelHenderConfig = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveConfigToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.readConfigToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.DgvExcelHenderConfig)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // DgvExcelHenderConfig
            // 
            this.DgvExcelHenderConfig.AllowUserToAddRows = false;
            this.DgvExcelHenderConfig.AllowUserToDeleteRows = false;
            this.DgvExcelHenderConfig.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvExcelHenderConfig.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DgvExcelHenderConfig.Location = new System.Drawing.Point(0, 25);
            this.DgvExcelHenderConfig.Name = "DgvExcelHenderConfig";
            this.DgvExcelHenderConfig.RowTemplate.Height = 23;
            this.DgvExcelHenderConfig.Size = new System.Drawing.Size(800, 425);
            this.DgvExcelHenderConfig.TabIndex = 0;
            this.DgvExcelHenderConfig.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.DgvExcelHenderConfig_CellBeginEdit);
            this.DgvExcelHenderConfig.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.DgvExcelHenderConfig_CellEndEdit);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 25);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveConfigToolStripMenuItem,
            this.readConfigToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(39, 21);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // saveConfigToolStripMenuItem
            // 
            this.saveConfigToolStripMenuItem.Name = "saveConfigToolStripMenuItem";
            this.saveConfigToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.saveConfigToolStripMenuItem.Text = "Save Config";
            this.saveConfigToolStripMenuItem.Click += new System.EventHandler(this.saveConfigToolStripMenuItem_Click);
            // 
            // readConfigToolStripMenuItem
            // 
            this.readConfigToolStripMenuItem.Name = "readConfigToolStripMenuItem";
            this.readConfigToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.readConfigToolStripMenuItem.Text = "Read Config";
            this.readConfigToolStripMenuItem.Click += new System.EventHandler(this.readConfigToolStripMenuItem_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.DgvExcelHenderConfig);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form2";
            this.Text = "Form2";
            ((System.ComponentModel.ISupportInitialize)(this.DgvExcelHenderConfig)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView DgvExcelHenderConfig;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem readConfigToolStripMenuItem;
        protected internal System.Windows.Forms.ToolStripMenuItem saveConfigToolStripMenuItem;
    }
}