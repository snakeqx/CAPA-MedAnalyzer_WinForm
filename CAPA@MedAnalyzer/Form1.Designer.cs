namespace CAPA_MedAnalyzer
{
    partial class Form1
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea3 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend3 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DgvConfig = new System.Windows.Forms.DataGridView();
            this.TbLog = new System.Windows.Forms.TextBox();
            this.TabControlMain = new System.Windows.Forms.TabControl();
            this.TabInputAndConfig = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.BtnCheckWeeklyNcm = new System.Windows.Forms.Button();
            this.DtpWeeklyNcm = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnCalculateFPY = new System.Windows.Forms.Button();
            this.TbMonth = new System.Windows.Forms.TextBox();
            this.TbYear = new System.Windows.Forms.TextBox();
            this.TabChartFPY = new System.Windows.Forms.TabPage();
            this.panel5 = new System.Windows.Forms.Panel();
            this.ChartFpy = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.panel4 = new System.Windows.Forms.Panel();
            this.DgvFpy = new System.Windows.Forms.DataGridView();
            this.BtnProductNcmDetail = new System.Windows.Forms.Button();
            this.TabDetail = new System.Windows.Forms.TabPage();
            this.DgvProductNcmDetail = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.BtnCheckMaterial = new System.Windows.Forms.Button();
            this.TabMaterialDetail = new System.Windows.Forms.TabPage();
            this.DgvMaterialDetail = new System.Windows.Forms.DataGridView();
            this.TabWeekly = new System.Windows.Forms.TabPage();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.DgvNcmWeeklyGeneral = new System.Windows.Forms.DataGridView();
            this.DgvNcmWeeklyDetail = new System.Windows.Forms.DataGridView();
            this.configToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelHeaderConfigurationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DgvConfig)).BeginInit();
            this.TabControlMain.SuspendLayout();
            this.TabInputAndConfig.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.TabChartFPY.SuspendLayout();
            this.panel5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ChartFpy)).BeginInit();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DgvFpy)).BeginInit();
            this.TabDetail.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DgvProductNcmDetail)).BeginInit();
            this.panel1.SuspendLayout();
            this.TabMaterialDetail.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DgvMaterialDetail)).BeginInit();
            this.TabWeekly.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DgvNcmWeeklyGeneral)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DgvNcmWeeklyDetail)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.configToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openExcelToolStripMenuItem,
            this.exportExcelToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // openExcelToolStripMenuItem
            // 
            this.openExcelToolStripMenuItem.Name = "openExcelToolStripMenuItem";
            this.openExcelToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.openExcelToolStripMenuItem.Text = "Open Excel";
            this.openExcelToolStripMenuItem.Click += new System.EventHandler(this.openExcelToolStripMenuItem_Click);
            // 
            // exportExcelToolStripMenuItem
            // 
            this.exportExcelToolStripMenuItem.Name = "exportExcelToolStripMenuItem";
            this.exportExcelToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.exportExcelToolStripMenuItem.Text = "Export Excel";
            this.exportExcelToolStripMenuItem.Click += new System.EventHandler(this.exportExcelToolStripMenuItem_Click);
            // 
            // DgvConfig
            // 
            this.DgvConfig.AllowUserToAddRows = false;
            this.DgvConfig.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DgvConfig.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvConfig.Location = new System.Drawing.Point(0, 6);
            this.DgvConfig.Name = "DgvConfig";
            this.DgvConfig.Size = new System.Drawing.Size(385, 391);
            this.DgvConfig.TabIndex = 1;
            this.DgvConfig.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.DgwConfig_CellBeginEdit);
            this.DgvConfig.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.DgwConfig_CellEndEdit);
            // 
            // TbLog
            // 
            this.TbLog.Dock = System.Windows.Forms.DockStyle.Top;
            this.TbLog.Location = new System.Drawing.Point(0, 0);
            this.TbLog.Multiline = true;
            this.TbLog.Name = "TbLog";
            this.TbLog.Size = new System.Drawing.Size(398, 244);
            this.TbLog.TabIndex = 2;
            // 
            // TabControlMain
            // 
            this.TabControlMain.Controls.Add(this.TabInputAndConfig);
            this.TabControlMain.Controls.Add(this.TabChartFPY);
            this.TabControlMain.Controls.Add(this.TabDetail);
            this.TabControlMain.Controls.Add(this.TabMaterialDetail);
            this.TabControlMain.Controls.Add(this.TabWeekly);
            this.TabControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TabControlMain.Location = new System.Drawing.Point(0, 24);
            this.TabControlMain.Name = "TabControlMain";
            this.TabControlMain.SelectedIndex = 0;
            this.TabControlMain.Size = new System.Drawing.Size(800, 426);
            this.TabControlMain.TabIndex = 3;
            // 
            // TabInputAndConfig
            // 
            this.TabInputAndConfig.Controls.Add(this.panel2);
            this.TabInputAndConfig.Controls.Add(this.DgvConfig);
            this.TabInputAndConfig.Location = new System.Drawing.Point(4, 22);
            this.TabInputAndConfig.Name = "TabInputAndConfig";
            this.TabInputAndConfig.Padding = new System.Windows.Forms.Padding(3);
            this.TabInputAndConfig.Size = new System.Drawing.Size(792, 400);
            this.TabInputAndConfig.TabIndex = 0;
            this.TabInputAndConfig.Text = "Input and Config";
            this.TabInputAndConfig.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Controls.Add(this.TbLog);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel2.Location = new System.Drawing.Point(391, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(398, 394);
            this.panel2.TabIndex = 5;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.BtnCheckWeeklyNcm);
            this.panel3.Controls.Add(this.DtpWeeklyNcm);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Controls.Add(this.label1);
            this.panel3.Controls.Add(this.BtnCalculateFPY);
            this.panel3.Controls.Add(this.TbMonth);
            this.panel3.Controls.Add(this.TbYear);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel3.Location = new System.Drawing.Point(0, 242);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(398, 152);
            this.panel3.TabIndex = 5;
            // 
            // BtnCheckWeeklyNcm
            // 
            this.BtnCheckWeeklyNcm.Location = new System.Drawing.Point(239, 81);
            this.BtnCheckWeeklyNcm.Name = "BtnCheckWeeklyNcm";
            this.BtnCheckWeeklyNcm.Size = new System.Drawing.Size(105, 49);
            this.BtnCheckWeeklyNcm.TabIndex = 8;
            this.BtnCheckWeeklyNcm.Text = "Check Weekly NCM";
            this.BtnCheckWeeklyNcm.UseVisualStyleBackColor = true;
            this.BtnCheckWeeklyNcm.Click += new System.EventHandler(this.BtnCheckWeeklyNcm_Click);
            // 
            // DtpWeeklyNcm
            // 
            this.DtpWeeklyNcm.Location = new System.Drawing.Point(27, 98);
            this.DtpWeeklyNcm.Name = "DtpWeeklyNcm";
            this.DtpWeeklyNcm.Size = new System.Drawing.Size(200, 20);
            this.DtpWeeklyNcm.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(24, 37);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Month";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Year";
            // 
            // BtnCalculateFPY
            // 
            this.BtnCalculateFPY.Location = new System.Drawing.Point(187, 8);
            this.BtnCalculateFPY.Name = "BtnCalculateFPY";
            this.BtnCalculateFPY.Size = new System.Drawing.Size(157, 46);
            this.BtnCalculateFPY.TabIndex = 4;
            this.BtnCalculateFPY.Text = "Calculate FPY";
            this.BtnCalculateFPY.UseVisualStyleBackColor = true;
            this.BtnCalculateFPY.Click += new System.EventHandler(this.BtnCalculateFPY_Click);
            // 
            // TbMonth
            // 
            this.TbMonth.Location = new System.Drawing.Point(59, 8);
            this.TbMonth.Name = "TbMonth";
            this.TbMonth.Size = new System.Drawing.Size(100, 20);
            this.TbMonth.TabIndex = 3;
            // 
            // TbYear
            // 
            this.TbYear.Location = new System.Drawing.Point(59, 34);
            this.TbYear.Name = "TbYear";
            this.TbYear.Size = new System.Drawing.Size(100, 20);
            this.TbYear.TabIndex = 3;
            // 
            // TabChartFPY
            // 
            this.TabChartFPY.Controls.Add(this.panel5);
            this.TabChartFPY.Controls.Add(this.panel4);
            this.TabChartFPY.Location = new System.Drawing.Point(4, 22);
            this.TabChartFPY.Name = "TabChartFPY";
            this.TabChartFPY.Padding = new System.Windows.Forms.Padding(3);
            this.TabChartFPY.Size = new System.Drawing.Size(792, 400);
            this.TabChartFPY.TabIndex = 1;
            this.TabChartFPY.Text = "Chart FPY";
            this.TabChartFPY.UseVisualStyleBackColor = true;
            // 
            // panel5
            // 
            this.panel5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel5.Controls.Add(this.ChartFpy);
            this.panel5.Location = new System.Drawing.Point(3, 3);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(341, 394);
            this.panel5.TabIndex = 4;
            // 
            // ChartFpy
            // 
            this.ChartFpy.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            chartArea3.Name = "ChartArea1";
            this.ChartFpy.ChartAreas.Add(chartArea3);
            legend3.Name = "Legend1";
            this.ChartFpy.Legends.Add(legend3);
            this.ChartFpy.Location = new System.Drawing.Point(0, 0);
            this.ChartFpy.Name = "ChartFpy";
            series3.ChartArea = "ChartArea1";
            series3.Legend = "Legend1";
            series3.Name = "FPY";
            this.ChartFpy.Series.Add(series3);
            this.ChartFpy.Size = new System.Drawing.Size(341, 394);
            this.ChartFpy.TabIndex = 0;
            this.ChartFpy.Text = "First Passing Yield";
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.DgvFpy);
            this.panel4.Controls.Add(this.BtnProductNcmDetail);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel4.Location = new System.Drawing.Point(350, 3);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(439, 394);
            this.panel4.TabIndex = 3;
            // 
            // DgvFpy
            // 
            this.DgvFpy.AllowUserToAddRows = false;
            this.DgvFpy.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvFpy.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DgvFpy.Location = new System.Drawing.Point(0, 0);
            this.DgvFpy.MultiSelect = false;
            this.DgvFpy.Name = "DgvFpy";
            this.DgvFpy.Size = new System.Drawing.Size(439, 331);
            this.DgvFpy.TabIndex = 1;
            this.DgvFpy.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.DgwFpy_CellBeginEdit);
            this.DgvFpy.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.DgwFpy_CellEndEdit);
            // 
            // BtnProductNcmDetail
            // 
            this.BtnProductNcmDetail.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.BtnProductNcmDetail.Location = new System.Drawing.Point(0, 331);
            this.BtnProductNcmDetail.Name = "BtnProductNcmDetail";
            this.BtnProductNcmDetail.Size = new System.Drawing.Size(439, 63);
            this.BtnProductNcmDetail.TabIndex = 2;
            this.BtnProductNcmDetail.Text = "See NCM detail in Product";
            this.BtnProductNcmDetail.UseVisualStyleBackColor = true;
            this.BtnProductNcmDetail.Click += new System.EventHandler(this.BtnProductNcmDetail_Click);
            // 
            // TabDetail
            // 
            this.TabDetail.Controls.Add(this.DgvProductNcmDetail);
            this.TabDetail.Controls.Add(this.panel1);
            this.TabDetail.Location = new System.Drawing.Point(4, 22);
            this.TabDetail.Name = "TabDetail";
            this.TabDetail.Padding = new System.Windows.Forms.Padding(3);
            this.TabDetail.Size = new System.Drawing.Size(792, 400);
            this.TabDetail.TabIndex = 2;
            this.TabDetail.Text = "NCM Detail";
            this.TabDetail.UseVisualStyleBackColor = true;
            // 
            // DgvProductNcmDetail
            // 
            this.DgvProductNcmDetail.AllowUserToAddRows = false;
            this.DgvProductNcmDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvProductNcmDetail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DgvProductNcmDetail.Location = new System.Drawing.Point(3, 60);
            this.DgvProductNcmDetail.Name = "DgvProductNcmDetail";
            this.DgvProductNcmDetail.Size = new System.Drawing.Size(786, 337);
            this.DgvProductNcmDetail.TabIndex = 1;
            this.DgvProductNcmDetail.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.DgvProductNcmDetail_CellBeginEdit);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.BtnCheckMaterial);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(786, 57);
            this.panel1.TabIndex = 0;
            // 
            // BtnCheckMaterial
            // 
            this.BtnCheckMaterial.Dock = System.Windows.Forms.DockStyle.Right;
            this.BtnCheckMaterial.Location = new System.Drawing.Point(676, 0);
            this.BtnCheckMaterial.Name = "BtnCheckMaterial";
            this.BtnCheckMaterial.Size = new System.Drawing.Size(110, 57);
            this.BtnCheckMaterial.TabIndex = 0;
            this.BtnCheckMaterial.Text = "Check Material History";
            this.BtnCheckMaterial.UseVisualStyleBackColor = true;
            this.BtnCheckMaterial.Click += new System.EventHandler(this.BtnCheckMaterial_Click);
            // 
            // TabMaterialDetail
            // 
            this.TabMaterialDetail.Controls.Add(this.DgvMaterialDetail);
            this.TabMaterialDetail.Location = new System.Drawing.Point(4, 22);
            this.TabMaterialDetail.Name = "TabMaterialDetail";
            this.TabMaterialDetail.Padding = new System.Windows.Forms.Padding(3);
            this.TabMaterialDetail.Size = new System.Drawing.Size(792, 400);
            this.TabMaterialDetail.TabIndex = 3;
            this.TabMaterialDetail.Text = "Material Detail";
            this.TabMaterialDetail.UseVisualStyleBackColor = true;
            // 
            // DgvMaterialDetail
            // 
            this.DgvMaterialDetail.AllowUserToAddRows = false;
            this.DgvMaterialDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvMaterialDetail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DgvMaterialDetail.Location = new System.Drawing.Point(3, 3);
            this.DgvMaterialDetail.MultiSelect = false;
            this.DgvMaterialDetail.Name = "DgvMaterialDetail";
            this.DgvMaterialDetail.Size = new System.Drawing.Size(786, 394);
            this.DgvMaterialDetail.TabIndex = 0;
            this.DgvMaterialDetail.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.DgvMaterialDetail_CellBeginEdit);
            // 
            // TabWeekly
            // 
            this.TabWeekly.Controls.Add(this.tableLayoutPanel1);
            this.TabWeekly.Location = new System.Drawing.Point(4, 22);
            this.TabWeekly.Name = "TabWeekly";
            this.TabWeekly.Padding = new System.Windows.Forms.Padding(3);
            this.TabWeekly.Size = new System.Drawing.Size(792, 400);
            this.TabWeekly.TabIndex = 4;
            this.TabWeekly.Text = "NCM Weekly";
            this.TabWeekly.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.59033F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 73.40967F));
            this.tableLayoutPanel1.Controls.Add(this.DgvNcmWeeklyGeneral, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.DgvNcmWeeklyDetail, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(786, 394);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // DgvNcmWeeklyGeneral
            // 
            this.DgvNcmWeeklyGeneral.AllowUserToAddRows = false;
            this.DgvNcmWeeklyGeneral.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvNcmWeeklyGeneral.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DgvNcmWeeklyGeneral.Location = new System.Drawing.Point(3, 3);
            this.DgvNcmWeeklyGeneral.Name = "DgvNcmWeeklyGeneral";
            this.DgvNcmWeeklyGeneral.Size = new System.Drawing.Size(202, 388);
            this.DgvNcmWeeklyGeneral.TabIndex = 0;
            this.DgvNcmWeeklyGeneral.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DgvNcmWeeklyGeneral_CellDoubleClick);
            // 
            // DgvNcmWeeklyDetail
            // 
            this.DgvNcmWeeklyDetail.AllowUserToAddRows = false;
            this.DgvNcmWeeklyDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvNcmWeeklyDetail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DgvNcmWeeklyDetail.Location = new System.Drawing.Point(211, 3);
            this.DgvNcmWeeklyDetail.Name = "DgvNcmWeeklyDetail";
            this.DgvNcmWeeklyDetail.Size = new System.Drawing.Size(572, 388);
            this.DgvNcmWeeklyDetail.TabIndex = 1;
            // 
            // configToolStripMenuItem
            // 
            this.configToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.excelHeaderConfigurationToolStripMenuItem});
            this.configToolStripMenuItem.Name = "configToolStripMenuItem";
            this.configToolStripMenuItem.Size = new System.Drawing.Size(55, 20);
            this.configToolStripMenuItem.Text = "Config";
            // 
            // excelHeaderConfigurationToolStripMenuItem
            // 
            this.excelHeaderConfigurationToolStripMenuItem.Name = "excelHeaderConfigurationToolStripMenuItem";
            this.excelHeaderConfigurationToolStripMenuItem.Size = new System.Drawing.Size(218, 22);
            this.excelHeaderConfigurationToolStripMenuItem.Text = "Excel Header Configuration";
            this.excelHeaderConfigurationToolStripMenuItem.Click += new System.EventHandler(this.excelHeaderConfigurationToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.TabControlMain);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Single Month FPY calculator";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DgvConfig)).EndInit();
            this.TabControlMain.ResumeLayout(false);
            this.TabInputAndConfig.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.TabChartFPY.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ChartFpy)).EndInit();
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DgvFpy)).EndInit();
            this.TabDetail.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DgvProductNcmDetail)).EndInit();
            this.panel1.ResumeLayout(false);
            this.TabMaterialDetail.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DgvMaterialDetail)).EndInit();
            this.TabWeekly.ResumeLayout(false);
            this.TabWeekly.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DgvNcmWeeklyGeneral)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DgvNcmWeeklyDetail)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openExcelToolStripMenuItem;
        private System.Windows.Forms.DataGridView DgvConfig;
        private System.Windows.Forms.TextBox TbLog;
        private System.Windows.Forms.TabControl TabControlMain;
        private System.Windows.Forms.TabPage TabInputAndConfig;
        private System.Windows.Forms.TabPage TabChartFPY;
        private System.Windows.Forms.DataVisualization.Charting.Chart ChartFpy;
        private System.Windows.Forms.TextBox TbMonth;
        private System.Windows.Forms.TextBox TbYear;
        private System.Windows.Forms.Button BtnCalculateFPY;
        private System.Windows.Forms.DataGridView DgvFpy;
        private System.Windows.Forms.TabPage TabDetail;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView DgvProductNcmDetail;
        private System.Windows.Forms.Button BtnCheckMaterial;
        private System.Windows.Forms.Button BtnProductNcmDetail;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.ToolStripMenuItem exportExcelToolStripMenuItem;
        private System.Windows.Forms.TabPage TabMaterialDetail;
        private System.Windows.Forms.DataGridView DgvMaterialDetail;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BtnCheckWeeklyNcm;
        private System.Windows.Forms.DateTimePicker DtpWeeklyNcm;
        private System.Windows.Forms.TabPage TabWeekly;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView DgvNcmWeeklyGeneral;
        private System.Windows.Forms.DataGridView DgvNcmWeeklyDetail;
        private System.Windows.Forms.ToolStripMenuItem configToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelHeaderConfigurationToolStripMenuItem;
    }
}