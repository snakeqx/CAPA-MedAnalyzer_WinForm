using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelHandler;
using System.Windows.Forms.DataVisualization.Charting;

namespace CAPA_MedAnalyzer
{
    public partial class Form1 : Form
    {
        private CapaSsme Ncm;
        private string CurrentPath;
        private string ConfigFilePath;
        private string InputFileName;
        private DataTable TableFpy;
        private DataTable TableProductNcmDetail;
        private DataTable TableMaterialDetail;
        private DataTable TableNcmWeeklyGeneral;
        private DataTable TableNcmWeeklyDetail;
        private string HeaderChartFPy = "FPY";
        private ExcelHeaderConfig RawExcelHeader = ExcelHeaderConfig.GetInstance();

        #region Initilization
        public Form1()
        {  
            InitializeComponent();

            CurrentPath = Application.StartupPath;
            ConfigFilePath = CurrentPath + @"/Config.xml";

            TbLog.Text = "Log Start!\r\n";
            DtpWeeklyNcm.Value = DateTime.Now.AddDays(-8);
            TbMonth.Text = DateTime.Now.AddMonths(-1).Month.ToString();
            TbYear.Text = DateTime.Now.AddMonths(-1).Year.ToString();
            BtnCalculateFPY.Enabled = false;
            exportExcelToolStripMenuItem.Enabled = false;
            BtnCheckWeeklyNcm.Enabled = false;
        }
        #endregion

        #region Event Handling
        private void openExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Open Excel file";
            dialog.ShowHelp = true;
            dialog.Filter = "Excel files|*.xlsx";
            dialog.FilterIndex = 1;
            dialog.RestoreDirectory = false;
            dialog.InitialDirectory = CurrentPath;
            dialog.Multiselect = false;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                InputFileName = dialog.FileName;
                Ncm = new CapaSsme(InputFileName, 1, ConfigFilePath);
                TbLog.Text += dialog.FileName + "Inited!\r\n" + "Data is init:" + Ncm.IsInitDone.ToString() + "\r\n";
            }
            else return;

            if (!Ncm.IsInitDone)
            {
                if (Ncm.ListofNewFoundName.Count == 0)
                    TbLog.Text += "Init failed, double check the excel file!\r\n";
                else
                {
                    TbLog.Text += "Find new names: \r\n";
                    foreach (string s in Ncm.ListofNewFoundName)
                    {
                        TbLog.Text += s + "\r\n";
                        Ncm.ConfigXml.AddData(s, "");
                    }
                }
            }
            // Binding the data to show
            DgvConfig.DataSource = Ncm.ConfigXml.DictToDataTable();
            DgvConfig.AutoResizeColumns();
            // Enable Analyze button and export excel menu
            BtnCalculateFPY.Enabled = true;
            exportExcelToolStripMenuItem.Enabled = true;
            BtnCheckWeeklyNcm.Enabled = true;
            excelHeaderConfigurationToolStripMenuItem.Enabled = false;
        }

        private void exportExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "save Excel file";
            dialog.ShowHelp = true;
            dialog.Filter = "Excel files|*.xlsx";
            dialog.FileName = InputFileName.Split('.')[0] + "_Modi." + InputFileName.Split('.')[1];
            dialog.FilterIndex = 1;
            dialog.RestoreDirectory = false;
            dialog.InitialDirectory = CurrentPath;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string filename = dialog.FileName;
                Ncm.InitExcel(filename);
                Ncm.AddtoExcel(Ncm.RawDataTable, "Rawdata");
                Ncm.AddtoExcel(Ncm.NcmDataTable.Table, "Modified Data");
                Ncm.SaveExcel();
            }
            else return;

        }

        private void BtnCalculateFPY_Click(object sender, EventArgs e)
        {
            DataTable dt = null;
            try
            {
                if (!Ncm.MatchProductName())
                {
                    MessageBox.Show("Table is still not initialized, Maybe there is some System has no Union Name.");
                    return;
                }
                int year = int.Parse(TbYear.Text);
                int month = int.Parse(TbMonth.Text);
                //ShowFpyChart(Ncm.AnalyzeFPY4SingleMonth(year, month));
                dt = Ncm.AnalyzeFPY4SingleMonth(year, month);
                TableFpy = dt.Copy();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + " For Year and Month");
            }

            
            DgvFpy.DataSource = TableFpy;
            DgvFpy.AutoResizeColumns();

            ChartFpy.BackColor = Color.Azure;
            ChartFpy.Series[HeaderChartFPy].Points.DataBindXY(GetFpyData().Item1, GetFpyData().Item2);
            ChartFpy.Series[HeaderChartFPy].ChartType = SeriesChartType.Bar;
            TabControlMain.SelectTab(1);
        }

        private void BtnProductNcmDetail_Click(object sender, EventArgs e)
        {
            if(DgvFpy.Rows.GetRowCount(DataGridViewElementStates.Selected) == 0)
            {
                MessageBox.Show("Please select a row");
                return;
            }
            // To get a selected row number, below code is not working perfectly
            // int selectedRow = DgvFpy.SelectedRows[0].Index;
            // Because if the datagridview is resorted after click on column name, the index will be different
            // Use below 2 lines to get the correct index
            // line 1: to get the whole row data from the datagridview
            // line 2: search the line in original datatable to get the correct index
            int year;
            int month;
            try
            {
                year = int.Parse(TbYear.Text);
                month = int.Parse(TbMonth.Text);
            }
            catch (Exception ex)
            {
                throw (ex);
            }

            DataRowView drv = DgvFpy.SelectedRows[0].DataBoundItem as DataRowView;
            int selectedRow = TableFpy.Rows.IndexOf(drv.Row);
            TableProductNcmDetail = Ncm.GetProductNcmDetail(TableFpy.Rows[selectedRow][Ncm.ExHdSystemUnionName].ToString(),
                                                            year, month).Copy();
            DgvProductNcmDetail.DataSource = TableProductNcmDetail;
            DgvProductNcmDetail.AutoResizeColumns();
            TabControlMain.SelectTab(2);
        }

        private void BtnCheckMaterial_Click(object sender, EventArgs e)
        {
            if (DgvProductNcmDetail.Rows.GetRowCount(DataGridViewElementStates.Selected) == 0)
            {
                MessageBox.Show("Please select a row");
                return;
            }
            // To get a selected row number, below code is not working perfectly
            // int selectedRow = DgvProductNcmDetail.SelectedRows[0].Index;
            // Because if the datagridview is resorted after click on column name, the index will be different
            // Use below 2 lines to get the correct index
            // line 1: to get the whole row data from the datagridview
            // line 2: search the line in original datatable to get the correct index
            DataRowView drv = DgvProductNcmDetail.SelectedRows[0].DataBoundItem as DataRowView;
            int selectedRow = TableProductNcmDetail.Rows.IndexOf(drv.Row);
            TableMaterialDetail = Ncm.GetMaterialDetail(TableProductNcmDetail.Rows[selectedRow][Ncm.RawExcelHeader.DefectivePartDescription].ToString()).Copy();
            DgvMaterialDetail.DataSource = TableMaterialDetail;
            DgvMaterialDetail.AutoResizeColumns();
            TabControlMain.SelectTab(3);
            
        }

        private void BtnCheckWeeklyNcm_Click(object sender, EventArgs e)
        {
            try
            {
                Ncm.MatchProductName();
                DateTime timeSpliter = DtpWeeklyNcm.Value.Date;
                Ncm.RecaclulateWeekSpliter(timeSpliter);
                TableNcmWeeklyGeneral = Ncm.GetWeeklyNcmGeneral();
                DgvNcmWeeklyGeneral.DataSource = TableNcmWeeklyGeneral;
                DgvNcmWeeklyGeneral.AutoResizeColumns();
                TabControlMain.SelectTab(4);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " For Year and Month");
            }

        }

        private void DgvNcmWeeklyGeneral_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0)
            {
                MessageBox.Show("Please double click on row selection area.");
                return;
            }
            // To get a selected row number, below code is not working perfectly
            // int selectedRow = DgvFpy.SelectedRows[0].Index;
            // Because if the datagridview is resorted after click on column name, the index will be different
            // Use below 2 lines to get the correct index
            // line 1: to get the whole row data from the datagridview
            // line 2: search the line in original datatable to get the correct index
            DataRowView drv = DgvNcmWeeklyGeneral.SelectedRows[0].DataBoundItem as DataRowView;
            int selectedRow = TableNcmWeeklyGeneral.Rows.IndexOf(drv.Row);
            TableNcmWeeklyDetail = Ncm.GetNcmWeeklyDetail(TableNcmWeeklyGeneral.Rows[selectedRow][Ncm.ExHdSystemUnionName].ToString()).Copy();
            DgvNcmWeeklyDetail.DataSource = TableNcmWeeklyDetail;
            DgvNcmWeeklyDetail.AutoResizeColumns();
        }

        private void excelHeaderConfigurationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 fm = new Form2();
            fm.ShowDialog();
            
        }

        #region Deal With Click On DataGridView
        private void DgwConfig_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Ncm.ConfigXml.UpdateDictByDataTable(DgvConfig.DataSource as DataTable);
            Ncm.ConfigXml.SaveXml(ConfigFilePath);
            TbLog.Text += "config.xml updated!\r\n";
        }

        private void DgwConfig_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            // Not allowed to change the 1st column
            if (e.ColumnIndex == 0 ||
                e.ColumnIndex == 1)
                e.Cancel = true;
        }

        private void DgwFpy_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataRow r in TableFpy.Rows)
            {
                try
                {
                    r[Ncm.HdFpy] = Math.Round(double.Parse(r[Ncm.HdQtyWithNcm].ToString()) /
                                              double.Parse(r[Ncm.HdTotalSysQty].ToString()),
                                              4);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            ChartFpy.Series[HeaderChartFPy].Points.DataBindXY(GetFpyData().Item1, GetFpyData().Item2);
        }

        private void DgwFpy_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            // Not allowed to change the 1st and last column 
            if (e.ColumnIndex == 0 ||
                e.ColumnIndex == 1 ||
                e.ColumnIndex == 2)
                e.Cancel = true;
        }

        private void DgvProductNcmDetail_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            e.Cancel = true;
        }

        private void DgvMaterialDetail_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            e.Cancel = true;
        }
        #endregion
        #endregion



        #region Helper functions
        private Tuple<List<string>, List<double>> GetFpyData()
        {
            List<string> chartName = new List<string>();
            List<double> chartFpy = new List<double>();

            foreach (DataRow r in TableFpy.Rows)
            {
                try
                {
                    chartName.Add(r[Ncm.ExHdSystemUnionName].ToString());
                    chartFpy.Add(double.Parse(r[Ncm.HdFpy].ToString()));
                }
                catch (Exception ex)
                {
                    chartName.Add("Error");
                    chartFpy.Add(1.00);
                    MessageBox.Show(ex.Message.ToString());
                }

            }
            Tuple<List<string>, List<double>> result = new Tuple<List<string>, List<double>>(chartName, chartFpy);
            return result;
        }
        #endregion

        


    }// end of class
}// end of name space
