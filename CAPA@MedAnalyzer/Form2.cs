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

namespace CAPA_MedAnalyzer
{
    public partial class Form2 : Form
    {
        ExcelHeaderConfig RawExcelHeader = ExcelHeaderConfig.GetInstance();
        private string CurrentPath;

        public Form2()
        {
            CurrentPath = Application.StartupPath;
            InitializeComponent();
            DgvExcelHenderConfig.DataSource = RawExcelHeader.GetFields();
            DgvExcelHenderConfig.AutoResizeColumns();
        }

        private void DgvExcelHenderConfig_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.ColumnIndex == 0) e.Cancel=true;
        }

        private void saveConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "save header config file";
            dialog.ShowHelp = true;
            dialog.Filter = "Header Config Files|*.hdconfig";
            dialog.FilterIndex = 1;
            dialog.RestoreDirectory = false;
            dialog.InitialDirectory = CurrentPath;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string filename = dialog.FileName;
                RawExcelHeader.SelfSerialized(filename);
            }
            else return;
        }

        private void readConfigToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "open header config file";
            dialog.ShowHelp = true;
            dialog.Filter = "Header Config Files|*.hdconfig";
            dialog.FilterIndex = 1;
            dialog.RestoreDirectory = false;
            dialog.InitialDirectory = CurrentPath;
            dialog.Multiselect = false;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string filename = dialog.FileName;
                RawExcelHeader.SelfDeSerialized(filename);
            }
            else return;
            DgvExcelHenderConfig.DataSource = RawExcelHeader.GetFields();
            DgvExcelHenderConfig.AutoResizeColumns();
        }
    }
}
