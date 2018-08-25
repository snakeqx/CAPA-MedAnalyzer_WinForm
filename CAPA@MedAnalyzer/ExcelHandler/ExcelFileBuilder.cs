using System;
using System.Data;
using OfficeOpenXml;
using System.IO;

namespace ExcelHandler
{
    class ExcelFileBuilder
    {
        public DataSet ToExcelDataSet = new DataSet();
        private string ExcelFileName;

        /// <summary>
        /// Initialize the class with a file name
        /// </summary>
        /// <param name="filePath"></param>
        public ExcelFileBuilder(string filePath)
        {
            ExcelFileName = filePath;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataTable"></param>
        public void AddDataTable(DataTable dataTable)
        {
            try
            { ToExcelDataSet.Tables.Add(dataTable); }
            catch (Exception ex) { throw (ex); }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="forceDelete"></param>
        /// <returns></returns>
        public bool ToExcel(bool forceDelete = true)
        {
            using (ExcelPackage objExcelPackage = new ExcelPackage())
            {
                foreach (DataTable dtSrc in ToExcelDataSet.Tables)
                {
                    //Create the worksheet    
                    ExcelWorksheet objWorksheet = objExcelPackage.Workbook.Worksheets.Add(dtSrc.TableName);
                    //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1    
                    objWorksheet.Cells["A1"].LoadFromDataTable(dtSrc, true);
                }

                //Write it back to the client
                if (File.Exists(ExcelFileName))
                {
                    if (forceDelete) File.Delete(ExcelFileName);
                    else return false;
                }

                //Create excel file on physical disk    
                FileStream objFileStrm = File.Create(ExcelFileName);
                objFileStrm.Close();

                //Write content to excel file    
                File.WriteAllBytes(ExcelFileName, objExcelPackage.GetAsByteArray());
            }
            return true;
        }

    }


}
