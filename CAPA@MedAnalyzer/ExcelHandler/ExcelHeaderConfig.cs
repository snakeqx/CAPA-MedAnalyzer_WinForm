using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace ExcelHandler
{
    /// <summary>
    /// A single instance class
    /// </summary>
    [Serializable]
    class ExcelHeaderConfig
    {
        // The unique instance for this single instance class
        public static ExcelHeaderConfig Instance;

        // Fields: raw excel headers
        #region unused headers
        /////////////////////////////
        // the dimmed line are unused headers
        //public string EntryDate = "Entry date";
        //public string DeliveryDate = "Delivery date";
        //public string CloseDate = "Close date";
        //public string Reporter = "Reporter";
        //public string DefectivePartMaterialNumber = "Defective part material number";
        //public string DefectivePartSerialNumber = "Defective part serial number";
        //public string Supplier = "Supplier";
        //public string Suppliername = "Suppliername";
        //public string FailureClassification = "PCategory";
        //public string Responsibility = "Responsibility";
        //public string ActionCodeFactor = "Action code factor";
        //public string FixTime = "Fix time (decimal)";
        //public string Creator = "Creator";
        //public string SystemMaterialNumber = "System material number";
        //public string SystemShortDescription = "System short description";
        //public string ComponentMaterialNumber = "Component material number";
        //public string ComponentSerialNumber = "Component serial number";
        //public string Source = "Source";
        //public string FilwCount = "File count";
        //////////////////////////////////
        #endregion
        private string sourceId = "Source ID";
        private string createDate = "Create date";
        private string defectivePartDescription = "Defective part description";
        private string problemDescription = "Problem description";
        private string conclusion = "Conclusion";
        private string systemDescription = "System description";
        private string systemSerialNumber = "System serial number";
        private string componentDescription = "Component description";
        // Fields for table header
        public string HdProperty = "Property";
        public string HdName = "Header Name";
        public string HdId = "Id";
        

        /// <summary>
        /// Properties for Excel header
        /// </summary>
        public string SourceId                      { get => sourceId; set => sourceId = value; }
        public string CreateDate                    { get => createDate; set => createDate = value; }
        public string DefectivePartDescription      { get => defectivePartDescription; set => defectivePartDescription = value; }
        public string ProblemDescription            { get => problemDescription; set => problemDescription = value; }
        public string Conclusion                    { get => conclusion; set => conclusion = value; }
        public string SystemDescription             { get => systemDescription; set => systemDescription = value; }
        public string SystemSerialNumber            { get => systemSerialNumber; set => systemSerialNumber = value; }
        public string ComponentDescription          { get => componentDescription; set => componentDescription = value; }
        



        // protect the constructor to be used externally
        private ExcelHeaderConfig() { }

        public static ExcelHeaderConfig GetInstance()
        {
            if (Instance == null)
                Instance = new ExcelHeaderConfig();
            return Instance;

        }

        public DataTable GetProperties()
        {
            DataTable tbResult = new DataTable("ExcelHeaderConfig");
            tbResult.Columns.Add(HdId, typeof(System.Int32));
            tbResult.Columns.Add(HdProperty, typeof(System.String));
            tbResult.Columns.Add(HdName, typeof(System.String));

            Type type = this.GetType();
            PropertyInfo[] propertyInfos = type.GetProperties();

            DataRow dr = null;
            int id = 1;
            foreach (PropertyInfo f in propertyInfos)
            {
                dr = tbResult.Rows.Add();
                dr[0] = id++;
                dr[1] = f.Name;
                dr[2] = f.GetValue(this) as string;
            }

            return tbResult;
        }

        public void SetProperties(DataTable tb)
        {
            if (tb.Columns.Count != 3) throw new Exception("The input datatable has wrong format!");

            Type type = this.GetType();
            PropertyInfo[] propertyInfos = type.GetProperties();

            //get property name list
            List<string> pNames = new List<string>();
            foreach(PropertyInfo p in propertyInfos)
            {
                pNames.Add(p.Name);
            }

            foreach (DataRow r in tb.Rows)
            {
                if (!pNames.Contains(r[1].ToString())) continue;
                PropertyInfo p = type.GetProperty(r[1].ToString());
                p.SetValue(this, r[2].ToString());
            }
        }
    }
}
