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
        // Current solution is a fixed solutin and ugly
        // maybe use reflection will be more elegant to be used.
        // TODO()

        // raw excel headers
        // the dimmed line are unused headers
        //public string EntryDate = "Entry date";
        //public string DeliveryDate = "Delivery date";
        //public string CloseDate = "Close date";
        public string SourceId = "Source ID";
        public string CreateDate = "Create date";
        //public string Reporter = "Reporter";
        //public string DefectivePartMaterialNumber = "Defective part material number";
        public string DefectivePartDescription = "Defective part description";
        //public string DefectivePartSerialNumber = "Defective part serial number";
        //public string Supplier = "Supplier";
        //public string Suppliername = "Suppliername";
        public string ProblemDescription = "Problem description";
        //public string FailureClassification = "PCategory";
        //public string Responsibility = "Responsibility";
        public string Conclusion = "Conclusion";
        //public string ActionCodeFactor = "Action code factor";
        //public string FixTime = "Fix time (decimal)";
        //public string Creator = "Creator";
        //public string SystemMaterialNumber = "System material number";
        public string SystemDescription = "System description";
        //public string SystemShortDescription = "System short description";
        public string SystemSerialNumber = "System serial number";
        //public string ComponentMaterialNumber = "Component material number";
        public string ComponentDescription = "Component description";
        //public string ComponentSerialNumber = "Component serial number";
        //public string Source = "Source";
        //public string FilwCount = "File count";

        // protect the constructor to be used externally
        private ExcelHeaderConfig() { }

        public static ExcelHeaderConfig GetInstance()
        {
            if (Instance == null)
                Instance = new ExcelHeaderConfig();
            return Instance;

        }

        public DataTable GetFields()
        {
            DataTable tbResult = new DataTable("ExcelHeaderConfig");
            tbResult.Columns.Add("Field", typeof(System.String));
            tbResult.Columns.Add("Name", typeof(System.String));

            Type type = this.GetType();
            FieldInfo[] fieldInfo = type.GetFields();

            DataRow dr = null;
            foreach (FieldInfo f in fieldInfo)
            {
                if (f.Name == "Instance") continue;
                dr = tbResult.Rows.Add();
                dr[0] = f.Name;
                dr[1] = f.GetValue(this) as string;
            }

            return tbResult;
        }

        public void SelfSerialized(string filePath="header.hdconfig")
        {
            Stream steam = File.Open(filePath, FileMode.Create);
            BinaryFormatter bf = new BinaryFormatter();
            bf.Serialize(steam, this);
            steam.Close();
        }

        public void SelfDeSerialized(string filePath="Header.config")
        {
            Stream steam = File.Open(filePath, FileMode.Open);
            BinaryFormatter bf2 = new BinaryFormatter();
            Instance = (ExcelHeaderConfig)bf2.Deserialize(steam);
            steam.Close();
        }
    }
}
