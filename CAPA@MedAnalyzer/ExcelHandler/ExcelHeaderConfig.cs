using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHandler
{
    class ExcelHeaderConfig
    {
        // Current solution is a fixed solutin and ugly
        // maybe use reflection will be more elegant to be used.
        // TODO()
        // raw excel headers
        public string EntryDate = "Entry date";
        public string DeliveryDate = "Delivery date";
        public string CloseDate = "Close date";
        public string SourceId = "Source ID";
        public string CreateDate = "Create date";
        public string Reporter = "Reporter";
        public string DefectivePartMaterialNumber = "Defective part material number";
        public string DefectivePartDescription = "Defective part description";
        public string DefectivePartSerialNumber = "Defective part serial number";
        public string Supplier = "Supplier";
        public string Suppliername = "Suppliername";
        public string ProblemDescription = "Problem description";
        public string FailureClassification = "PCategory";
        public string Responsibility = "Responsibility";
        public string Conclusion = "Conclusion";
        public string ActionCodeFactor = "Action code factor";
        public string FixTime = "Fix time (decimal)";
        public string Creator = "Creator";
        public string SystemMaterialNumber = "System material number";
        public string SystemDescription = "System description";
        public string SystemShortDescription = "System short description";
        public string SystemSerialNumber = "System serial number";
        public string ComponentMaterialNumber = "Component material number";
        public string ComponentDescription = "Component description";
        public string ComponentSerialNumber = "Component serial number";
        public string Source = "Source";
        public string FilwCount = "File count";
    }
}
