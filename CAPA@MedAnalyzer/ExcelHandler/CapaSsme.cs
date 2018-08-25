using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.Data;
using System.IO;


namespace ExcelHandler
{
    class CapaSsme
    {
        public bool IsInitDone = false;
        public DataTable RawDataTable;
        public DataTableBuilder NcmDataTable;
        public ExcelFileBuilder NcmExcelFile;
        public ExcelHeaderConfig RawExcelHeader = new ExcelHeaderConfig();
        public XmlDictHelper ConfigXml;
        public string ConfigFilePath;
        public List<string> ListofInstinctProductName = new List<string>();
        public List<string> ListofNewFoundName = new List<string>();

        // new added excel headers
        public string ExHdSystemDescriptionNormalized = "System description normalized(string)";
        public string ExHdWeek = "Week";
        public string ExHdMonth = "Month";
        public string ExHdSystemUnionName = "System Union Name";

        // Addtional Table headers
        public string HdQtyWithNcm = "NCM Qty";
        public string HdTotalSysQty = "Sys Qty";
        public string HdFpy = "FPY";
        public string HdId = "Id";

        // Week Spliter
        public string WkBefore = "Before";
        public string WkLast = "Week";
        public string WkLast8 = "Last 8 Weeks";

        #region Init
        /// <summary>
        /// Initialize the class
        /// </summary>
        /// <param name="filePath">The input Excel File</param>
        /// <param name="page">The excel page number to analyze, defualt=1</param>
        /// <param name="cfgFile">The config file, default is "./config.xml"</param>
        public CapaSsme(string filePath, int page = 1, string cfgFile=@"./config.xml")
        {
            if (!File.Exists(filePath)) throw new FileNotFoundException();
            ConfigFilePath = cfgFile;
            NcmDataTable = new DataTableBuilder(filePath, page);
            RawDataTable = NcmDataTable.Table.Copy();
            InitDataTable();
            if (File.Exists(ConfigFilePath))
            {
                ReadConfiXml();
                if (!CheckIfNewNameInExcel()) IsInitDone = true;
            }
            else CreateConfigXml();
            if (MatchProductName()) IsInitDone = true;
        }
        #endregion

        #region Public functions
        /// <summary>
        /// To check the excel file has product name which does not exist in the config file.
        /// </summary>
        /// <returns>if it has unconfigured name, return false and the new found names will be put into ListofNewFoundName; otherwise return true</returns>
        public bool CheckIfNewNameInExcel()
        {
            // Clear ListofNewFoundName
            ListofNewFoundName.Clear();
            // analyze product name
            foreach (string name in ListofInstinctProductName)
            {
                if (!ConfigXml.Names.Keys.Contains(name)) ListofNewFoundName.Add(name);
            }
            if (ListofNewFoundName.Count != 0) return true;
            return false;
        }

        /// <summary>
        /// Add the new column of System Union Name according to config
        /// If the Column has been existing, remove the column first
        /// </summary>
        /// <returns>True: Success; false: Not success. E.g. There is a new name not found in the config</returns>
        public bool MatchProductName()
        {
            if (CheckIfNewNameInExcel()) return false;
            // if the System Union Name column has been there, remove it first
            if (NcmDataTable.Table.Columns.Contains(ExHdSystemUnionName))
                NcmDataTable.Table.Columns.Remove(NcmDataTable.Table.Columns[ExHdSystemUnionName]);
            NcmDataTable.Table.Columns.Add(ExHdSystemUnionName, typeof(System.String));
            // r is a reference, so changing r will change the original data also
            foreach (DataRow r in NcmDataTable.Table.Rows)
                r[ExHdSystemUnionName] = ConfigXml.Names[r[ExHdSystemDescriptionNormalized].ToString()];
            IsInitDone = true;
            return true;
        }

        /// <summary>
        /// ExcelFileBuilder constructor wrapper
        /// </summary>
        /// <param name="filePath">The excel file name to be saved.</param>
        public void InitExcel(string filePath)
        {
            NcmExcelFile = new ExcelFileBuilder(filePath);
        }

        /// <summary>
        /// ExcelFileBuilder.AddDataTable wrapper
        /// </summary>
        /// <param name="table">The DataTable to be converted</param>
        /// <param name="tableName">The Name of the datatable</param>
        public void AddtoExcel(DataTable table, string tableName)
        {
            table.TableName = tableName;
            NcmExcelFile.AddDataTable(table);
        }

        /// <summary>
        /// ExcelFileBuilder.ToExcel wrapper
        /// </summary>
        /// <param name="isForceDelete">If the file already existed, force delete or not. Default=true</param>
        public void SaveExcel(bool isForceDelete = true)
        {
            NcmExcelFile.ToExcel(isForceDelete);
        }

        /// <summary>
        /// Analyze a single month for FPY
        /// </summary>
        /// <param name="year">year as int. E.g. 2018</param>
        /// <param name="month">month as int. E.g. 11</param>
        /// <returns></returns>
        public DataTable AnalyzeFPY4SingleMonth(int year, int month)
        {
            if (!IsInitDone) throw new Exception("Data is not initialized. Maybe found new systerm type.");
            double fMonth = year + month * 0.01;
            // remove unecessary data
            DataTable dt = (from x in NcmDataTable.Table.AsEnumerable()
                            where
                            (x.Field<double>(ExHdMonth) == fMonth &&
                            x.Field<string>(RawExcelHeader.SystemSerialNumber) != "ERROR" &&
                            x.Field<string>(RawExcelHeader.SystemSerialNumber) != "EMPTY" &&
                            x.Field<string>(RawExcelHeader.Conclusion) != "Retest" &&
                            x.Field<string>(ExHdSystemUnionName) != "")
                                                        select x).
                            CopyToDataTable();
            // get name and make instinct
            List<string> names = (from x in dt.AsEnumerable()
                                  select x.Field<string>(ExHdSystemUnionName)).Distinct().ToList();
            // analyze each union name and count distinct serial number
            Dictionary<string, Int32> result = new Dictionary<string, int>();
            foreach (string n in names)
            {
                // calculate distinct count
                int count = (from x in dt.AsEnumerable()
                             where x.Field<string>(ExHdSystemUnionName) == n
                             select x.Field<string>(RawExcelHeader.SystemSerialNumber)).Distinct().Count();
                result.Add(n, count);
            }
            // convert dict to datatable and return
            DataTable dtResult = new DataTable("FPY" + year.ToString() + month.ToString());
            dtResult.Columns.Add(HdId, typeof(System.Int32));
            dtResult.Columns.Add(ExHdSystemUnionName, typeof(System.String));
            dtResult.Columns.Add(HdQtyWithNcm, typeof(System.Int32));
            dtResult.Columns.Add(HdTotalSysQty, typeof(System.Int32));
            dtResult.Columns.Add(HdFpy, typeof(System.Double));
            DataRow dr = null;
            Int32 id = 1;
            foreach (KeyValuePair<string, Int32> kv in result)
            {
                dr = dtResult.Rows.Add();
                dr[0] = id++;
                dr[1] = kv.Key;
                dr[2] = kv.Value;
                dr[3] = kv.Value;
                dr[4] = 1.00;
            }

            return dtResult;
        }

        /// <summary>
        /// Search all NCMs happened in the given product name (Union Name)
        /// </summary>
        /// <param name="sysUnionName">The System Union Name</param>
        /// <returns></returns>
        public DataTable GetProductNcmDetail(string sysUnionName, int year, int month)
        {
            if (!IsInitDone) throw new Exception("Data is not initialized. Maybe found new systerm type.");
            double fMonth = year + 0.01 * month;
            DataTable resultTable = new DataTable("Product NCM detail");
            resultTable.Columns.Add(HdId, typeof(System.Int32));
            resultTable.Columns.Add(ExHdSystemUnionName, typeof(System.String));
            resultTable.Columns.Add(RawExcelHeader.DefectivePartDescription, typeof(System.String));
            resultTable.Columns.Add("Count", typeof(System.String));

            // Don't get the sense for linq!!
            /////////////////////////////////////////////////////////////
            //var dt = from x in NcmDataTable.Table.AsEnumerable()
            //         where x.Field<string>(StrSystemUnionName) == productName
            //         group x by new
            //         {
            //             product1 = x.Field<string>(StrSystemUnionName),
            //             product2 = x.Field<string>(StrSystemDescriptionNormalized),
            //             name = x.Field<string>(DefectivePartDescription),
            //         } into g
            //         select new
            //         {
            //             productname1 = g.Key.product1,
            //             productname2 = g.Key.product2,
            //             materialname = g.Key.name,
            //             namecount = g.Count(x => x.Field<string>(DefectivePartDescription) == productName)
            //         };
            /////////////////////////////////////////////////
            // Give up linq, manual count

            Dictionary<string, int> count = new Dictionary<string, int>();
            // 1st, Generate a datatable only has the "productName"
            DataTable dt = (from x in NcmDataTable.Table.AsEnumerable()
                            where x.Field<string>(ExHdSystemUnionName) == sysUnionName &&
                                  x.Field<double>(ExHdMonth) == fMonth
                            select x).CopyToDataTable();

            // 2nd, Get distinct material name and count in the same time
            foreach (DataRow r in dt.Rows)
            {
                // if not added yet, added into key and set number to 1
                if (!count.Keys.Contains(r[RawExcelHeader.DefectivePartDescription].ToString()))
                    count.Add(r[RawExcelHeader.DefectivePartDescription].ToString(), 1);
                // if already added, increase the number
                else count[r[RawExcelHeader.DefectivePartDescription].ToString()] += 1;
            }

            // 3rd, convert dict to a datatable, remember to reorder the dict according to value fist
            count = count.OrderByDescending(p => p.Value).ToDictionary(p=>p.Key, o=>o.Value);
            DataRow dr = null;
            Int32 id = 1;
            foreach (KeyValuePair<string, int> kv in count)
            {
                dr = resultTable.Rows.Add();
                dr[0] = id++;
                dr[1] = sysUnionName;
                dr[2] = kv.Key.ToString();
                dr[3] = kv.Value.ToString();
            }

            // 4th, return the fuck table
            return resultTable;
        }

        /// <summary>
        /// Search all NCMs for a given maternal
        /// </summary>
        /// <param name="materialName">The material name</param>
        /// <returns></returns>
        public DataTable GetMaterialDetail(string materialName)
        {
            if (!IsInitDone) throw new Exception("Data is not initialized. Maybe found new systerm type.");
            DataTable resultTable = new DataTable("Product NCM detail");
            resultTable.Columns.Add(HdId, typeof(System.Int32));
            resultTable.Columns.Add(ExHdSystemUnionName, typeof(System.String));
            resultTable.Columns.Add(ExHdSystemDescriptionNormalized, typeof(System.String));
            resultTable.Columns.Add(RawExcelHeader.ComponentDescription, typeof(System.String));
            resultTable.Columns.Add(ExHdMonth, typeof(System.String));
            resultTable.Columns.Add(RawExcelHeader.DefectivePartDescription, typeof(System.String));
            resultTable.Columns.Add(RawExcelHeader.ProblemDescription, typeof(System.String));
            resultTable.Columns.Add(RawExcelHeader.Conclusion, typeof(System.String));

            DataTable dt = (from x in NcmDataTable.Table.AsEnumerable()
                            where x.Field<string>(RawExcelHeader.DefectivePartDescription) == materialName
                            orderby x.Field<double>(ExHdMonth) descending
                            select x).CopyToDataTable();
            DataRow dr = null;
            Int32 id = 1;
            foreach (DataRow r in dt.Rows)
            {
                dr = resultTable.Rows.Add();
                dr[0] = id++;
                dr[1] = r[ExHdSystemUnionName].ToString();
                dr[2] = r[ExHdSystemDescriptionNormalized].ToString();
                dr[3] = r[RawExcelHeader.ComponentDescription].ToString();
                dr[4] = r[ExHdMonth].ToString();
                dr[5] = r[RawExcelHeader.DefectivePartDescription].ToString();
                dr[6] = r[RawExcelHeader.ProblemDescription].ToString();
                dr[7] = r[RawExcelHeader.Conclusion].ToString();
            }

            return resultTable;
        }

        /// <summary>
        /// Search All NCMs happened for the WEEK, only show the system name and NCM counts
        /// </summary>
        /// <returns></returns>
        public DataTable GetWeeklyNcmGeneral()
        {
            if (!IsInitDone) throw new Exception("Data is not initialized. Maybe found new systerm type.");
            DataTable tbResult = new DataTable("Weekly Ncm General");
            tbResult.Columns.Add(HdId, typeof(System.Int32));
            tbResult.Columns.Add(ExHdSystemUnionName, typeof(System.String));
            tbResult.Columns.Add("Count", typeof(System.String));
            Dictionary<string, int> count = new Dictionary<string, int>();

            try
            {
                // filter table only has defined week
                DataTable dt = (from x in NcmDataTable.Table.AsEnumerable()
                                where x.Field<string>(ExHdWeek) == WkLast
                                select x).CopyToDataTable();
                // count
                foreach (DataRow r in dt.Rows)
                {
                    // if not added yet, added into key and set number to 1
                    if (!count.Keys.Contains(r[ExHdSystemUnionName].ToString()))
                        count.Add(r[ExHdSystemUnionName].ToString(), 1);
                    // if already added, increase the number
                    else count[r[ExHdSystemUnionName].ToString()] += 1;
                }
            }
            catch (InvalidOperationException)
            {
                // It could be there is no new NCM happened in the defined week
                return null;
            }
            
            // sort
            count = count.OrderByDescending(p => p.Value).ToDictionary(p => p.Key, o => o.Value);
            // convert to datatable
            DataRow dr = null;
            Int32 id = 1;
            foreach (KeyValuePair<string, int> kv in count)
            {
                dr = tbResult.Rows.Add();
                dr[0] = id++;
                dr[1] = kv.Key.ToString();
                dr[2] = kv.Value.ToString();
            }
            return tbResult;
        }

        /// <summary>
        /// Search all ncms and details according to system name
        /// </summary>
        /// <param name="sysUnionName"></param>
        /// <returns></returns>
        public DataTable GetNcmWeeklyDetail(string sysUnionName)
        {
            if (!IsInitDone) throw new Exception("Data is not initialized. Maybe found new systerm type.");
            DataTable tbResult = new DataTable("Weekly Ncm Detail");
            tbResult.Columns.Add(HdId, typeof(System.Int32));
            tbResult.Columns.Add(ExHdSystemUnionName, typeof(System.String));
            tbResult.Columns.Add(ExHdSystemDescriptionNormalized, typeof(System.String));
            tbResult.Columns.Add(RawExcelHeader.ComponentDescription, typeof(System.String));
            tbResult.Columns.Add(ExHdMonth, typeof(System.String));
            tbResult.Columns.Add(RawExcelHeader.DefectivePartDescription, typeof(System.String));
            tbResult.Columns.Add(RawExcelHeader.ProblemDescription, typeof(System.String));
            tbResult.Columns.Add(RawExcelHeader.Conclusion, typeof(System.String));
            tbResult.Columns.Add(ExHdWeek, typeof(System.String));

            DataRow dr = null;
            Int32 id = 1;

            try
            {
                // filter table only has defined week
                DataTable dt = (from x in NcmDataTable.Table.AsEnumerable()
                                where x.Field<string>(ExHdWeek) == WkLast &&
                                x.Field<string>(ExHdSystemUnionName) == sysUnionName
                                orderby x.Field<string>(RawExcelHeader.DefectivePartDescription)
                                select x).CopyToDataTable();
                // copy the data to the result table
                foreach (DataRow r in dt.Rows)
                {
                    dr = tbResult.Rows.Add();
                    dr[0] = id++;
                    dr[1] = r[ExHdSystemUnionName].ToString();
                    dr[2] = r[ExHdSystemDescriptionNormalized].ToString();
                    dr[3] = r[RawExcelHeader.ComponentDescription].ToString();
                    dr[4] = r[ExHdMonth].ToString();
                    dr[5] = r[RawExcelHeader.DefectivePartDescription].ToString();
                    dr[6] = r[RawExcelHeader.ProblemDescription].ToString();
                    dr[7] = r[RawExcelHeader.Conclusion].ToString();
                    dr[8] = r[ExHdWeek].ToString();
                }
            }
            catch (InvalidOperationException)
            {
                // It could be there is no new NCM happened in the defined week
                return null;
            }
            return tbResult;
        }

        /// <summary>
        /// To split a given date to specify WEEK and rewrite in the Table
        /// </summary>
        /// <param name="splitTime">a DataTime, any dates after the time will be recongnized as WEEK, else Last 8 weeks or before</param>
        /// <returns>3 options: "Week", "Last 8 Weeks" and "Before"</returns>
        public void RecaclulateWeekSpliter(DateTime splitTime)
        {
            if (!IsInitDone) throw new Exception("Data is not initialized. Maybe found new systerm type.");
            DateTime tempDateTime;
            foreach (DataRow r in NcmDataTable.Table.Rows)
            {
                // change colum for split week
                tempDateTime = DateTime.ParseExact(r[RawExcelHeader.CreateDate].ToString(), "yyyy-MM-dd HH:mm",
                                       System.Globalization.CultureInfo.InvariantCulture);
                r[ExHdWeek] = JudgeWeekSpliter(tempDateTime, splitTime);
            }
        }



        public List<string> GetExcelHeaderConfigProperties()
        {
            List<string> lstResult = new List<string>();
            foreach(var p in RawExcelHeader.GetType().GetProperties())
            {
                lstResult.Add(p.Name);
            }
            return lstResult;
        }
        #endregion

        #region private functions
        private bool InitDataTable()
        {
            // Add a new col with normalized product name
            NcmDataTable.Table.Columns.Add(ExHdSystemDescriptionNormalized, typeof(System.String));
            NcmDataTable.Table.Columns.Add(ExHdWeek, typeof(System.String));
            NcmDataTable.Table.Columns.Add(ExHdMonth, typeof(System.Double));
            string tempProductNameNormal;
            DateTime tempDateTime;
            // the local r is a reference, so changing r will change NCMDataTable.Table.Rows
            foreach (DataRow r in NcmDataTable.Table.Rows)
            {
                // make the new colum for normalized name
                tempProductNameNormal = r[RawExcelHeader.SystemDescription].ToString().ToUpper();
                tempProductNameNormal = NormalizeString(tempProductNameNormal);
                r[ExHdSystemDescriptionNormalized] = tempProductNameNormal;
                // make the new colum for split week
                tempDateTime = DateTime.ParseExact(r[RawExcelHeader.CreateDate].ToString(), "yyyy-MM-dd HH:mm",
                                       System.Globalization.CultureInfo.InvariantCulture);
                r[ExHdWeek] = JudgeWeekSpliter(tempDateTime, DateTime.Now);
                // make the new coum for month format with yyyy.mm which is a float number
                r[ExHdMonth] = tempDateTime.Year + tempDateTime.Month * 0.01;
            }
            // analyze the normlized product name and remove dedunt
            foreach (DataRow r in NcmDataTable.Table.Rows)
            {
                if (ListofInstinctProductName.Contains(r[ExHdSystemDescriptionNormalized])) continue;
                ListofInstinctProductName.Add(r[ExHdSystemDescriptionNormalized].ToString());
            }
            return true;
        }

        private string JudgeWeekSpliter(DateTime dtInData, DateTime dtCompareTo)
        {
            if (dtInData >= dtCompareTo) return WkLast;
            else if ((dtCompareTo - dtInData).Days <= 56) return WkLast8;
            else return WkBefore;
        }

        private void CreateConfigXml()
        {
            ConfigXml = new XmlDictHelper();
            foreach (DataRow r in NcmDataTable.Table.Rows)
            {
                ConfigXml.AddData(r[ExHdSystemDescriptionNormalized].ToString(), r[ExHdSystemDescriptionNormalized].ToString());
            }
            ConfigXml.SaveXml(ConfigFilePath);
        }

        private void ReadConfiXml()
        {
            ConfigXml = new XmlDictHelper(ConfigFilePath);
        }

        private static string NormalizeString(string str)
        {
            return str.Replace(" ", "").
                // Be carefule \u00a0 is a "NO_BRAKE_SPACE".
                Replace("\u00a0", "").
                Replace("\u0028", "").
                // Chinese SPACE
                Replace("\u3000", "").
                Replace("\u0029", "");
        }
        #endregion
    }// end of class
}// end of namespace
