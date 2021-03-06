﻿using System;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using System.Data;

namespace ExcelHandler
{
    public class XmlDictHelper
    {
        public Dictionary<string, string> Names = new Dictionary<string, string>();
        private string key = "Product Name";
        private string value = "Union Name";
        private string Hdid = "Id";
        private string FirstLevelName = "root";
        private string SecondLevelName = "item";
        private string ThirdLevelName1 = "key";
        private string ThirdLevelName2 = "value";

        #region Init
        /// <summary>
        /// Read a file to import the xml
        /// </summary>
        /// <param name="filePath">File path.</param>
        public XmlDictHelper(string filePath)
        {
            
            if (!File.Exists(filePath)) throw new FileNotFoundException();
            ReadFileToFillDictionary(filePath);   
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="T:ExcelTest.XmlHelper"/> class.
        /// </summary>
        public XmlDictHelper() { }
        #endregion

        #region Public Functions
        /// <summary>
        /// Adds the data.
        /// </summary>
        /// <returns><c>true</c>, if data was added, <c>false</c> if data is updated.</returns>
        /// <param name="key">Key.</param>
        /// <param name="value">Value.</param>
        /// <param name="isForceUpdate">If set to <c>true</c> is force update when key is duplicated.</param>
        public bool AddData(string key, string value, bool isForceUpdate = true)
        {
            if (IsKeyInDictionary(key))
            {
                if (!isForceUpdate) throw new Exception("Key alrady exists! Maybe use isForceUpdate=true to force update data.");
                else
                {
                    Names[key] = value;
                    return false;
                }
            }
            Names.Add(key, value);
            return true;
        }

        /// <summary>
        /// Saves the xml.
        /// </summary>
        /// <returns><c>true</c>, if xml was saved, <c>false</c> otherwise.</returns>
        /// <param name="filePath">File path.</param>
        /// <param name="isForceDelete">If set to <c>true</c> is force delete.</param>
        public bool SaveXml(string filePath, bool isForceDelete = true)
        {
            // check and deal if file already exists
            if (File.Exists(filePath))
            {
                if (!isForceDelete) throw new Exception("File is already existed. To force save use isForceDelete=true");
                else File.Delete(filePath);
            }

            // Initial xml
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode header = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
            xmlDoc.AppendChild(header);
            XmlElement root = xmlDoc.CreateElement(FirstLevelName);
            // export Names to xml and save file
            foreach (KeyValuePair<string, string> kv in Names)
            {
                XmlElement xn = InsertItem(kv, xmlDoc);
                root.AppendChild(xn);
            }
            xmlDoc.AppendChild(root);
            xmlDoc.Save(filePath);
            return true;
        }

        /// <summary>
        /// Reads the file to fill dictionary.
        /// </summary>
        /// <returns><c>true</c>, if file to fill dictionary was  read, <c>false</c> otherwise.</returns>
        /// <param name="filePath">File path.</param>
        public bool ReadFileToFillDictionary(string filePath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load(filePath);
                XmlNodeList itemNodeList = xmlDoc.SelectNodes('/' + FirstLevelName + '/' + SecondLevelName);
                if (itemNodeList == null) return false;
                foreach (XmlNode nl in itemNodeList)
                {
                    string key = nl.FirstChild.InnerXml;
                    string value = nl.LastChild.InnerXml;
                    AddData(key, value);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return true;
        }

        /// <summary>
        /// Return Dictionary as a table
        /// </summary>
        /// <returns></returns>
        public DataTable DictToDataTable()
        {
            DataTable NameTable = new DataTable("Config");
            NameTable.Columns.Add(Hdid, typeof(System.Int32));
            NameTable.Columns.Add(key, typeof(System.String));
            NameTable.Columns.Add(value, typeof(System.String));
            DataRow row = null;
            Int32 id = 1;
            foreach (KeyValuePair<string, string> kv in Names)
            {
                row = NameTable.Rows.Add();
                row[Hdid] = id++;
                row[key] = kv.Key;
                row[value] = kv.Value;
            }
            return NameTable;
        }

        /// <summary>
        /// Use the input datatable to override dict
        /// </summary>
        /// <param name="dt"></param>
        public void UpdateDictByDataTable(DataTable dt)
        {
            Names.Clear();
            foreach(DataRow row in dt.Rows)
            {
                if (row[2].ToString().Trim() == "") continue;
                AddData(row[1].ToString(), row[2].ToString());   
            }
        }
        #endregion

        #region Private functions
        private XmlElement InsertItem(KeyValuePair<string, string> kv, XmlDocument xmlDoc)
        {
            XmlElement xn = xmlDoc.CreateElement(SecondLevelName);
            xn.AppendChild(GetXmlNode(xmlDoc, ThirdLevelName1, kv.Key));
            xn.AppendChild(GetXmlNode(xmlDoc, ThirdLevelName2, kv.Value));
            return xn;
        }

        private XmlNode GetXmlNode(XmlDocument xmlDoc, string name, string value)
        {
            XmlElement xn = xmlDoc.CreateElement(name);
            xn.InnerText = value;
            return xn;
        }

        private bool IsKeyInDictionary(string key)
        {
            if (Names.ContainsKey(key)) return true;
            else return false;
        }
        #endregion


    }// end of class
}
