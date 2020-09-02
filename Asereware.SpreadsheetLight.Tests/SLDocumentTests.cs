﻿using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetLight;
using SpreadsheetLight.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetLight.Tests
{
    [TestClass()]
    public class SLDocumentTests
    {
        private const string _coFileName = "Hello World.xlsx";

        [TestMethod()]
        public void SetCellValueTest()
        {
            using (var sd = new SLDocument())
            {
                sd.SetCellValue("A1", "My text value");
                sd.SetCellValue("B1", true);
                sd.SetCellValue(1, 3, DateTime.Now.ToString("g"));
                //
                sd.SetCellValue("A2", "My second value");
                sd.SetCellValue("B2", false);
                sd.SetCellValue(2, 3, DateTime.Now.ToString("g"));
                //
                sd.SetCellValue("A3", "The last");
                sd.SetCellValue("B3", true);
                sd.SetCellValue(3, 3, DateTime.Now.ToString("g"));

                sd.SaveAs(_coFileName);
            }

            Assert.IsTrue(File.Exists(_coFileName));
            CreateDataTableExtension(_coFileName);
            CreateDataViewExtension(_coFileName);
        }

        [TestMethod()]
        public void SetCellValueHeaderTest()
        {
            var fileName = $"Headers - {_coFileName}";
            using (var sd = new SLDocument())
            {
                //Headers.
                sd.SetCellValue("A1", "One");
                sd.SetCellValue("B1", "Two");
                sd.SetCellValue(1, 3, "Three");
                //Values.
                sd.SetCellValue("A2", "My text value");
                sd.SetCellValue("B2", true);
                sd.SetCellValue(2, 3, DateTime.Now.ToString("g"));
                //
                sd.SetCellValue("A3", "My second value");
                sd.SetCellValue("B3", false);
                sd.SetCellValue(3, 3, DateTime.Now.ToString("g"));
                //
                sd.SetCellValue("A4", "The last");
                sd.SetCellValue("B4", true);
                sd.SetCellValue(4, 3, DateTime.Now.ToString("g"));

                sd.SaveAs(fileName);
            }

            Assert.IsTrue(File.Exists(fileName));
            CreateDataTableExtension(fileName, hasHeaders: true);
            CreateDataViewExtension(fileName, hasHeaders: true);
        }


        private void CreateDataTableExtension(string fileName, bool hasHeaders = false)
        {
            System.Data.DataTable dt = null;
            using (var sd = new SLDocument(fileName))
            {
                dt = sd.CreateDataTable(hasHeaders);                
            }

            Assert.IsTrue(dt.Columns.Count == 3);
            Assert.IsTrue(dt.Rows.Count == 3);
        }

        private void CreateDataViewExtension(string fileName, bool hasHeaders = false)
        {
            System.Data.DataView dv = null;
            using (var sd = new SLDocument(fileName))
            {
                dv = sd.CreateDataView(hasHeaders);
            }

            Assert.IsTrue(dv.Count == 3);
            Assert.IsTrue(dv.Table.Columns.Count == 3);
            Assert.IsTrue(dv.Table.Rows.Count == 3);

            var col1Name = dv.Table.Columns[0].ColumnName;
            dv.RowFilter = $"{col1Name} LIKE '%text%'";
            Assert.IsTrue(dv.Count == 1);
            dv.RowFilter = $"{col1Name} LIKE '%none%'";
            Assert.IsTrue(dv.Count == 0);
            dv.RowFilter = String.Empty;
            Assert.IsTrue(dv.Count == 3);
        }
    }
}