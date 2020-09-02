using DocumentFormat.OpenXml.Drawing.Charts;
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
                sd.SetCellValue(1, 3, DateTime.Now.ToString("s"));
                sd.SaveAs(_coFileName);
            }

            Assert.IsTrue(File.Exists(_coFileName));
            CreateDataTableExtension();
            CreateDataViewExtension();
        }

        private void CreateDataTableExtension()
        {
            System.Data.DataTable dt = null;
            using (var sd = new SLDocument(_coFileName))
            {
                dt = sd.CreateDataTable(hasHeaders: false);                
            }

            Assert.IsTrue(dt.Columns.Count == 3);
            Assert.IsTrue(dt.Rows.Count == 1);
        }

        private void CreateDataViewExtension()
        {
            System.Data.DataView dv = null;
            using (var sd = new SLDocument(_coFileName))
            {
                dv = sd.CreateDataView(hasHeaders: false);
            }

            Assert.IsTrue(dv.Count == 1);
            Assert.IsTrue(dv.Table.Columns.Count == 3);
            Assert.IsTrue(dv.Table.Rows.Count == 1);

            var col1Name = dv.Table.Columns[0].ColumnName;
            dv.RowFilter = $"{col1Name} LIKE '%text%'";
            Assert.IsTrue(dv.Count == 1);
            dv.RowFilter = $"{col1Name} LIKE '%none%'";
            Assert.IsTrue(dv.Count == 0);
            dv.RowFilter = String.Empty;
            Assert.IsTrue(dv.Count == 1);
        }
    }
}