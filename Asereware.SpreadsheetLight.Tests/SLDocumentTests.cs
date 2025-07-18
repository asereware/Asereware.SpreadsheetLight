using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetLight.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace SpreadsheetLight.Tests
{
    [TestClass()]
    public class SLDocumentTests
    {
        private const string _coFileName = "Hello World.xlsx";

        [TestMethod]
        public void SetCellValueOnMacroTest()
        {
            var path = Path.Combine(Environment.CurrentDirectory, "Files");
            var filePath = Path.Combine(path, "FileMacro1.xlsm");
            if (File.Exists(filePath))
            {
                using (var sd = new SLDocument(filePath, "Hoja1"))
                {
                    Debug.Print("Loaded.");

                    sd.SetCellValue("A2", $"New Value {DateTime.Now.ToString("s")}");
                    sd.SaveAs($"{path}\\FileMacroResult.xlsm");

                }
            }

            Assert.IsTrue(true);
        }

        [TestMethod]
        public void SetCellValueOnMacroOnProtectedWorkbookTest()
        {
            var path = Path.Combine(Environment.CurrentDirectory, "Files");
            var filePath = Path.Combine(path, "1695985_RCC - 2024 v3.3 - Apr 2024 Ergo.xlsm");
            var copyFilePath = $"{path}\\Updated-1695985_RCC - 2024 v3.3 - Apr 2024 Ergo.xlsm";
            if (File.Exists(filePath))
            {
                using (var sd = new SLDocument(filePath, "Hoja2"))
                {
                    Debug.Print("Loaded.");
                    sd.HideWorksheet("Hoja2", IsVeryHidden: true);
                    sd.SetCellValue("A2", $"New Value {DateTime.Now.ToString("s")}");
                    sd.SaveAs(copyFilePath);
                }

                using (var sd = new SLDocument(copyFilePath))
                {
                    if (!sd.IsWorksheetHidden("Hoja2"))
                    {
                        sd.HideWorksheet("Hoja2", IsVeryHidden: true);
                        sd.SetCellValue("A2", $"New Value {DateTime.Now.ToString("s")}");
                        sd.Save();
                    }
                }
            }
        }

        [TestMethod]
        public void LocalFilexlsmTest()
        {
            var fileName = "1609721_RCC - 2024 v3.2 - Feb 2024 Ergo (PROD).xlsm";
            var indexPage = "Index";
            var regulatoryPage = "Regulatory Instruments";
            var path = Path.Combine(Environment.CurrentDirectory, "Files");
            var filePath = Path.Combine(path, fileName);
            var copyFilePath = $"{path}\\Updated-{fileName}";

            List<string> files = new List<string>() { "uno", "dos" };

            if (File.Exists(filePath))
            {
                using (var doc = new SLDocument(filePath))
                {

                    if (doc.SelectWorksheet(regulatoryPage))
                    {
                        CleanData(doc);
                        int row = 5;
                        foreach (var instr in files)
                        {
                            doc.SetCellValue(row, 2, instr);
                            doc.SetCellValue(row, 3, instr);
                            doc.SetCellValue(row, 4, instr);
                            doc.SetCellValue(row, 5, instr);
                            row++;
                        }

                        doc.SelectWorksheet(indexPage);
                        doc.HideWorksheet(regulatoryPage);
                        Debug.Print($"Saving file {filePath}...");
                        doc.SaveAs(filePath);
                        Debug.Print("Saved OK.");
                    }
                }
            }
        }

        private void CleanData(SLDocument sl)
        {
            for (int i = 5; i <= 104; ++i)
            {
                for (int j = 2; j <= 5; ++j)
                {
                    sl.SetCellValue(i, j, String.Empty);
                }
            }
        }

        [TestMethod]
        public void SetCellValueOnMacroOnProtectedWorkbookTest2()
        {
            var path = Path.Combine(Environment.CurrentDirectory, "Files");
            var filePath = Path.Combine(path, "RCC-Protected.xlsm");
            var copyFilePath = $"{path}\\RCC-Copy.xlsm";
            if (File.Exists(filePath))
            {
                using (var sd = new SLDocument(filePath, "Regulatory Instruments"))
                {
                    Debug.Print("Loaded.");
                    sd.HideWorksheet("Hoja2", IsVeryHidden: true);                    
                    sd.SetCellValue("B5", $"New Value {DateTime.Now.ToString("s")}");
                    sd.SaveAs(copyFilePath);
                }
            }

            Assert.IsTrue(true);
        }


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