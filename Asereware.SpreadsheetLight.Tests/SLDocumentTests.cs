using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetLight.Tests
{
    [TestClass()]
    public class SLDocumentTests
    {
        [TestMethod()]
        public void SetCellValueTest()
        {

            using (var sd = new SLDocument())
            {
                sd.SetCellValue("A1", "Mi valor de texto");
                sd.SetCellValue("B1", true);
                sd.SetCellValue(1, 3, DateTime.Now);
                sd.SaveAs("Hello World.xlsx");
            }
        }
    }
}