# Asereware.SpreadsheetLight
A new version of SpreadsheetLight from  original Copyright (c) 2011 Vincent Tan Wai Lip version, but with support for DocumentFormat.OpenXml.

https://spreadsheetlight.com

# Dependencies:
- DocumentFormat.OpenXml v2.10.1 to support save Excel macro files. Before we have v2.19.0.0 of OpenXml, but does not support save macro files.
- Net Framework v4.8.1

Please got to https://spreadsheetlight.com/developers/ for more information.

# Constraints
- Office SmartTag support has been removed (but still commented in source code).

# To install from Nuget go to:

https://www.nuget.org/packages/asereware.spreadsheetlight/

# Use it
Create xlsx file and get data table and data view.
```csharp
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
```
