// ------------------------------------
// Excel Export POC
// ------------------------------------

using ClosedXML.Excel;

namespace ExcelExporter;
internal class WorkbookFactory
{
    public static void CreateWorkbook(List<SampleData1> sampleData1, List<SampleData2> sampleData2)
    {
        using var workbook = new XLWorkbook();
        var ws1 = workbook.Worksheets.Add("Sample Data 1");
        var ws2 = workbook.Worksheets.Add("Sample Data 2");

        // Insert the data as a table
        var table1 = ws1.Cell(1, 1).InsertTable(sampleData1.AsEnumerable(), "SampleData1Table", true);
        var table2 = ws2.Cell(1, 1).InsertTable(sampleData2.AsEnumerable(), "SampleData2Table", true);

        // Optional: Adjust column widths to content
        ws1.Columns().AdjustToContents();
        ws2.Columns().AdjustToContents();

        workbook.SaveAs("SampleWorkbook2.xlsx");
    }
}
