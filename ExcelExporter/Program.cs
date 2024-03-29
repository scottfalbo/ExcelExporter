// ------------------------------------
// Excel Export POC
// ------------------------------------

using ExcelExporter;

Console.WriteLine("-- Excel Exporter --");

Export();

static void Export()
{
    Console.WriteLine("Building workbook...");

    var workbookName = "SampleWorkbook.xlsx";

    var sampleDataBuilder = new SampleDataBuilder();
    var sampleData1 = sampleDataBuilder.BuildSampleData1Collection();
    var sampleData2 = sampleDataBuilder.BuildSampleData2Collection();

    var workbook = WorkbookFactory.CreateWorkbook(sampleData1, sampleData2);

    if (!workbookName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
    {
        workbookName += ".xlsx";
    }

    File.WriteAllBytes(workbookName, workbook);

    Console.WriteLine($"Workbook saved to {Path.GetFullPath(workbookName)}");
}
