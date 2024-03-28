// ------------------------------------
// Excel Export POC
// ------------------------------------

using ExcelExporter;

Console.WriteLine("Excel Exporter");

Export();

static void Export()
{
    var sampleData1 = GetSampleData1();
    var sampleData2 = GetSampleData2();

    
}



static List<SampleData1> GetSampleData1()
{
    var dataList = new List<SampleData1>();

    for (var i = 0; i < 20; i++)
    {
        var data = new SampleDataBuilder().BuildSampleData1();
        dataList.Add(data);
    }

    return dataList;
}

static List<SampleData2> GetSampleData2()
{
    var dataList = new List<SampleData2>();

    for (var i = 0; i < 20; i++)
    {
        var data = new SampleDataBuilder().BuildSampleData2();
        dataList.Add(data);
    }

    return dataList;
}