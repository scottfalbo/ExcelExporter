// ------------------------------------
// Excel Export POC
// ------------------------------------

namespace ExcelExporter;
internal class SampleDataBuilder
{
    private Guid _id = Guid.NewGuid();
    private DateTimeOffset _date = DateTimeOffset.Now;
    private decimal _amount;
    private bool _hasSomething;
    private int _quantity;
    private bool _isWhatever;
    private string _description;

    public SampleDataBuilder()
    {
        var random = new Random();

        _amount = (decimal)random.NextDouble();
        _hasSomething = random.Next(0, 2) == 1;
        _quantity = random.Next(1, 100);
        _isWhatever = random.Next(0, 2) == 1;
        _description = "I'm a sample description, yep.";
    }

    public List<SampleData1> BuildSampleData1Collection()
    {
        var dataList = new List<SampleData1>();

        for (var i = 0; i < 20; i++)
        {
            var data = BuildSampleData1();
            dataList.Add(data);
        }

        return dataList;
    }

    public List<SampleData2> BuildSampleData2Collection()
    {
        var dataList = new List<SampleData2>();

        for (var i = 0; i < 20; i++)
        {
            var data = BuildSampleData2();
            dataList.Add(data);
        }

        return dataList;
    }

    private SampleData1 BuildSampleData1()
    {
        return new SampleData1
        {
            Id = _id,
            Date = _date,
            Amount = _amount,
            IsWhatever = _isWhatever,
            Description = _description
        };
    }

    private SampleData2 BuildSampleData2()
    {
        return new SampleData2
        {
            Id = _id,
            Date = _date,
            Amount = _amount,
            HasSomething = _hasSomething,
            Quantity = _quantity
        };
    }
}
