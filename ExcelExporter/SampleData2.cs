namespace ExcelExporter;

internal class SampleData2
{
    public Guid Id { get; set; } = Guid.NewGuid();

    public DateTimeOffset Date { get; set; } = DateTimeOffset.Now;

    public decimal Amount { get; set; }

    public bool HasSomething { get; set; }

    public int Quantity { get; set; }
}
