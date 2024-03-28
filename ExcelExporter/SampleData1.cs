// ------------------------------------
// Excel Export POC
// ------------------------------------

namespace ExcelExporter;

internal class SampleData1
{
    public Guid Id { get; set; } = Guid.NewGuid();

    public DateTimeOffset Date { get; set; } = DateTimeOffset.Now;

    public decimal Amount { get; set; }

    public bool IsWhatever { get; set; }

    public string Description { get; set; }
}
