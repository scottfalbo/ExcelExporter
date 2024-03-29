// ------------------------------------
// Excel Export POC
// ------------------------------------

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace ExcelExporter;
internal class WorkbookFactory
{
    public static byte[] CreateWorkbook(List<SampleData1> sampleData1, List<SampleData2> sampleData2)
    {
        using var memoryStream = new MemoryStream();
        using (var spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
        {
            var workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            var sheets = workbookpart.Workbook.AppendChild(new Sheets());

            AddSheetToWorkbook(sampleData1, workbookpart, sheets, "sample1");
            AddSheetToWorkbook(sampleData2, workbookpart, sheets, "sample2");
        }

        return memoryStream.ToArray();
    }

    public static void AddSheetToWorkbook<T>(IEnumerable<T> data, WorkbookPart workbookPart, Sheets sheets, string sheetName)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        var sheet = new Sheet()
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = (uint)sheets.Elements<Sheet>().Count() + 1,
            Name = sheetName
        };
        sheets.Append(sheet);

        AddRowsToSheet(data, sheetData);

        CreateTableDefinition(worksheetPart, sheetName + "Table", sheet.SheetId);
    }

    private static void AddRowsToSheet<T>(IEnumerable<T> data, SheetData sheetData)
    {
        var headerRow = new Row();

        foreach (var property in typeof(T).GetProperties())
        {
            headerRow.Append(new Cell { CellValue = new CellValue(property.Name), DataType = CellValues.String });
        }

        sheetData.AppendChild(headerRow);

        foreach (var item in data)
        {
            var row = new Row();

            foreach (var property in typeof(T).GetProperties())
            {
                var value = property.GetValue(item);
                var cell = new Cell();

                if (value is DateTimeOffset dto)
                {
                    var dateTime = dto.UtcDateTime;
                    var oaDateValue = dateTime.ToOADate();

                    cell.CellValue = new CellValue(oaDateValue.ToString());
                    cell.DataType = CellValues.Number;
                }
                else if (value is decimal dec)
                {
                    cell.CellValue = new CellValue(dec.ToString());
                    cell.DataType = CellValues.Number;
                }
                else if (value != null)
                {
                    cell.CellValue = new CellValue(value.ToString());
                    cell.DataType = CellValues.String;
                }

                row.Append(cell);
            }

            sheetData.AppendChild(row);
        }
    }

    private static void CreateTableDefinition(WorksheetPart worksheetPart, string tableName, uint sheetId)
    {
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        var rowCount = sheetData.Elements<Row>().Count();
        var columnCount = sheetData.Elements<Row>().First().Elements<Cell>().Count();

        string reference = $"A1:{GetColumnName(columnCount)}{rowCount}";

        var table = new Table
        {
            Id = GetUniqueTableId(sheetId), // This needs to be unique across the entire workbook
            Name = tableName,
            DisplayName = tableName,
            Reference = reference,
            TotalsRowShown = false,
            TableStyleInfo = new TableStyleInfo
            {
                Name = "TableStyleMedium2",
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false
            }
        };

        var tableColumns = new TableColumns { Count = (uint)columnCount };
        for (int i = 0; i < columnCount; i++)
        {
            var columnName = $"Column{i + 1}";
            tableColumns.Append(new TableColumn { Id = (uint)(i + 1), Name = columnName });
        }
        table.Append(tableColumns);

        TableDefinitionPart tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
        tablePart.Table = table;

        // After you create the TableDefinitionPart, you need to add a relationship.
        // This is a critical step that is often missed.
        var rId = worksheetPart.GetIdOfPart(tablePart);
        tablePart.Table.Save();

        // Link the TablePart to the Worksheet through TableParts
        TableParts tableParts = worksheetPart.Worksheet.Elements<TableParts>().FirstOrDefault();
        if (tableParts == null)
        {
            tableParts = new TableParts { Count = 1 };
            worksheetPart.Worksheet.Append(tableParts);
        }
        else
        {
            tableParts.Count.Value++;
        }

        tableParts.Append(new TablePart { Id = rId });
        worksheetPart.Worksheet.Save();
    }

    private static uint GetUniqueTableId(uint sheetId)
    {
        // Implement logic to generate a unique ID for each table.
        // This could be a static counter that you increment for each new table.
        // For this example, let's just return 1, but you need to make sure this is actually unique.
        return sheetId;
    }

    // Helper method to convert a column index to its corresponding Excel column name
    private static string GetColumnName(int columnIndex)
    {
        int dividend = columnIndex;
        string columnName = String.Empty;
        int modulo;

        while (dividend > 0)
        {
            modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            dividend = (int)((dividend - modulo) / 26);
        }

        return columnName;
    }
}
