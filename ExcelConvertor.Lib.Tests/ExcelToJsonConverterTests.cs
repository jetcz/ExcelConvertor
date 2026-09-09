using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace ExcelConvertor.Lib.Tests;

public sealed class ExcelToJsonConverterTests
{
    private readonly ExcelToJsonConverter _sut = new();

    [Fact]
    public async Task ConvertAsync_ConvertsClientsAndProjects()
    {
        await using var workbook = CreateWorkbook(
            ["Klient", "IČ", "Zakázka", "2025-01", "2025-02"],
            [
                ["Acme", "123", "Project A", "4", ""],
                ["Acme", "123", "Project B", "2", long.MaxValue.ToString()],
                ["Beta", "456", "Project C", "", "7"]
            ]);
        await using var output = new MemoryStream();

        await _sut.ConvertAsync(workbook, output);

        using var document = JsonDocument.Parse(output.ToArray());
        var clients = document.RootElement;
        Assert.Equal(2, clients.GetArrayLength());
        Assert.Equal("Acme", clients[0].GetProperty("nazevKlienta").GetString());
        Assert.Equal("123", clients[0].GetProperty("icoKlienta").GetString());
        Assert.Equal(2, clients[0].GetProperty("zakazky").GetArrayLength());
        Assert.Equal(1, clients[0].GetProperty("zakazky")[0].GetProperty("vyrobeneKusy").GetArrayLength());
        Assert.Equal(4, clients[0].GetProperty("zakazky")[0].GetProperty("vyrobeneKusy")[0].GetProperty("pocetKusu").GetInt32());
        Assert.Equal(1, clients[1].GetProperty("zakazky")[0].GetProperty("vyrobeneKusy").GetArrayLength());
    }

    [Fact]
    public async Task ConvertAsync_ThrowsWhenPeriodIsMissing()
    {
        await using var workbook = CreateWorkbook(
            ["Klient", "IČ", "Zakázka", "2025-01", ""],
            [["Acme", "123", "Project A", "1", "2"]]);

        var exception = await Assert.ThrowsAsync<InvalidDataException>(() =>
            _sut.ConvertAsync(workbook, new MemoryStream()));

        Assert.Contains("period header", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertAsync_ThrowsWhenRequiredValueIsMissing()
    {
        await using var workbook = CreateWorkbook(
            ["Klient", "IČ", "Zakázka", "2025-01"],
            [["Acme", "", "Project A", "1"]]);

        var exception = await Assert.ThrowsAsync<InvalidDataException>(() =>
            _sut.ConvertAsync(workbook, new MemoryStream()));

        Assert.Contains("client ID", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("row 2", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertAsync_ThrowsWhenPiecesAreNotNumeric()
    {
        await using var workbook = CreateWorkbook(
            ["Klient", "IČ", "Zakázka", "2025-01"],
            [["Acme", "123", "Project A", "not-a-number"]]);

        var exception = await Assert.ThrowsAsync<InvalidDataException>(() =>
            _sut.ConvertAsync(workbook, new MemoryStream()));

        Assert.Contains("not a valid number", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("column 4", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("row 2", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertAsync_ThrowsWhenWorksheetHasTooFewColumns()
    {
        await using var workbook = CreateWorkbook(
            ["Klient", "IČ", "Zakázka"],
            [["Acme", "123", "Project A"]]);

        var exception = await Assert.ThrowsAsync<InvalidDataException>(() =>
            _sut.ConvertAsync(workbook, new MemoryStream()));

        Assert.Contains("at least 4 columns", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertAsync_ThrowsWhenHeaderIsInvalid()
    {
        await using var workbook = CreateWorkbook(
            ["Customer", "IČ", "Zakázka", "2025-01"],
            [["Acme", "123", "Project A", "1"]]);

        var exception = await Assert.ThrowsAsync<InvalidDataException>(() =>
            _sut.ConvertAsync(workbook, new MemoryStream()));

        Assert.Contains("Klient", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("-1")]
    [InlineData("1.5")]
    public async Task ConvertAsync_ThrowsWhenPiecesAreNotNonNegativeIntegers(string pieces)
    {
        await using var workbook = CreateWorkbook(
            ["Klient", "IČ", "Zakázka", "2025-01"],
            [["Acme", "123", "Project A", pieces]]);

        var exception = await Assert.ThrowsAsync<InvalidDataException>(() =>
            _sut.ConvertAsync(workbook, new MemoryStream()));

        Assert.Contains("non-negative integer", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConvertAsync_ThrowsForNullInput()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            _sut.ConvertAsync(null!, new MemoryStream()));
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            _sut.ConvertAsync(new MemoryStream(), null!));
    }

    [Fact]
    public async Task ConvertAsync_ThrowsWhenCanceled()
    {
        await using var workbook = CreateWorkbook(
            ["Klient", "IČ", "Zakázka", "2025-01"], []);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            _sut.ConvertAsync(workbook, new MemoryStream(), cancellation.Token));
    }

    private static MemoryStream CreateWorkbook(
        string[] headers,
        string[][] rows)
    {
        var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            });

            sheetData.Append(CreateRow(headers));
            foreach (var row in rows)
            {
                sheetData.Append(CreateRow(row));
            }

            workbookPart.Workbook.Save();
        }

        stream.Position = 0;
        return stream;
    }

    private static Row CreateRow(IEnumerable<string> values)
    {
        var row = new Row();
        foreach (var value in values)
        {
            row.Append(new Cell
            {
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(value))
            });
        }

        return row;
    }
}
