using System.Globalization;
using System.Text;
using System.Text.Json;
using ExcelDataReader;

namespace ExcelConvertor.Lib;

/// <summary>
/// Provides functionality to convert an XLSX worksheet to JSON without buffering the complete input workbook or output document in memory.
/// Be aware that the worksheet must be sorted by client before processing. Rows for the same client must be contiguous; otherwise, the converter cannot close a client in the output stream without retaining previously written data, and the client will be emitted as multiple groups.
/// </summary>
public sealed class ExcelToJsonConverter : IExcelToJsonConverter
{
    private static readonly (int Index, string Name) ClientNameColumn = (0, "Klient");
    private static readonly (int Index, string Name) ClientIdColumn = (1, "IČ");
    private static readonly (int Index, string Name) ProjectNameColumn = (2, "Zakázka");
    private const int DatesColumnsStart = 3;

    static ExcelToJsonConverter()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); //to enable Windows-1252
    }

    /// <summary>
    /// Converts an XLSX worksheet to JSON
    /// </summary>
    /// <param name="xlsxStream"></param>
    /// <param name="jsonOutput"></param>
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    /// <exception cref="InvalidDataException"></exception>
    public async Task ConvertAsync(
        Stream xlsxStream,
        Stream jsonOutput,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(xlsxStream);
        ArgumentNullException.ThrowIfNull(jsonOutput);

        cancellationToken.ThrowIfCancellationRequested();

        using var r = ExcelReaderFactory.CreateReader(xlsxStream);

        if (!r.Read())
        {
            throw new InvalidDataException("The worksheet does not contain any data rows.");
        }

        if (r.FieldCount < DatesColumnsStart + 1)
        {
            throw new InvalidDataException(
                $"The worksheet must contain at least {DatesColumnsStart + 1} columns.");
        }

        ValidateHeader(r, ClientNameColumn);
        ValidateHeader(r, ClientIdColumn);
        ValidateHeader(r, ProjectNameColumn);

        //read header rows for dates (allow any arbitrary date format, we dont care)
        var periods = new string[r.FieldCount];
        for (int i = DatesColumnsStart; i < r.FieldCount; i++)
        {
            periods[i] = ReadString(r, i) ?? throw new InvalidDataException($"The period header in worksheet column {i + 1} is missing.");
        }

        using var writer = new Utf8JsonWriter(jsonOutput, new() { Indented = true });

        writer.WriteStartArray();

        bool clientStarted = false;
        string currentClientId = string.Empty;

        while (r.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();

            string? clientName = ReadString(r, ClientNameColumn.Index);
            if (string.IsNullOrWhiteSpace(clientName))
            {
                throw new InvalidDataException(
                    $"The client name is missing on worksheet row {r.Depth + 1}.");
            }

            string? clientId = ReadString(r, ClientIdColumn.Index);
            if (string.IsNullOrWhiteSpace(clientId))
            {
                throw new InvalidDataException(
                    $"The client ID is missing on worksheet row {r.Depth + 1}.");
            }

            string? projectName = ReadString(r, ProjectNameColumn.Index);
            if (string.IsNullOrWhiteSpace(projectName))
            {
                throw new InvalidDataException(
                    $"The project name is missing on worksheet row {r.Depth + 1}.");
            }

            if (!clientStarted || !StringComparer.OrdinalIgnoreCase.Equals(clientId, currentClientId))
            {
                if (clientStarted)
                {
                    writer.WriteEndArray();
                    writer.WriteEndObject();
                    await writer.FlushAsync(cancellationToken);
                }

                writer.WriteStartObject();
                writer.WriteString("nazevKlienta", clientName);
                writer.WriteString("icoKlienta", clientId);
                writer.WriteStartArray("zakazky");

                currentClientId = clientId;
                clientStarted = true;
            }

            writer.WriteStartObject();
            writer.WriteString("nazevZakazky", projectName);
            writer.WriteStartArray("vyrobeneKusy");

            for (int i = DatesColumnsStart; i < r.FieldCount; i++)
            {
                long? pieces = ReadPieces(r, i);
                if (!pieces.HasValue)
                {
                    continue;
                }

                writer.WriteStartObject();
                writer.WriteString("obdobi", periods[i]);
                writer.WriteNumber("pocetKusu", pieces.Value);
                writer.WriteEndObject();
            }

            writer.WriteEndArray();
            writer.WriteEndObject();
        }

        if (clientStarted)
        {
            writer.WriteEndArray();
            writer.WriteEndObject();
        }

        writer.WriteEndArray();
        await writer.FlushAsync(cancellationToken);
    }


    private static void ValidateHeader(IExcelDataReader reader, (int Index, string Name) column)
    {
        var actual = ReadString(reader, column.Index);
        if (!string.Equals(actual, column.Name, StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                $"Worksheet column {column.Index + 1} must have the header '{column.Name}'.");
        }
    }

    private static string? ReadString(IExcelDataReader reader, int column)
    {
        var value = reader.GetValue(column);

        return value switch
        {
            null => null,
            DateTime dateTime => dateTime.ToString(
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture),
            _ => value.ToString()?.Trim()
        };
    }

    private static long? ReadPieces(IExcelDataReader reader, int column)
    {
        var value = reader.GetValue(column);
        if (value is null || string.IsNullOrWhiteSpace(value.ToString()))
        {
            return null;
        }

        string text = value is IFormattable formattable
            ? formattable.ToString(null, CultureInfo.InvariantCulture) ?? string.Empty
            : value.ToString() ?? string.Empty;

        if (!long.TryParse(text, NumberStyles.Integer | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out long pieces)
            || pieces < 0
            || pieces > long.MaxValue)
        {
            throw new InvalidDataException(
                $"The value '{value}' in worksheet column {column + 1}, row {reader.Depth + 1} is not a valid number; expected a non-negative integer.");
        }

        return pieces;
    }


}