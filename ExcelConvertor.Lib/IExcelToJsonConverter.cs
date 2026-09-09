namespace ExcelConvertor.Lib;

public interface IExcelToJsonConverter
{
    /// <summary>
    /// Converts an XLSX worksheet to JSON without buffering the complete input
    /// workbook or output document in memory.
    /// </summary>
    /// <remarks>
    /// The worksheet must be sorted by client before processing. Rows for the
    /// same client must be contiguous; otherwise the converter cannot close a
    /// client in the output stream without retaining previously written data,
    /// and the client will be emitted as multiple groups.
    /// </remarks>
    Task ConvertAsync(
        Stream xlsxStream,
        Stream jsonOutput,
        CancellationToken cancellationToken = default);
}
