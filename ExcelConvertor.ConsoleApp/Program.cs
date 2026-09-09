using ExcelConvertor.Lib;
using Serilog;

Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Information()
    .WriteTo.Console(
        outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
    .CreateLogger();

try
{
    string dir = Environment.CurrentDirectory;
    var files = Directory.EnumerateFiles(dir, "*.xlsx", SearchOption.TopDirectoryOnly).ToList();

    if (files.Count == 0)
    {
        Log.Information("No XLSX files found in {Directory}.", dir);
        return 0;
    }

    Log.Information("Found {FileCount} XLSX file(s) in {Directory}.", files.Count, dir);

    var converter = new ExcelToJsonConverter();
    int failedFiles = 0;

    foreach (var inPath in files)
    {
        string outPath = Path.ChangeExtension(inPath, ".json");
        string temporaryPath = $"{outPath}.{Guid.NewGuid():N}.tmp";

        try
        {
            await using var input = new FileStream(inPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            await using var output = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 64 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan);

            await converter.ConvertAsync(input, output);

            File.Move(temporaryPath, outPath, overwrite: true);
            Log.Information("Converted {InputFile} to {OutputFile}.", Path.GetFileName(inPath), Path.GetFileName(outPath));
        }
        catch (Exception exception)
        {
            failedFiles++;
            Log.Error(exception, "Failed to convert {InputFile}.", Path.GetFileName(inPath));
        }
        finally
        {
            File.Delete(temporaryPath);
        }
    }

    if (failedFiles > 0)
    {
        Log.Error("Conversion completed with {FailedFileCount} failure(s).", failedFiles);
        return 1;
    }

    Log.Information("Conversion completed successfully.");
    return 0;
}
catch (Exception exception)
{
    Log.Fatal(exception, "The conversion process failed.");
    return 1;
}
finally
{
    await Log.CloseAndFlushAsync();
}
