using System;
using System.IO;
using System.Threading.Tasks;
using DotVVM.Framework.Controls;
using DotVVM.Framework.Hosting;
using DotVVM.Core.Storage;
using ExcelConvertor.Lib;

namespace ExcelConvertor.WebApp.ViewModels
{
    public class DefaultViewModel(
        IExcelToJsonConverter converter,
        IUploadedFileStorage uploadedFileStorage,
        IDotvvmRequestContext context) : MasterPageViewModel
    {

        public UploadedFilesCollection UploadedFiles { get; set; } = new();

        public async Task Convert()
        {
            if (UploadedFiles.Files.Count != 1)
            {
                throw new InvalidOperationException("Please select one XLSX file.");
            }

            var uploadedFile = UploadedFiles.Files[0];
            if (!uploadedFile.IsAllowed ||
                !string.Equals(Path.GetExtension(uploadedFile.FileName), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Please select a valid XLSX file.");
            }

            await using var xlsxStream = await uploadedFileStorage.GetFileAsync(uploadedFile.FileId);
            string temporaryPath = Path.Combine(
                Path.GetTempPath(),
                $"{Guid.NewGuid():N}.converted.json");

            try
            {
                await using var jsonStream = new FileStream(
                    temporaryPath,
                    FileMode.CreateNew,
                    FileAccess.ReadWrite,
                    FileShare.Read,
                    bufferSize: 64 * 1024,
                    options: FileOptions.Asynchronous | FileOptions.SequentialScan);

                await converter.ConvertAsync(xlsxStream, jsonStream);
                jsonStream.Position = 0;
                await context.ReturnFileAsync(jsonStream, "converted.json", "application/json");
            }
            finally
            {
                File.Delete(temporaryPath);
            }
        }

    }
}
