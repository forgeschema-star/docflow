using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocFlow.Core.Helpers;
using DocFlow.Core.Interfaces;

namespace DocFlow.Core.Services
{
    public sealed class CsvService : ICsvService
    {
        private readonly IExcelService _excelService;
        private readonly ILogger _logger;
        private readonly Models.DocFlowSettings _settings;

        public CsvService(IExcelService excelService, ILogger logger = null, Models.DocFlowSettings settings = null)
        {
            _excelService = excelService ?? throw new ArgumentNullException(nameof(excelService));
            _logger = logger ?? new NullLogger();
            _settings = settings ?? Models.DocFlowSettings.CreateDefault();
        }

        public void CreateCsv(string path, List<Dictionary<string, string>> data)
        {
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Csv);
            FileHelper.EnsureCanWriteOutput(path, _settings.AllowOverwrite);
            File.WriteAllText(path, CsvHelper.Serialize(data));
        }

        public List<Dictionary<string, string>> ReadCsv(string path)
        {
            FileHelper.EnsureInputFileExists(path);
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Csv);
            FileHelper.EnsureFileSize(path, _settings.MaxFileSizeBytes);
            return CsvHelper.Deserialize(File.ReadAllText(path));
        }

        public void ConvertCsvToExcel(string inputPath, string outputPath)
        {
            _excelService.CreateExcel(outputPath, ReadCsv(inputPath));
        }

        public void ConvertExcelToCsv(string inputPath, string outputPath)
        {
            CreateCsv(outputPath, _excelService.ReadExcel(inputPath));
        }

        public void ConvertCsvToPdf(string inputPath, string outputPath)
        {
            var tempExcel = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                ConvertCsvToExcel(inputPath, tempExcel);
                _excelService.ConvertExcelToPdf(tempExcel, outputPath);
            }
            finally
            {
                if (File.Exists(tempExcel))
                {
                    File.Delete(tempExcel);
                }
            }
        }

        public Task ConvertCsvToExcelAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ConvertCsvToExcel(inputPath, outputPath), cancellationToken);
        }

        public Task ConvertExcelToCsvAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ConvertExcelToCsv(inputPath, outputPath), cancellationToken);
        }

        public Task ConvertCsvToPdfAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ConvertCsvToPdf(inputPath, outputPath), cancellationToken);
        }
    }
}
