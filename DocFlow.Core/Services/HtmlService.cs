using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocFlow.Core.Helpers;
using DocFlow.Core.Interfaces;

namespace DocFlow.Core.Services
{
    public sealed class HtmlService : IHtmlService
    {
        private readonly IWordService _wordService;
        private readonly IExcelService _excelService;
        private readonly ILogger _logger;
        private readonly Models.DocFlowSettings _settings;

        public HtmlService(IWordService wordService, IExcelService excelService, ILogger logger = null, Models.DocFlowSettings settings = null)
        {
            _wordService = wordService ?? throw new ArgumentNullException(nameof(wordService));
            _excelService = excelService ?? throw new ArgumentNullException(nameof(excelService));
            _logger = logger ?? new NullLogger();
            _settings = settings ?? Models.DocFlowSettings.CreateDefault();
        }

        public string ReadHtml(string path)
        {
            FileHelper.EnsureInputFileExists(path);
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Html);
            FileHelper.EnsureFileSize(path, _settings.MaxFileSizeBytes);
            return HtmlHelper.ExtractText(File.ReadAllText(path));
        }

        public void ConvertHtmlToWord(string inputPath, string outputPath)
        {
            _wordService.CreateWord(outputPath, ReadHtml(inputPath));
        }

        public void ConvertHtmlToPdf(string inputPath, string outputPath)
        {
            var tempWord = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
            try
            {
                ConvertHtmlToWord(inputPath, tempWord);
                _wordService.ConvertWordToPdf(tempWord, outputPath);
            }
            finally
            {
                if (File.Exists(tempWord))
                {
                    File.Delete(tempWord);
                }
            }
        }

        public void ConvertHtmlToExcel(string inputPath, string outputPath)
        {
            var tables = ExtractTables(inputPath);
            if (tables.Count == 0)
            {
                tables = new List<Dictionary<string, string>>
                {
                    new Dictionary<string, string> { { "Content", ReadHtml(inputPath) } }
                };
            }

            _excelService.CreateExcel(outputPath, tables);
        }

        public List<Dictionary<string, string>> ExtractTables(string path)
        {
            FileHelper.EnsureInputFileExists(path);
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Html);
            FileHelper.EnsureFileSize(path, _settings.MaxFileSizeBytes);
            return HtmlHelper.ExtractTables(File.ReadAllText(path));
        }

        public Task ConvertHtmlToWordAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ConvertHtmlToWord(inputPath, outputPath), cancellationToken);
        }

        public Task ConvertHtmlToPdfAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ConvertHtmlToPdf(inputPath, outputPath), cancellationToken);
        }

        public Task ConvertHtmlToExcelAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ConvertHtmlToExcel(inputPath, outputPath), cancellationToken);
        }
    }
}
