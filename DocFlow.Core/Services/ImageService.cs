using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocFlow.Core.Interfaces;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using Tesseract;

namespace DocFlow.Core.Services
{
    public sealed class ImageService : IImageService
    {
        private readonly IWordService _wordService;
        private readonly IExcelService _excelService;
        private readonly ILogger _logger;
        private readonly Models.DocFlowSettings _settings;

public ImageService(IWordService wordService, IExcelService excelService, ILogger logger = null, Models.DocFlowSettings settings = null)
        {
            _wordService = wordService ?? throw new ArgumentNullException(nameof(wordService));
            _excelService = excelService ?? throw new ArgumentNullException(nameof(excelService));
            _logger = logger ?? new Helpers.NullLogger();
            _settings = settings ?? Models.DocFlowSettings.CreateDefault();
        }

        public string ReadImageText(string path)
        {
            Helpers.FileHelper.EnsureInputFileExists(path);
            Helpers.FileHelper.EnsureDocumentType(path, Models.DocumentType.Image);
            Helpers.FileHelper.EnsureFileSize(path, _settings.MaxFileSizeBytes);

            var tessDataPath = ResolveTessDataPath();
            using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default))
            using (var image = Pix.LoadFromFile(path))
            using (var page = engine.Process(image))
            {
                return page.GetText() ?? string.Empty;
            }
        }

        public void ConvertImageToPdf(string inputPath, string outputPath)
        {
            Helpers.FileHelper.EnsureInputFileExists(inputPath);
            Helpers.FileHelper.EnsureDocumentType(inputPath, Models.DocumentType.Image);
            Helpers.FileHelper.EnsureDocumentType(outputPath, Models.DocumentType.Pdf);
            Helpers.FileHelper.EnsureCanWriteOutput(outputPath, _settings.AllowOverwrite);

            var bytes = File.ReadAllBytes(inputPath);
            using (var pdf = new PdfDocument())
            {
                var page = pdf.AddPage();
                page.Size = PdfSharpCore.PageSize.A4;
                using (var gfx = XGraphics.FromPdfPage(page))
                {
                    const double margin = 20;
                    var img = XImage.FromStream(() => new MemoryStream(bytes));
                    gfx.DrawImage(img, margin, margin, page.Width - margin * 2, page.Height - margin * 2);
                }
                pdf.Save(outputPath);
            }
        }

        public void ConvertImageToWord(string inputPath, string outputPath)
        {
            _wordService.CreateWord(outputPath, ReadImageText(inputPath));
        }

        public void ConvertImageToExcel(string inputPath, string outputPath)
        {
            var lines = ReadImageText(inputPath).Replace("\r\n", "\n").Split('\n');
            var data = new List<Dictionary<string, string>>();
            foreach (var line in lines)
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    data.Add(new Dictionary<string, string> { { "Content", line.Trim() } });
                }
            }

            if (data.Count == 0)
            {
                data.Add(new Dictionary<string, string> { { "Content", string.Empty } });
            }

            _excelService.CreateExcel(outputPath, data);
        }

        public Task<string> ReadImageTextAsync(string path, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ReadImageText(path), cancellationToken);
        }

        public Task ConvertImageToPdfAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ConvertImageToPdf(inputPath, outputPath), cancellationToken);
        }

        public Task ConvertImageToWordAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ConvertImageToWord(inputPath, outputPath), cancellationToken);
        }

        public Task ConvertImageToExcelAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ConvertImageToExcel(inputPath, outputPath), cancellationToken);
        }

        private string ResolveTessDataPath()
        {
            if (!string.IsNullOrWhiteSpace(_settings.OcrDataPath) && Directory.Exists(_settings.OcrDataPath))
            {
                return _settings.OcrDataPath;
            }

            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var candidate = Path.Combine(baseDirectory, "tessdata");
            if (Directory.Exists(candidate))
            {
                return candidate;
            }

            throw new InvalidOperationException("Tesseract data directory was not found. Configure OcrDataPath or deploy a tessdata folder.");
        }
    }
}
