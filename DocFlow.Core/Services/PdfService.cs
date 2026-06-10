using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocFlow.Core.Helpers;
using DocFlow.Core.Interfaces;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace DocFlow.Core.Services
{
    public sealed class PdfService : IPdfService
    {
        private readonly IWordService _wordService;
        private readonly IExcelService _excelService;
        private readonly ILogger _logger;
        private readonly Models.DocFlowSettings _settings;

        public PdfService(IWordService wordService, IExcelService excelService, ILogger logger = null, Models.DocFlowSettings settings = null)
        {
            _wordService = wordService ?? throw new ArgumentNullException(nameof(wordService));
            _excelService = excelService ?? throw new ArgumentNullException(nameof(excelService));
            _logger = logger ?? new NullLogger();
            _settings = settings ?? Models.DocFlowSettings.CreateDefault();
        }

        public string ReadPdf(string path)
        {
            FileHelper.EnsureInputFileExists(path);
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Pdf);
            FileHelper.EnsureFileSize(path, _settings.MaxFileSizeBytes);
            LoggingHelper.LogStart(_logger, _settings, "ReadPdf", path);

            using (var stream = File.OpenRead(path))
            {
                var result = ReadPdf(stream);
                LoggingHelper.LogEnd(_logger, _settings, "ReadPdf", path);
                return result;
            }
        }

        public string ReadPdf(Stream input)
        {
            StreamHelper.EnsureReadable(input, nameof(input));

            try
            {
                return string.Join(Environment.NewLine, ExtractLines(input)).Trim();
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to read PDF document.", exception);
                throw new InvalidOperationException("PDF document read failed.", exception);
            }
        }

        public string ReadPdf(byte[] inputBytes)
        {
            using (var stream = StreamHelper.ToMemoryStream(inputBytes, nameof(inputBytes)))
            {
                return ReadPdf(stream);
            }
        }

        public IList<string> ExtractImages(string inputPath, string outputDirectory)
        {
            FileHelper.EnsureInputFileExists(inputPath);
            FileHelper.EnsureDocumentType(inputPath, Models.DocumentType.Pdf);
            Directory.CreateDirectory(outputDirectory);

            var savedFiles = new List<string>();
            using (var document = PdfDocument.Open(inputPath))
            {
                var pageNumber = 0;
                foreach (var page in document.GetPages())
                {
                    pageNumber++;
                    var imageIndex = 0;
                    foreach (var image in page.GetImages())
                    {
                        imageIndex++;
                        var filePath = Path.Combine(outputDirectory, string.Format("page-{0}-image-{1}.bin", pageNumber, imageIndex));
                        byte[] bytes;
                        if (image.TryGetPng(out bytes))
                        {
                            filePath = Path.Combine(outputDirectory, string.Format("page-{0}-image-{1}.png", pageNumber, imageIndex));
                        }
                        else
                        {
                            IReadOnlyList<byte> rawList;
                            bytes = image.TryGetBytes(out rawList) ? rawList.ToArray() : image.RawBytes.ToArray();
                        }

                        File.WriteAllBytes(filePath, bytes);
                        savedFiles.Add(filePath);
                    }
                }
            }

            return savedFiles;
        }

        public Task<string> ReadPdfAsync(string path, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReadPdf(path);
            }, cancellationToken);
        }

        public Task<string> ReadPdfAsync(Stream input, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReadPdf(input);
            }, cancellationToken);
        }

        public Task<string> ReadPdfAsync(byte[] inputBytes, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReadPdf(inputBytes);
            }, cancellationToken);
        }

        public Task<IList<string>> ExtractImagesAsync(string inputPath, string outputDirectory, CancellationToken cancellationToken = default)
        {
            return Task.Run(() => ExtractImages(inputPath, outputDirectory), cancellationToken);
        }

        public void ConvertPdfToWord(string inputPath, string outputPath)
        {
            FileHelper.EnsureInputFileExists(inputPath);
            FileHelper.EnsureDocumentType(inputPath, Models.DocumentType.Pdf);
            FileHelper.EnsureDocumentType(outputPath, Models.DocumentType.Word);
            FileHelper.EnsureFileSize(inputPath, _settings.MaxFileSizeBytes);
            FileHelper.EnsureCanWriteOutput(outputPath, _settings.AllowOverwrite);
            LoggingHelper.LogStart(_logger, _settings, "ConvertPdfToWord", inputPath + " => " + outputPath);

            using (var input = File.OpenRead(inputPath))
            using (var output = File.Create(outputPath))
            {
                ConvertPdfToWord(input, output);
            }

            LoggingHelper.LogEnd(_logger, _settings, "ConvertPdfToWord", outputPath);
        }

        public void ConvertPdfToWord(Stream input, Stream output)
        {
            StreamHelper.EnsureReadable(input, nameof(input));
            StreamHelper.EnsureWritable(output, nameof(output));

            try
            {
                var blocks = BuildBlocksFromPdf(input);
                _wordService.CreateWord(output, SerializeBlocks(blocks));
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to convert PDF to Word.", exception);
                throw new InvalidOperationException("PDF to Word conversion failed.", exception);
            }
        }

        public byte[] ConvertPdfToWord(byte[] inputBytes)
        {
            using (var input = StreamHelper.ToMemoryStream(inputBytes, nameof(inputBytes)))
            using (var output = new MemoryStream())
            {
                ConvertPdfToWord(input, output);
                return output.ToArray();
            }
        }

        public Task ConvertPdfToWordAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ConvertPdfToWord(inputPath, outputPath);
            }, cancellationToken);
        }

        public Task ConvertPdfToWordAsync(Stream input, Stream output, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ConvertPdfToWord(input, output);
            }, cancellationToken);
        }

        public Task<byte[]> ConvertPdfToWordAsync(byte[] inputBytes, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ConvertPdfToWord(inputBytes);
            }, cancellationToken);
        }

        public void ConvertPdfToExcel(string inputPath, string outputPath)
        {
            FileHelper.EnsureInputFileExists(inputPath);
            FileHelper.EnsureDocumentType(inputPath, Models.DocumentType.Pdf);
            FileHelper.EnsureDocumentType(outputPath, Models.DocumentType.Excel);
            FileHelper.EnsureFileSize(inputPath, _settings.MaxFileSizeBytes);
            FileHelper.EnsureCanWriteOutput(outputPath, _settings.AllowOverwrite);
            LoggingHelper.LogStart(_logger, _settings, "ConvertPdfToExcel", inputPath + " => " + outputPath);

            using (var input = File.OpenRead(inputPath))
            using (var output = File.Create(outputPath))
            {
                ConvertPdfToExcel(input, output);
            }

            LoggingHelper.LogEnd(_logger, _settings, "ConvertPdfToExcel", outputPath);
        }

        public void ConvertPdfToExcel(Stream input, Stream output)
        {
            StreamHelper.EnsureReadable(input, nameof(input));
            StreamHelper.EnsureWritable(output, nameof(output));

            try
            {
                _excelService.CreateExcel(output, BuildTabularDataFromPdf(input));
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to convert PDF to Excel.", exception);
                throw new InvalidOperationException("PDF to Excel conversion failed.", exception);
            }
        }

        public byte[] ConvertPdfToExcel(byte[] inputBytes)
        {
            using (var input = StreamHelper.ToMemoryStream(inputBytes, nameof(inputBytes)))
            using (var output = new MemoryStream())
            {
                ConvertPdfToExcel(input, output);
                return output.ToArray();
            }
        }

        public Task ConvertPdfToExcelAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ConvertPdfToExcel(inputPath, outputPath);
            }, cancellationToken);
        }

        public Task ConvertPdfToExcelAsync(Stream input, Stream output, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ConvertPdfToExcel(input, output);
            }, cancellationToken);
        }

        public Task<byte[]> ConvertPdfToExcelAsync(byte[] inputBytes, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ConvertPdfToExcel(inputBytes);
            }, cancellationToken);
        }

        private static string SerializeBlocks(IEnumerable<TextBlock> blocks)
        {
            var builder = new StringBuilder();
            foreach (var block in blocks)
            {
                if (block.Type == TextBlockType.Heading)
                {
                    builder.Append(new string('#', Math.Max(1, Math.Min(3, block.Level))));
                    builder.Append(' ');
                }

                builder.AppendLine(block.Text);
                builder.AppendLine();
            }

            return builder.ToString().Trim();
        }

        private IList<TextBlock> BuildBlocksFromPdf(Stream input)
        {
            return TextStructureHelper.ParseBlocks(string.Join(Environment.NewLine, ExtractLines(input)));
        }

        private List<Dictionary<string, string>> BuildTabularDataFromPdf(Stream input)
        {
            var lines = ExtractLines(input);
            var tableRows = new List<List<string>>();

            foreach (var line in lines)
            {
                List<string> cells;
                if (TextStructureHelper.TrySplitTableRow(line, out cells))
                {
                    tableRows.Add(cells);
                }
            }

            if (tableRows.Count == 0)
            {
                var fallback = lines
                    .Where(line => !string.IsNullOrWhiteSpace(line))
                    .Select(line => new Dictionary<string, string> { { "Content", line } })
                    .ToList();

                return fallback.Count > 0
                    ? fallback
                    : new List<Dictionary<string, string>> { new Dictionary<string, string> { { "Content", string.Empty } } };
            }

            var columnCount = tableRows.Max(row => row.Count);
            foreach (var row in tableRows)
            {
                while (row.Count < columnCount)
                {
                    row.Add(string.Empty);
                }
            }

            var headers = tableRows[0]
                .Select((value, index) => string.IsNullOrWhiteSpace(value) ? "Column" + (index + 1) : value.Trim())
                .ToList();

            var data = new List<Dictionary<string, string>>();
            for (var rowIndex = 1; rowIndex < tableRows.Count; rowIndex++)
            {
                var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    row[headers[columnIndex]] = tableRows[rowIndex][columnIndex];
                }

                data.Add(row);
            }

            if (data.Count == 0)
            {
                var emptyRow = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (var index = 0; index < headers.Count; index++)
                {
                    emptyRow[headers[index]] = string.Empty;
                }

                data.Add(emptyRow);
            }

            return data;
        }

        private IList<string> ExtractLines(Stream input)
        {
            using (var seekable = StreamHelper.EnsureSeekable(input))
            using (var document = PdfDocument.Open(seekable))
            {
                var lines = new List<string>();
                foreach (var page in document.GetPages())
                {
                    lines.AddRange(ReconstructPageLines(page));
                    lines.Add(string.Empty);
                }

                while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[lines.Count - 1]))
                {
                    lines.RemoveAt(lines.Count - 1);
                }

                return lines;
            }
        }

        private static IList<string> ReconstructPageLines(Page page)
        {
            var groups = new List<LineGroup>();
            foreach (var letter in page.Letters
                         .Where(letter => !string.IsNullOrEmpty(letter.Value) && !char.IsControl(letter.Value[0]))
                         .OrderByDescending(letter => letter.GlyphRectangle.Bottom)
                         .ThenBy(letter => letter.GlyphRectangle.Left))
            {
                var line = groups.FirstOrDefault(group => Math.Abs(group.Y - letter.GlyphRectangle.Bottom) <= 2.0);
                if (line == null)
                {
                    line = new LineGroup(letter.GlyphRectangle.Bottom);
                    groups.Add(line);
                }

                line.Letters.Add(letter);
            }

            return groups
                .OrderByDescending(group => group.Y)
                .Select(BuildLineText)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();
        }

        private static string BuildLineText(LineGroup group)
        {
            var letters = group.Letters.OrderBy(letter => letter.GlyphRectangle.Left).ToList();
            if (letters.Count == 0)
            {
                return string.Empty;
            }

            var averageWidth = letters.Average(letter => Math.Max(1.0, letter.GlyphRectangle.Width));
            var builder = new StringBuilder();
            var first = true;
            double previousRight = 0;

            foreach (var letter in letters)
            {
                if (!first)
                {
                    var gap = letter.GlyphRectangle.Left - previousRight;
                    if (gap > averageWidth * 3)
                    {
                        builder.Append("   ");
                    }
                    else if (gap > averageWidth * 1.2)
                    {
                        builder.Append(' ');
                    }
                }

                builder.Append(!string.IsNullOrEmpty(letter.Value) && char.IsWhiteSpace(letter.Value[0]) ? " " : letter.Value);
                previousRight = letter.GlyphRectangle.Right;
                first = false;
            }

            return builder.ToString().Trim();
        }

        private sealed class LineGroup
        {
            public LineGroup(double y)
            {
                Y = y;
                Letters = new List<Letter>();
            }

            public double Y { get; private set; }

            public List<Letter> Letters { get; private set; }
        }
    }
}
