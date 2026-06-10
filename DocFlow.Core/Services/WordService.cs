using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocFlow.Core.Helpers;
using DocFlow.Core.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocFlow.Core.Services
{
    public sealed class WordService : IWordService
    {
        private readonly ILogger _logger;
        private readonly Models.DocFlowSettings _settings;

        static WordService()
        {
            QuestPDF.Settings.License = LicenseType.Community;
        }

        public WordService(ILogger logger = null, Models.DocFlowSettings settings = null)
        {
            _logger = logger ?? new NullLogger();
            _settings = settings ?? Models.DocFlowSettings.CreateDefault();
        }

        public void CreateWord(string path, string content)
        {
            FileHelper.EnsurePath(path, nameof(path));
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Word);
            FileHelper.EnsureCanWriteOutput(path, _settings.AllowOverwrite);
            LoggingHelper.LogStart(_logger, _settings, "CreateWord", path);

            using (var stream = File.Create(path))
            {
                CreateWord(stream, content);
            }

            LoggingHelper.LogEnd(_logger, _settings, "CreateWord", path);
        }

        public void CreateWord(Stream output, string content)
        {
            StreamHelper.EnsureWritable(output, nameof(output));
            if (content == null)
            {
                throw new ArgumentNullException(nameof(content));
            }

            try
            {
                if (output.CanSeek)
                {
                    output.SetLength(0);
                    output.Position = 0;
                }

                using (var document = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document, true))
                {
                    var mainPart = document.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    EnsureStyles(mainPart);

                    var body = new Body();
                    foreach (var block in TextStructureHelper.ParseBlocks(content))
                    {
                        body.Append(CreateParagraph(block));
                    }

                    mainPart.Document.Append(body);
                    mainPart.Document.Save();
                }

                if (output.CanSeek)
                {
                    output.Position = 0;
                }
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to create Word document.", exception);
                throw new InvalidOperationException("Word document creation failed.", exception);
            }
        }

        public byte[] CreateWord(string content)
        {
            using (var stream = new MemoryStream())
            {
                CreateWord(stream, content);
                return stream.ToArray();
            }
        }

        public Task CreateWordAsync(string path, string content, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                CreateWord(path, content);
            }, cancellationToken);
        }

        public Task CreateWordAsync(Stream output, string content, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                CreateWord(output, content);
            }, cancellationToken);
        }

        public Task<byte[]> CreateWordAsync(string content, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return CreateWord(content);
            }, cancellationToken);
        }

        public string ReadWord(string path)
        {
            FileHelper.EnsureInputFileExists(path);
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Word);
            FileHelper.EnsureFileSize(path, _settings.MaxFileSizeBytes);
            LoggingHelper.LogStart(_logger, _settings, "ReadWord", path);

            using (var stream = File.OpenRead(path))
            {
                var result = ReadWord(stream);
                LoggingHelper.LogEnd(_logger, _settings, "ReadWord", path);
                return result;
            }
        }

        public string ReadWord(Stream input)
        {
            return SerializeBlocks(ReadBlocks(input));
        }

        public string ReadWord(byte[] inputBytes)
        {
            using (var stream = StreamHelper.ToMemoryStream(inputBytes, nameof(inputBytes)))
            {
                return ReadWord(stream);
            }
        }

        public void ReplacePlaceholders(string templatePath, string outputPath, IDictionary<string, string> placeholders)
        {
            FileHelper.EnsureInputFileExists(templatePath);
            FileHelper.EnsureDocumentType(templatePath, Models.DocumentType.Word);
            FileHelper.EnsureDocumentType(outputPath, Models.DocumentType.Word);
            FileHelper.EnsureFileSize(templatePath, _settings.MaxFileSizeBytes);
            FileHelper.EnsureCanWriteOutput(outputPath, _settings.AllowOverwrite);
            LoggingHelper.LogStart(_logger, _settings, "ReplaceWordPlaceholders", templatePath + " => " + outputPath);

            using (var input = File.OpenRead(templatePath))
            using (var output = File.Create(outputPath))
            {
                ReplacePlaceholders(input, output, placeholders);
            }

            LoggingHelper.LogEnd(_logger, _settings, "ReplaceWordPlaceholders", outputPath);
        }

        public void ReplacePlaceholders(Stream templateStream, Stream outputStream, IDictionary<string, string> placeholders)
        {
            StreamHelper.EnsureReadable(templateStream, nameof(templateStream));
            StreamHelper.EnsureWritable(outputStream, nameof(outputStream));
            PlaceholderHelper.ValidatePlaceholders(placeholders, nameof(placeholders));

            try
            {
                using (var seekable = StreamHelper.EnsureSeekable(templateStream))
                {
                    var bytes = seekable.ToArray();
                    outputStream.Write(bytes, 0, bytes.Length);
                    if (outputStream.CanSeek)
                    {
                        outputStream.Position = 0;
                    }

                    using (var document = WordprocessingDocument.Open(outputStream, true))
                    {
                        ReplacePlaceholdersInWordDocument(document, placeholders);
                        document.MainDocumentPart.Document.Save();
                    }

                    if (outputStream.CanSeek)
                    {
                        outputStream.Position = 0;
                    }
                }
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to replace placeholders in Word document.", exception);
                throw new InvalidOperationException("Word placeholder replacement failed.", exception);
            }
        }

        public byte[] ReplacePlaceholders(byte[] templateBytes, IDictionary<string, string> placeholders)
        {
            using (var input = StreamHelper.ToMemoryStream(templateBytes, nameof(templateBytes)))
            using (var output = new MemoryStream())
            {
                ReplacePlaceholders(input, output, placeholders);
                return output.ToArray();
            }
        }

        public Task<string> ReadWordAsync(string path, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReadWord(path);
            }, cancellationToken);
        }

        public Task<string> ReadWordAsync(Stream input, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReadWord(input);
            }, cancellationToken);
        }

        public Task<string> ReadWordAsync(byte[] inputBytes, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReadWord(inputBytes);
            }, cancellationToken);
        }

        public Task ReplacePlaceholdersAsync(string templatePath, string outputPath, IDictionary<string, string> placeholders, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ReplacePlaceholders(templatePath, outputPath, placeholders);
            }, cancellationToken);
        }

        public Task ReplacePlaceholdersAsync(Stream templateStream, Stream outputStream, IDictionary<string, string> placeholders, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ReplacePlaceholders(templateStream, outputStream, placeholders);
            }, cancellationToken);
        }

        public Task<byte[]> ReplacePlaceholdersAsync(byte[] templateBytes, IDictionary<string, string> placeholders, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReplacePlaceholders(templateBytes, placeholders);
            }, cancellationToken);
        }

        public void ConvertWordToPdf(string inputPath, string outputPath)
        {
            FileHelper.EnsureInputFileExists(inputPath);
            FileHelper.EnsureDocumentType(inputPath, Models.DocumentType.Word);
            FileHelper.EnsureDocumentType(outputPath, Models.DocumentType.Pdf);
            FileHelper.EnsureFileSize(inputPath, _settings.MaxFileSizeBytes);
            FileHelper.EnsureCanWriteOutput(outputPath, _settings.AllowOverwrite);
            LoggingHelper.LogStart(_logger, _settings, "ConvertWordToPdf", inputPath + " => " + outputPath);

            using (var input = File.OpenRead(inputPath))
            using (var output = File.Create(outputPath))
            {
                ConvertWordToPdf(input, output);
            }

            LoggingHelper.LogEnd(_logger, _settings, "ConvertWordToPdf", outputPath);
        }

        public void ConvertWordToPdf(Stream input, Stream output)
        {
            StreamHelper.EnsureReadable(input, nameof(input));
            StreamHelper.EnsureWritable(output, nameof(output));

            try
            {
                RenderBlocksToPdf(ReadBlocks(input), output);
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to convert Word document to PDF.", exception);
                throw new InvalidOperationException("Word to PDF conversion failed.", exception);
            }
        }

        public byte[] ConvertWordToPdf(byte[] inputBytes)
        {
            using (var input = StreamHelper.ToMemoryStream(inputBytes, nameof(inputBytes)))
            using (var output = new MemoryStream())
            {
                ConvertWordToPdf(input, output);
                return output.ToArray();
            }
        }

        public Task ConvertWordToPdfAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ConvertWordToPdf(inputPath, outputPath);
            }, cancellationToken);
        }

        public Task ConvertWordToPdfAsync(Stream input, Stream output, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ConvertWordToPdf(input, output);
            }, cancellationToken);
        }

        public Task<byte[]> ConvertWordToPdfAsync(byte[] inputBytes, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ConvertWordToPdf(inputBytes);
            }, cancellationToken);
        }

        internal IList<TextBlock> ReadBlocks(Stream input)
        {
            StreamHelper.EnsureReadable(input, nameof(input));

            try
            {
                using (var seekable = StreamHelper.EnsureSeekable(input))
                using (var document = WordprocessingDocument.Open(seekable, false))
                {
                    var blocks = new List<TextBlock>();
                    foreach (var paragraph in document.MainDocumentPart.Document.Body.Elements<Paragraph>())
                    {
                        var text = string.Concat(paragraph.Descendants<Text>().Select(t => t.Text)).Trim();
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            continue;
                        }

                        var styleId = paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.ParagraphStyleId != null
                            ? paragraph.ParagraphProperties.ParagraphStyleId.Val.Value
                            : string.Empty;

                        var level = ResolveHeadingLevel(styleId);
                        blocks.Add(level > 0
                            ? new TextBlock(TextBlockType.Heading, text, level)
                            : new TextBlock(TextBlockType.Paragraph, text, 0));
                    }

                    return blocks;
                }
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to read Word document.", exception);
                throw new InvalidOperationException("Word document read failed.", exception);
            }
        }

        internal string SerializeBlocks(IEnumerable<TextBlock> blocks)
        {
            var lines = new List<string>();
            foreach (var block in blocks)
            {
                lines.Add(block.Type == TextBlockType.Heading
                    ? new string('#', Math.Max(1, Math.Min(3, block.Level))) + " " + block.Text
                    : block.Text);
                lines.Add(string.Empty);
            }

            return string.Join(Environment.NewLine, lines).Trim();
        }

        private static void EnsureStyles(MainDocumentPart mainPart)
        {
            var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                CreateStyle("Normal", "Normal", false, "22"),
                CreateStyle("Heading1", "Heading 1", true, "32"),
                CreateStyle("Heading2", "Heading 2", true, "28"),
                CreateStyle("Heading3", "Heading 3", true, "24"));
        }

        private static Style CreateStyle(string styleId, string styleName, bool bold, string fontSize)
        {
            var style = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            };

            style.Append(new StyleName { Val = styleName });
            style.Append(new BasedOn { Val = "Normal" });
            style.Append(new PrimaryStyle());
            style.Append(new StyleRunProperties(
                bold ? new Bold() : null,
                new FontSize { Val = fontSize }));

            return style;
        }

        private static Paragraph CreateParagraph(TextBlock block)
        {
            var paragraph = new Paragraph();
            var properties = new ParagraphProperties(new SpacingBetweenLines
            {
                Before = block.Type == TextBlockType.Heading ? "120" : "80",
                After = block.Type == TextBlockType.Heading ? "160" : "120"
            });

            if (block.Type == TextBlockType.Heading)
            {
                properties.Append(new ParagraphStyleId { Val = "Heading" + Math.Max(1, Math.Min(3, block.Level)) });
            }

            var runProperties = new RunProperties();
            if (block.Type == TextBlockType.Heading)
            {
                runProperties.Append(new Bold());
                runProperties.Append(new FontSize { Val = block.Level == 1 ? "32" : block.Level == 2 ? "28" : "24" });
            }
            else
            {
                runProperties.Append(new FontSize { Val = "22" });
            }

            paragraph.Append(properties);
            paragraph.Append(new Run(runProperties, new Text(block.Text) { Space = SpaceProcessingModeValues.Preserve }));
            return paragraph;
        }

        private static int ResolveHeadingLevel(string styleId)
        {
            if (string.Equals(styleId, "Heading1", StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }

            if (string.Equals(styleId, "Heading2", StringComparison.OrdinalIgnoreCase))
            {
                return 2;
            }

            if (string.Equals(styleId, "Heading3", StringComparison.OrdinalIgnoreCase))
            {
                return 3;
            }

            return 0;
        }

        private static void RenderBlocksToPdf(IList<TextBlock> blocks, Stream output)
        {
            if (output.CanSeek)
            {
                output.SetLength(0);
                output.Position = 0;
            }

            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Margin(36);
                    page.Size(PageSizes.A4);
                    page.DefaultTextStyle(x => x.FontSize(11));
                    page.Content().Column(column =>
                    {
                        foreach (var block in blocks)
                        {
                            if (block.Type == TextBlockType.Heading)
                            {
                                column.Item().PaddingBottom(8).Text(block.Text).FontSize(block.Level == 1 ? 18 : block.Level == 2 ? 16 : 14).Bold();
                            }
                            else
                            {
                                column.Item().PaddingBottom(10).Text(block.Text).FontSize(11);
                            }
                        }
                    });
                });
            }).GeneratePdf(output);

            if (output.CanSeek)
            {
                output.Position = 0;
            }
        }

        private static void ReplacePlaceholdersInWordDocument(WordprocessingDocument document, IDictionary<string, string> placeholders)
        {
            var texts = document.MainDocumentPart.Document.Descendants<Text>().ToList();
            foreach (var text in texts)
            {
                text.Text = PlaceholderHelper.ReplacePlaceholders(text.Text, placeholders);
            }

            foreach (var headerPart in document.MainDocumentPart.HeaderParts)
            {
                foreach (var text in headerPart.RootElement.Descendants<Text>())
                {
                    text.Text = PlaceholderHelper.ReplacePlaceholders(text.Text, placeholders);
                }
            }

            foreach (var footerPart in document.MainDocumentPart.FooterParts)
            {
                foreach (var text in footerPart.RootElement.Descendants<Text>())
                {
                    text.Text = PlaceholderHelper.ReplacePlaceholders(text.Text, placeholders);
                }
            }
        }
    }
}
