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
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

namespace DocFlow.Core.Services
{
    public sealed class WordService : IWordService
    {
        private readonly ILogger _logger;
        private readonly Models.DocFlowSettings _settings;

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

        // ── Styled WordBlock overloads ───────────────────────────────────────

        public void CreateWord(Stream output, System.Collections.Generic.IList<Models.WordBlock> blocks)
        {
            StreamHelper.EnsureWritable(output, nameof(output));
            if (blocks == null) { throw new ArgumentNullException(nameof(blocks)); }

            try
            {
                if (output.CanSeek) { output.SetLength(0); output.Position = 0; }

                using (var document = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document, true))
                {
                    var mainPart = document.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    EnsureStyles(mainPart);
                    var body = new Body();

                    foreach (var block in blocks)
                    {
                        body.Append(CreateStyledParagraph(block));
                    }

                    mainPart.Document.Append(body);
                    mainPart.Document.Save();
                }

                if (output.CanSeek) { output.Position = 0; }
            }
            catch (Exception ex)
            {
                _logger.LogError("Failed to create styled Word document.", ex);
                throw new InvalidOperationException("Word document creation failed.", ex);
            }
        }

        public byte[] CreateWord(System.Collections.Generic.IList<Models.WordBlock> blocks)
        {
            using (var ms = new MemoryStream()) { CreateWord(ms, blocks); return ms.ToArray(); }
        }

        public void CreateWord(string path, System.Collections.Generic.IList<Models.WordBlock> blocks)
        {
            FileHelper.EnsurePath(path, nameof(path));
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Word);
            FileHelper.EnsureCanWriteOutput(path, _settings.AllowOverwrite);
            using (var stream = File.Create(path)) { CreateWord(stream, blocks); }
        }

        public Task CreateWordAsync(string path, System.Collections.Generic.IList<Models.WordBlock> blocks, CancellationToken cancellationToken = default)
            => Task.Run(() => { cancellationToken.ThrowIfCancellationRequested(); CreateWord(path, blocks); }, cancellationToken);

        public Task CreateWordAsync(Stream output, System.Collections.Generic.IList<Models.WordBlock> blocks, CancellationToken cancellationToken = default)
            => Task.Run(() => { cancellationToken.ThrowIfCancellationRequested(); CreateWord(output, blocks); }, cancellationToken);

        public Task<byte[]> CreateWordAsync(System.Collections.Generic.IList<Models.WordBlock> blocks, CancellationToken cancellationToken = default)
            => Task.Run(() => { cancellationToken.ThrowIfCancellationRequested(); return CreateWord(blocks); }, cancellationToken);

        private static Paragraph CreateStyledParagraph(Models.WordBlock block)
        {
            var para = new Paragraph();
            var pPr  = new ParagraphProperties();

            if (block.Type == Models.WordBlockType.Heading)
            {
                pPr.Append(new ParagraphStyleId { Val = $"Heading{Math.Max(1, Math.Min(3, block.Level))}" });
            }

            if (!string.IsNullOrEmpty(block.Alignment))
            {
                JustificationValues jv;
                switch (block.Alignment.ToLowerInvariant())
                {
                    case "center":  jv = JustificationValues.Center; break;
                    case "right":   jv = JustificationValues.Right;  break;
                    case "justify": jv = JustificationValues.Both;   break;
                    default:        jv = JustificationValues.Left;   break;
                }
                pPr.Append(new Justification { Val = jv });
            }

            if (!string.IsNullOrEmpty(block.BackgroundColor))
            {
                pPr.Append(new Shading
                {
                    Val   = ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill  = block.BackgroundColor.TrimStart('#')
                });
            }

            para.Append(pPr);

            var run = new Run();
            var rPr = new RunProperties();

            bool bold = block.FontBold || block.Type == Models.WordBlockType.Heading;
            if (bold)              { rPr.Append(new Bold()); }
            if (block.FontItalic)  { rPr.Append(new Italic()); }
            if (block.Underline)   { rPr.Append(new Underline { Val = UnderlineValues.Single }); }
            if (block.FontSize.HasValue)
            {
                string halfPt = ((int)(block.FontSize.Value * 2)).ToString();
                rPr.Append(new FontSize { Val = halfPt });
            }
            if (!string.IsNullOrEmpty(block.FontColor))
            {
                rPr.Append(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = block.FontColor.TrimStart('#') });
            }

            run.Append(rPr);
            run.Append(new Text(block.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
            para.Append(run);
            return para;
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
            const double margin = 36;

            using (var pdf = new PdfDocument())
            {
                var page = pdf.AddPage();
                page.Size = PdfSharpCore.PageSize.A4;
                XGraphics gfx = XGraphics.FromPdfPage(page);
                double yPos = margin;

                try
                {
                    foreach (var block in blocks)
                    {
                        double fontSize  = block.Type == TextBlockType.Heading ? (block.Level == 1 ? 18 : block.Level == 2 ? 16 : 14) : 11;
                        XFontStyle style = block.Type == TextBlockType.Heading ? XFontStyle.Bold : XFontStyle.Regular;
                        double lineH     = fontSize + 6;
                        double spaceAfter = block.Type == TextBlockType.Heading ? 8 : 10;

                        if (yPos + lineH > page.Height - margin)
                        {
                            gfx.Dispose();
                            page      = pdf.AddPage();
                            page.Size = PdfSharpCore.PageSize.A4;
                            gfx       = XGraphics.FromPdfPage(page);
                            yPos      = margin;
                        }

                        var font = new XFont("Arial", fontSize, style);
                        gfx.DrawString(block.Text ?? string.Empty, font, XBrushes.Black,
                            new XRect(margin, yPos, page.Width - margin * 2, lineH),
                            XStringFormats.TopLeft);
                        yPos += lineH + spaceAfter;
                    }
                }
                finally
                {
                    gfx.Dispose();
                }

                using (var ms = new MemoryStream())
                {
                    pdf.Save(ms);
                    ms.Position = 0;
                    if (output.CanSeek) { output.SetLength(0); output.Position = 0; }
                    ms.CopyTo(output);
                    if (output.CanSeek) { output.Position = 0; }
                }
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
