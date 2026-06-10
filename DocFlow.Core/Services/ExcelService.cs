using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocFlow.Core.Helpers;
using DocFlow.Core.Interfaces;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

namespace DocFlow.Core.Services
{
    public sealed class ExcelService : IExcelService
    {
        private readonly ILogger _logger;
        private readonly Models.DocFlowSettings _settings;

public ExcelService(ILogger logger = null, Models.DocFlowSettings settings = null)
        {
            _logger = logger ?? new NullLogger();
            _settings = settings ?? Models.DocFlowSettings.CreateDefault();
        }

        public void CreateExcel(string path, List<Dictionary<string, string>> data)
        {
            FileHelper.EnsurePath(path, nameof(path));
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Excel);
            FileHelper.EnsureCanWriteOutput(path, _settings.AllowOverwrite);
            LoggingHelper.LogStart(_logger, _settings, "CreateExcel", path);

            using (var stream = File.Create(path))
            {
                CreateExcel(stream, data);
            }

            LoggingHelper.LogEnd(_logger, _settings, "CreateExcel", path);
        }

        public void CreateExcel(Stream output, List<Dictionary<string, string>> data)
        {
            StreamHelper.EnsureWritable(output, nameof(output));
            ValidateData(data);

            try
            {
                if (output.CanSeek)
                {
                    output.SetLength(0);
                    output.Position = 0;
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");
                    var headers = ResolveHeaders(data);

                    for (var column = 0; column < headers.Count; column++)
                    {
                        worksheet.Cell(1, column + 1).Value = headers[column];
                    }

                    var headerRange = worksheet.Range(1, 1, 1, Math.Max(1, headers.Count));
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    for (var rowIndex = 0; rowIndex < data.Count; rowIndex++)
                    {
                        for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
                        {
                            string value;
                            data[rowIndex].TryGetValue(headers[columnIndex], out value);
                            worksheet.Cell(rowIndex + 2, columnIndex + 1).Value = value ?? string.Empty;
                        }
                    }

                    if (headers.Count > 0)
                    {
                        worksheet.Range(1, 1, Math.Max(1, data.Count + 1), headers.Count)
                            .CreateTable("DocFlowTable")
                            .Theme = XLTableTheme.TableStyleMedium2;
                    }

                    worksheet.Columns().AdjustToContents();
                    workbook.SaveAs(output);
                }

                if (output.CanSeek)
                {
                    output.Position = 0;
                }
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to create Excel document.", exception);
                throw new InvalidOperationException("Excel document creation failed.", exception);
            }
        }

        public byte[] CreateExcel(List<Dictionary<string, string>> data)
        {
            using (var stream = new MemoryStream())
            {
                CreateExcel(stream, data);
                return stream.ToArray();
            }
        }

        public Task CreateExcelAsync(string path, List<Dictionary<string, string>> data, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                CreateExcel(path, data);
            }, cancellationToken);
        }

        public Task CreateExcelAsync(Stream output, List<Dictionary<string, string>> data, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                CreateExcel(output, data);
            }, cancellationToken);
        }

        public Task<byte[]> CreateExcelAsync(List<Dictionary<string, string>> data, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return CreateExcel(data);
            }, cancellationToken);
        }

        public List<Dictionary<string, string>> ReadExcel(string path)
        {
            FileHelper.EnsureInputFileExists(path);
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Excel);
            FileHelper.EnsureFileSize(path, _settings.MaxFileSizeBytes);
            LoggingHelper.LogStart(_logger, _settings, "ReadExcel", path);

            using (var stream = File.OpenRead(path))
            {
                var result = ReadExcel(stream);
                LoggingHelper.LogEnd(_logger, _settings, "ReadExcel", path);
                return result;
            }
        }

        public List<Dictionary<string, string>> ReadExcel(Stream input)
        {
            StreamHelper.EnsureReadable(input, nameof(input));

            try
            {
                using (var seekable = StreamHelper.EnsureSeekable(input))
                using (var workbook = new XLWorkbook(seekable))
                {
                    var worksheet = workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null || worksheet.LastRowUsed() == null || worksheet.LastColumnUsed() == null)
                    {
                        return new List<Dictionary<string, string>>();
                    }

                    var headers = new List<string>();
                    var lastColumn = worksheet.LastColumnUsed().ColumnNumber();
                    for (var column = 1; column <= lastColumn; column++)
                    {
                        headers.Add(worksheet.Cell(1, column).GetValue<string>().Trim());
                    }

                    var lastRow = worksheet.LastRowUsed().RowNumber();
                    var result = new List<Dictionary<string, string>>();
                    for (var row = 2; row <= lastRow; row++)
                    {
                        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        for (var column = 1; column <= lastColumn; column++)
                        {
                            values[headers[column - 1]] = worksheet.Cell(row, column).GetValue<string>();
                        }

                        if (values.Values.Any(value => !string.IsNullOrWhiteSpace(value)))
                        {
                            result.Add(values);
                        }
                    }

                    return result;
                }
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to read Excel document.", exception);
                throw new InvalidOperationException("Excel document read failed.", exception);
            }
        }

        public List<Dictionary<string, string>> ReadExcel(byte[] inputBytes)
        {
            using (var stream = StreamHelper.ToMemoryStream(inputBytes, nameof(inputBytes)))
            {
                return ReadExcel(stream);
            }
        }

        public void ReplacePlaceholders(string templatePath, string outputPath, IDictionary<string, string> placeholders)
        {
            FileHelper.EnsureInputFileExists(templatePath);
            FileHelper.EnsureDocumentType(templatePath, Models.DocumentType.Excel);
            FileHelper.EnsureDocumentType(outputPath, Models.DocumentType.Excel);
            FileHelper.EnsureFileSize(templatePath, _settings.MaxFileSizeBytes);
            FileHelper.EnsureCanWriteOutput(outputPath, _settings.AllowOverwrite);
            LoggingHelper.LogStart(_logger, _settings, "ReplaceExcelPlaceholders", templatePath + " => " + outputPath);

            using (var input = File.OpenRead(templatePath))
            using (var output = File.Create(outputPath))
            {
                ReplacePlaceholders(input, output, placeholders);
            }

            LoggingHelper.LogEnd(_logger, _settings, "ReplaceExcelPlaceholders", outputPath);
        }

        public void ReplacePlaceholders(Stream templateStream, Stream outputStream, IDictionary<string, string> placeholders)
        {
            StreamHelper.EnsureReadable(templateStream, nameof(templateStream));
            StreamHelper.EnsureWritable(outputStream, nameof(outputStream));
            PlaceholderHelper.ValidatePlaceholders(placeholders, nameof(placeholders));

            try
            {
                using (var seekable = StreamHelper.EnsureSeekable(templateStream))
                using (var workbook = new XLWorkbook(seekable))
                {
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        foreach (var cell in worksheet.CellsUsed())
                        {
                            if (!cell.HasFormula)
                            {
                                cell.Value = PlaceholderHelper.ReplacePlaceholders(cell.GetValue<string>(), placeholders);
                            }
                        }
                    }

                    if (outputStream.CanSeek)
                    {
                        outputStream.SetLength(0);
                        outputStream.Position = 0;
                    }

                    workbook.SaveAs(outputStream);

                    if (outputStream.CanSeek)
                    {
                        outputStream.Position = 0;
                    }
                }
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to replace placeholders in Excel document.", exception);
                throw new InvalidOperationException("Excel placeholder replacement failed.", exception);
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

        public Task<List<Dictionary<string, string>>> ReadExcelAsync(string path, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReadExcel(path);
            }, cancellationToken);
        }

        public Task<List<Dictionary<string, string>>> ReadExcelAsync(Stream input, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReadExcel(input);
            }, cancellationToken);
        }

        public Task<List<Dictionary<string, string>>> ReadExcelAsync(byte[] inputBytes, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ReadExcel(inputBytes);
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

        public void ConvertExcelToPdf(string inputPath, string outputPath)
        {
            FileHelper.EnsureInputFileExists(inputPath);
            FileHelper.EnsureDocumentType(inputPath, Models.DocumentType.Excel);
            FileHelper.EnsureDocumentType(outputPath, Models.DocumentType.Pdf);
            FileHelper.EnsureFileSize(inputPath, _settings.MaxFileSizeBytes);
            FileHelper.EnsureCanWriteOutput(outputPath, _settings.AllowOverwrite);
            LoggingHelper.LogStart(_logger, _settings, "ConvertExcelToPdf", inputPath + " => " + outputPath);

            using (var input = File.OpenRead(inputPath))
            using (var output = File.Create(outputPath))
            {
                ConvertExcelToPdf(input, output);
            }

            LoggingHelper.LogEnd(_logger, _settings, "ConvertExcelToPdf", outputPath);
        }

        public void ConvertExcelToPdf(Stream input, Stream output)
        {
            StreamHelper.EnsureReadable(input, nameof(input));
            StreamHelper.EnsureWritable(output, nameof(output));

            try
            {
                RenderTableToPdf(ReadExcel(input), output);
            }
            catch (Exception exception)
            {
                _logger.LogError("Failed to convert Excel document to PDF.", exception);
                throw new InvalidOperationException("Excel to PDF conversion failed.", exception);
            }
        }

        public byte[] ConvertExcelToPdf(byte[] inputBytes)
        {
            using (var input = StreamHelper.ToMemoryStream(inputBytes, nameof(inputBytes)))
            using (var output = new MemoryStream())
            {
                ConvertExcelToPdf(input, output);
                return output.ToArray();
            }
        }

        public Task ConvertExcelToPdfAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ConvertExcelToPdf(inputPath, outputPath);
            }, cancellationToken);
        }

        public Task ConvertExcelToPdfAsync(Stream input, Stream output, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                ConvertExcelToPdf(input, output);
            }, cancellationToken);
        }

        public Task<byte[]> ConvertExcelToPdfAsync(byte[] inputBytes, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ConvertExcelToPdf(inputBytes);
            }, cancellationToken);
        }

        private static List<string> ResolveHeaders(List<Dictionary<string, string>> data)
        {
            var headers = new List<string>();
            foreach (var row in data)
            {
                foreach (var key in row.Keys)
                {
                    if (!headers.Contains(key))
                    {
                        headers.Add(key);
                    }
                }
            }

            return headers;
        }

        private static void ValidateData(List<Dictionary<string, string>> data)
        {
            if (data == null)
            {
                throw new ArgumentNullException(nameof(data));
            }

            if (data.Count == 0)
            {
                throw new ArgumentException("Excel data must contain at least one row.", nameof(data));
            }
        }

        private static void RenderTableToPdf(List<Dictionary<string, string>> data, Stream output)
        {
            var headers = ResolveHeaders(data);

            const double margin    = 24;
            const double rowHeight = 18;
            const double titleGap  = 30;

            using (var pdf = new PdfDocument())
            {
                var page = pdf.AddPage();
                page.Size        = PdfSharpCore.PageSize.A4;
                page.Orientation = PdfSharpCore.PageOrientation.Landscape;

                double pageWidth  = page.Width;
                double pageHeight = page.Height;
                double usableW    = pageWidth - margin * 2;
                int    colCount   = Math.Max(1, headers.Count);
                double colWidth   = usableW / colCount;

                XGraphics gfx  = XGraphics.FromPdfPage(page);
                double    yPos = margin;

                try
                {
                    // Title
                    var titleFont = new XFont("Arial", 16, XFontStyle.Bold);
                    gfx.DrawString("Excel Export", titleFont, XBrushes.Black,
                        new XRect(margin, yPos, usableW, titleGap), XStringFormats.TopLeft);
                    yPos += titleGap;

                    // Header row
                    var headerFont = new XFont("Arial", 10, XFontStyle.Bold);
                    var headerBg   = new XSolidBrush(XColor.FromArgb(0xDD, 0xDD, 0xDD));
                    var borderPen  = new XPen(XColor.FromArgb(0xBB, 0xBB, 0xBB), 0.5);

                    for (int i = 0; i < headers.Count; i++)
                    {
                        double x = margin + i * colWidth;
                        gfx.DrawRectangle(borderPen, headerBg, x, yPos, colWidth, rowHeight);
                        gfx.DrawString(headers[i], headerFont, XBrushes.Black,
                            new XRect(x + 2, yPos + 2, colWidth - 4, rowHeight - 4), XStringFormats.TopLeft);
                    }
                    yPos += rowHeight;

                    // Data rows
                    var bodyFont = new XFont("Arial", 9, XFontStyle.Regular);
                    foreach (var row in data)
                    {
                        if (yPos + rowHeight > pageHeight - margin)
                        {
                            gfx.Dispose();
                            page             = pdf.AddPage();
                            page.Size        = PdfSharpCore.PageSize.A4;
                            page.Orientation = PdfSharpCore.PageOrientation.Landscape;
                            gfx              = XGraphics.FromPdfPage(page);
                            yPos             = margin;
                        }

                        for (int i = 0; i < headers.Count; i++)
                        {
                            string value;
                            row.TryGetValue(headers[i], out value);
                            double x = margin + i * colWidth;
                            gfx.DrawRectangle(borderPen, XBrushes.White, x, yPos, colWidth, rowHeight);
                            gfx.DrawString(value ?? string.Empty, bodyFont, XBrushes.Black,
                                new XRect(x + 2, yPos + 2, colWidth - 4, rowHeight - 4), XStringFormats.TopLeft);
                        }
                        yPos += rowHeight;
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

        // ── Rich-cell overloads ──────────────────────────────────────────────

        public void CreateExcel(Stream output, System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows)
        {
            StreamHelper.EnsureWritable(output, nameof(output));
            if (rows == null) { throw new ArgumentNullException(nameof(rows)); }

            try
            {
                if (output.CanSeek) { output.SetLength(0); output.Position = 0; }

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Sheet1");
                    for (int r = 0; r < rows.Count; r++)
                    {
                        var row = rows[r];
                        if (row == null) { continue; }
                        for (int c = 0; c < row.Count; c++)
                        {
                            var def = row[c];
                            if (def == null) { continue; }
                            var cell = ws.Cell(r + 1, c + 1);
                            ApplyCellData(cell, def);
                            ApplyCellStyle(cell, def);
                        }
                    }
                    ws.Columns().AdjustToContents();
                    workbook.SaveAs(output);
                }

                if (output.CanSeek) { output.Position = 0; }
            }
            catch (Exception ex)
            {
                _logger.LogError("Failed to create rich Excel document.", ex);
                throw new InvalidOperationException("Excel document creation failed.", ex);
            }
        }

        public byte[] CreateExcel(System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows)
        {
            using (var ms = new MemoryStream()) { CreateExcel(ms, rows); return ms.ToArray(); }
        }

        public void CreateExcel(string path, System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows)
        {
            FileHelper.EnsurePath(path, nameof(path));
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Excel);
            FileHelper.EnsureCanWriteOutput(path, _settings.AllowOverwrite);
            using (var stream = File.Create(path)) { CreateExcel(stream, rows); }
        }

        public Task CreateExcelAsync(string path, System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows, CancellationToken cancellationToken = default)
            => Task.Run(() => { cancellationToken.ThrowIfCancellationRequested(); CreateExcel(path, rows); }, cancellationToken);

        public Task CreateExcelAsync(Stream output, System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows, CancellationToken cancellationToken = default)
            => Task.Run(() => { cancellationToken.ThrowIfCancellationRequested(); CreateExcel(output, rows); }, cancellationToken);

        public Task<byte[]> CreateExcelAsync(System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows, CancellationToken cancellationToken = default)
            => Task.Run(() => { cancellationToken.ThrowIfCancellationRequested(); return CreateExcel(rows); }, cancellationToken);

        // ── Chart overloads ──────────────────────────────────────────────────

        public void CreateExcelWithChart(Stream output, System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows, Models.ChartDefinition chart)
        {
            StreamHelper.EnsureWritable(output, nameof(output));
            if (rows == null) { throw new ArgumentNullException(nameof(rows)); }
            if (chart == null) { throw new ArgumentNullException(nameof(chart)); }

            using (var ms = new MemoryStream())
            {
                CreateExcel(ms, rows);
                Helpers.ExcelChartHelper.InjectChart(ms, chart, rows.Count, "Sheet1");
                ms.Position = 0;
                if (output.CanSeek) { output.SetLength(0); output.Position = 0; }
                ms.CopyTo(output);
                if (output.CanSeek) { output.Position = 0; }
            }
        }

        public byte[] CreateExcelWithChart(System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows, Models.ChartDefinition chart)
        {
            using (var ms = new MemoryStream()) { CreateExcelWithChart(ms, rows, chart); return ms.ToArray(); }
        }

        public void CreateExcelWithChart(string path, System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows, Models.ChartDefinition chart)
        {
            FileHelper.EnsurePath(path, nameof(path));
            FileHelper.EnsureDocumentType(path, Models.DocumentType.Excel);
            FileHelper.EnsureCanWriteOutput(path, _settings.AllowOverwrite);
            using (var stream = File.Create(path)) { CreateExcelWithChart(stream, rows, chart); }
        }

        public Task CreateExcelWithChartAsync(string path, System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows, Models.ChartDefinition chart, CancellationToken cancellationToken = default)
            => Task.Run(() => { cancellationToken.ThrowIfCancellationRequested(); CreateExcelWithChart(path, rows, chart); }, cancellationToken);

        public Task CreateExcelWithChartAsync(Stream output, System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows, Models.ChartDefinition chart, CancellationToken cancellationToken = default)
            => Task.Run(() => { cancellationToken.ThrowIfCancellationRequested(); CreateExcelWithChart(output, rows, chart); }, cancellationToken);

        public Task<byte[]> CreateExcelWithChartAsync(System.Collections.Generic.IList<System.Collections.Generic.IList<Models.ExcelCell>> rows, Models.ChartDefinition chart, CancellationToken cancellationToken = default)
            => Task.Run(() => { cancellationToken.ThrowIfCancellationRequested(); return CreateExcelWithChart(rows, chart); }, cancellationToken);

        private static void ApplyCellData(IXLCell cell, Models.ExcelCell def)
        {
            if (!string.IsNullOrEmpty(def.Formula))
                cell.FormulaA1 = def.Formula;
            else
                cell.Value = def.Value ?? string.Empty;
        }

        private static void ApplyCellStyle(IXLCell cell, Models.ExcelCell def)
        {
            if (def.FontBold)    { cell.Style.Font.Bold = true; }
            if (def.FontItalic)  { cell.Style.Font.Italic = true; }
            if (def.FontSize.HasValue) { cell.Style.Font.FontSize = def.FontSize.Value; }
            if (!string.IsNullOrEmpty(def.FontColor))       { cell.Style.Font.FontColor = XLColor.FromHtml(def.FontColor); }
            if (!string.IsNullOrEmpty(def.BackgroundColor))
            {
                cell.Style.Fill.PatternType = XLFillPatternValues.Solid;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml(def.BackgroundColor);
            }
            if (!string.IsNullOrEmpty(def.BorderColor))
            {
                var bc = XLColor.FromHtml(def.BorderColor);
                cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                cell.Style.Border.SetOutsideBorderColor(bc);
            }
            if (def.Alignment != Models.ExcelHorizontalAlignment.Default)
            {
                cell.Style.Alignment.Horizontal = def.Alignment == Models.ExcelHorizontalAlignment.Center ? XLAlignmentHorizontalValues.Center
                    : def.Alignment == Models.ExcelHorizontalAlignment.Right ? XLAlignmentHorizontalValues.Right
                    : XLAlignmentHorizontalValues.Left;
            }
        }
    }
}
