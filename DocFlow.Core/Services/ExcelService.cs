using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocFlow.Core.Helpers;
using DocFlow.Core.Interfaces;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace DocFlow.Core.Services
{
    public sealed class ExcelService : IExcelService
    {
        private readonly ILogger _logger;
        private readonly Models.DocFlowSettings _settings;

        static ExcelService()
        {
            QuestPDF.Settings.License = LicenseType.Community;
        }

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
                            if (!cell.IsFormula)
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
            if (output.CanSeek)
            {
                output.SetLength(0);
                output.Position = 0;
            }

            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Margin(24);
                    page.Size(PageSizes.A4.Landscape());
                    page.DefaultTextStyle(x => x.FontSize(10));
                    page.Header().Text("Excel Export").Bold().FontSize(16);
                    page.Content().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            foreach (var header in headers)
                            {
                                columns.RelativeColumn();
                            }
                        });

                        table.Header(header =>
                        {
                            foreach (var title in headers)
                            {
                                header.Cell().Element(CellStyle).Background(Colors.Grey.Lighten2).Text(title).Bold();
                            }
                        });

                        foreach (var row in data)
                        {
                            foreach (var header in headers)
                            {
                                string value;
                                row.TryGetValue(header, out value);
                                table.Cell().Element(CellStyle).Text(value ?? string.Empty);
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

        private static IContainer CellStyle(IContainer container)
        {
            return container.Border(1).BorderColor(Colors.Grey.Lighten2).Padding(4);
        }
    }
}
