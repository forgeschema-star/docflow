using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocFlow.Core.Helpers;
using DocFlow.Core.Interfaces;
using DocFlow.Core.Models;

namespace DocFlow.Core.Services
{
    public sealed class ConversionService : IConversionService
    {
        private readonly IWordService _wordService;
        private readonly IExcelService _excelService;
        private readonly IPdfService _pdfService;
        private readonly ICsvService _csvService;
        private readonly IHtmlService _htmlService;
        private readonly IImageService _imageService;
        private readonly ILogger _logger;
        private readonly DocFlowSettings _settings;

        public ConversionService(IWordService wordService, IExcelService excelService, IPdfService pdfService, ICsvService csvService, IHtmlService htmlService, IImageService imageService, ILogger logger = null, DocFlowSettings settings = null)
        {
            _wordService = wordService ?? throw new ArgumentNullException(nameof(wordService));
            _excelService = excelService ?? throw new ArgumentNullException(nameof(excelService));
            _pdfService = pdfService ?? throw new ArgumentNullException(nameof(pdfService));
            _csvService = csvService ?? throw new ArgumentNullException(nameof(csvService));
            _htmlService = htmlService ?? throw new ArgumentNullException(nameof(htmlService));
            _imageService = imageService ?? throw new ArgumentNullException(nameof(imageService));
            _logger = logger ?? new NullLogger();
            _settings = settings ?? DocFlowSettings.CreateDefault();
        }

        public ConversionResult Convert(DocumentType from, DocumentType to, string inputPath, string outputPath)
        {
            try
            {
                FileHelper.EnsureInputFileExists(inputPath);
                FileHelper.EnsureDocumentType(inputPath, from);
                FileHelper.EnsureDocumentType(outputPath, to);
                FileHelper.EnsureFileSize(inputPath, _settings.MaxFileSizeBytes);
                FileHelper.EnsureCanWriteOutput(outputPath, _settings.AllowOverwrite);
                LoggingHelper.LogStart(_logger, _settings, "Convert", from + " => " + to);
                ConvertFileCore(from, to, inputPath, outputPath);
                LoggingHelper.LogEnd(_logger, _settings, "Convert", outputPath);

                return new ConversionResult
                {
                    Success = true,
                    ErrorCode = ConversionErrorCode.None,
                    SourceType = from,
                    TargetType = to,
                    OutputPath = outputPath,
                    Message = "Conversion completed successfully."
                };
            }
            catch (Exception exception)
            {
                _logger.LogError("Document conversion failed.", exception);
                return CreateFailure(from, to, outputPath, exception);
            }
        }

        public ConversionResult Convert(DocumentType from, DocumentType to, Stream input, Stream output)
        {
            try
            {
                StreamHelper.EnsureReadable(input, nameof(input));
                StreamHelper.EnsureWritable(output, nameof(output));
                ConvertCore(from, to, input, output);

                return new ConversionResult
                {
                    Success = true,
                    ErrorCode = ConversionErrorCode.None,
                    SourceType = from,
                    TargetType = to,
                    Message = "Conversion completed successfully."
                };
            }
            catch (Exception exception)
            {
                _logger.LogError("Document conversion failed.", exception);
                return CreateFailure(from, to, null, exception);
            }
        }

        public ConversionResult Convert(DocumentType from, DocumentType to, byte[] inputBytes)
        {
            try
            {
                using (var input = StreamHelper.ToMemoryStream(inputBytes, nameof(inputBytes)))
                using (var output = new MemoryStream())
                {
                    ConvertCore(from, to, input, output);
                    return new ConversionResult
                    {
                        Success = true,
                        ErrorCode = ConversionErrorCode.None,
                        SourceType = from,
                        TargetType = to,
                        OutputBytes = output.ToArray(),
                        Message = "Conversion completed successfully."
                    };
                }
            }
            catch (Exception exception)
            {
                _logger.LogError("Document conversion failed.", exception);
                return CreateFailure(from, to, null, exception);
            }
        }

        public Task<ConversionResult> ConvertAsync(DocumentType from, DocumentType to, string inputPath, string outputPath, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return Convert(from, to, inputPath, outputPath);
            }, cancellationToken);
        }

        public Task<ConversionResult> ConvertAsync(DocumentType from, DocumentType to, Stream input, Stream output, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return Convert(from, to, input, output);
            }, cancellationToken);
        }

        public Task<ConversionResult> ConvertAsync(DocumentType from, DocumentType to, byte[] inputBytes, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return Convert(from, to, inputBytes);
            }, cancellationToken);
        }

        private void ConvertCore(DocumentType from, DocumentType to, Stream input, Stream output)
        {
            if (from == to)
            {
                using (var copy = StreamHelper.EnsureSeekable(input))
                {
                    copy.CopyTo(output);
                }

                if (output.CanSeek)
                {
                    output.Position = 0;
                }

                return;
            }

            if (from == DocumentType.Word && to == DocumentType.Pdf)
            {
                _wordService.ConvertWordToPdf(input, output);
                return;
            }

            if (from == DocumentType.Excel && to == DocumentType.Pdf)
            {
                _excelService.ConvertExcelToPdf(input, output);
                return;
            }

            if (from == DocumentType.Pdf && to == DocumentType.Word)
            {
                _pdfService.ConvertPdfToWord(input, output);
                return;
            }

            if (from == DocumentType.Pdf && to == DocumentType.Excel)
            {
                _pdfService.ConvertPdfToExcel(input, output);
                return;
            }

            var inputExtension = GetExtension(from);
            var outputExtension = GetExtension(to);
            WriteOutputBytes(output, BuildBytesFromFileAction(input, inputExtension, outputExtension, (source, destination) => ConvertFileCore(from, to, source, destination)));
            return;
        }

        private void ConvertFileCore(DocumentType from, DocumentType to, string inputPath, string outputPath)
        {
            if (from == to)
            {
                File.Copy(inputPath, outputPath, true);
                return;
            }

            if (from == DocumentType.Word && to == DocumentType.Pdf)
            {
                _wordService.ConvertWordToPdf(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Excel && to == DocumentType.Pdf)
            {
                _excelService.ConvertExcelToPdf(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Pdf && to == DocumentType.Word)
            {
                _pdfService.ConvertPdfToWord(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Pdf && to == DocumentType.Excel)
            {
                _pdfService.ConvertPdfToExcel(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Csv && to == DocumentType.Excel)
            {
                _csvService.ConvertCsvToExcel(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Excel && to == DocumentType.Csv)
            {
                _csvService.ConvertExcelToCsv(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Csv && to == DocumentType.Pdf)
            {
                _csvService.ConvertCsvToPdf(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Html && to == DocumentType.Word)
            {
                _htmlService.ConvertHtmlToWord(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Html && to == DocumentType.Pdf)
            {
                _htmlService.ConvertHtmlToPdf(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Html && to == DocumentType.Excel)
            {
                _htmlService.ConvertHtmlToExcel(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Image && to == DocumentType.Pdf)
            {
                _imageService.ConvertImageToPdf(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Image && to == DocumentType.Word)
            {
                _imageService.ConvertImageToWord(inputPath, outputPath);
                return;
            }

            if (from == DocumentType.Image && to == DocumentType.Excel)
            {
                _imageService.ConvertImageToExcel(inputPath, outputPath);
                return;
            }

            throw new NotSupportedException(string.Format("Conversion from {0} to {1} is not supported.", from, to));
        }

        private static void WriteOutputBytes(Stream output, byte[] bytes)
        {
            output.Write(bytes, 0, bytes.Length);
            if (output.CanSeek)
            {
                output.Position = 0;
            }
        }

        private byte[] BuildBytesFromFileAction(Stream input, string inputExtension, string outputExtension, Action<string, string> conversion)
        {
            var sourcePath = CreateTempPath(inputExtension);
            var outputPath = CreateTempPath(outputExtension);
            try
            {
                using (var file = File.Create(sourcePath))
                {
                    using (var copy = StreamHelper.EnsureSeekable(input))
                    {
                        copy.CopyTo(file);
                    }
                }

                conversion(sourcePath, outputPath);
                try
                {
                    return File.ReadAllBytes(outputPath);
                }
                finally
                {
                    if (File.Exists(outputPath))
                    {
                        File.Delete(outputPath);
                    }
                }
            }
            finally
            {
                if (File.Exists(sourcePath))
                {
                    File.Delete(sourcePath);
                }
            }
        }

        private static string CreateTempPath(string extension)
        {
            return Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        }

        private static string GetExtension(DocumentType documentType)
        {
            switch (documentType)
            {
                case DocumentType.Word:
                    return ".docx";
                case DocumentType.Pdf:
                    return ".pdf";
                case DocumentType.Excel:
                    return ".xlsx";
                case DocumentType.Csv:
                    return ".csv";
                case DocumentType.Html:
                    return ".html";
                case DocumentType.Image:
                    return ".png";
                default:
                    throw new NotSupportedException("Unsupported document type.");
            }
        }

        private static ConversionResult CreateFailure(DocumentType from, DocumentType to, string outputPath, Exception exception)
        {
            return new ConversionResult
            {
                Success = false,
                ErrorCode = MapErrorCode(exception),
                SourceType = from,
                TargetType = to,
                OutputPath = outputPath,
                Message = exception.Message
            };
        }

        private static ConversionErrorCode MapErrorCode(Exception exception)
        {
            if (exception is FileNotFoundException)
            {
                return ConversionErrorCode.FileNotFound;
            }

            if (exception is NotSupportedException)
            {
                return ConversionErrorCode.UnsupportedConversion;
            }

            if (exception is InvalidOperationException)
            {
                var message = exception.Message ?? string.Empty;
                if (message.IndexOf("size limit", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return ConversionErrorCode.FileTooLarge;
                }

                if (message.IndexOf("already exists", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return ConversionErrorCode.OutputAlreadyExists;
                }

                if (message.IndexOf("valid", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return ConversionErrorCode.InvalidFileType;
                }

                return ConversionErrorCode.ValidationFailed;
            }

            return ConversionErrorCode.ProcessingFailed;
        }
    }
}
