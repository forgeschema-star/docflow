using System.IO;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace DocFlow.Core.Interfaces
{
    public interface IPdfService
    {
        string ReadPdf(string path);

        string ReadPdf(Stream input);

        string ReadPdf(byte[] inputBytes);

        IList<string> ExtractImages(string inputPath, string outputDirectory);

        Task<string> ReadPdfAsync(string path, CancellationToken cancellationToken = default);

        Task<string> ReadPdfAsync(Stream input, CancellationToken cancellationToken = default);

        Task<string> ReadPdfAsync(byte[] inputBytes, CancellationToken cancellationToken = default);

        Task<IList<string>> ExtractImagesAsync(string inputPath, string outputDirectory, CancellationToken cancellationToken = default);

        void ConvertPdfToWord(string inputPath, string outputPath);

        void ConvertPdfToWord(Stream input, Stream output);

        byte[] ConvertPdfToWord(byte[] inputBytes);

        Task ConvertPdfToWordAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task ConvertPdfToWordAsync(Stream input, Stream output, CancellationToken cancellationToken = default);

        Task<byte[]> ConvertPdfToWordAsync(byte[] inputBytes, CancellationToken cancellationToken = default);

        void ConvertPdfToExcel(string inputPath, string outputPath);

        void ConvertPdfToExcel(Stream input, Stream output);

        byte[] ConvertPdfToExcel(byte[] inputBytes);

        Task ConvertPdfToExcelAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task ConvertPdfToExcelAsync(Stream input, Stream output, CancellationToken cancellationToken = default);

        Task<byte[]> ConvertPdfToExcelAsync(byte[] inputBytes, CancellationToken cancellationToken = default);
    }
}
