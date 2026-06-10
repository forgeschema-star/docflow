using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace DocFlow.Core.Interfaces
{
    public interface IExcelService
    {
        void CreateExcel(string path, List<Dictionary<string, string>> data);

        void CreateExcel(Stream output, List<Dictionary<string, string>> data);

        byte[] CreateExcel(List<Dictionary<string, string>> data);

        Task CreateExcelAsync(string path, List<Dictionary<string, string>> data, CancellationToken cancellationToken = default);

        Task CreateExcelAsync(Stream output, List<Dictionary<string, string>> data, CancellationToken cancellationToken = default);

        Task<byte[]> CreateExcelAsync(List<Dictionary<string, string>> data, CancellationToken cancellationToken = default);

        List<Dictionary<string, string>> ReadExcel(string path);

        List<Dictionary<string, string>> ReadExcel(Stream input);

        List<Dictionary<string, string>> ReadExcel(byte[] inputBytes);

        void ReplacePlaceholders(string templatePath, string outputPath, IDictionary<string, string> placeholders);

        void ReplacePlaceholders(Stream templateStream, Stream outputStream, IDictionary<string, string> placeholders);

        byte[] ReplacePlaceholders(byte[] templateBytes, IDictionary<string, string> placeholders);

        Task<List<Dictionary<string, string>>> ReadExcelAsync(string path, CancellationToken cancellationToken = default);

        Task<List<Dictionary<string, string>>> ReadExcelAsync(Stream input, CancellationToken cancellationToken = default);

        Task<List<Dictionary<string, string>>> ReadExcelAsync(byte[] inputBytes, CancellationToken cancellationToken = default);

        Task ReplacePlaceholdersAsync(string templatePath, string outputPath, IDictionary<string, string> placeholders, CancellationToken cancellationToken = default);

        Task ReplacePlaceholdersAsync(Stream templateStream, Stream outputStream, IDictionary<string, string> placeholders, CancellationToken cancellationToken = default);

        Task<byte[]> ReplacePlaceholdersAsync(byte[] templateBytes, IDictionary<string, string> placeholders, CancellationToken cancellationToken = default);

        void ConvertExcelToPdf(string inputPath, string outputPath);

        void ConvertExcelToPdf(Stream input, Stream output);

        byte[] ConvertExcelToPdf(byte[] inputBytes);

        Task ConvertExcelToPdfAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task ConvertExcelToPdfAsync(Stream input, Stream output, CancellationToken cancellationToken = default);

        Task<byte[]> ConvertExcelToPdfAsync(byte[] inputBytes, CancellationToken cancellationToken = default);
    }
}
