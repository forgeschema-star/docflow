using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace DocFlow.Core.Interfaces
{
    public interface IHtmlService
    {
        string ReadHtml(string path);

        void ConvertHtmlToWord(string inputPath, string outputPath);

        void ConvertHtmlToPdf(string inputPath, string outputPath);

        void ConvertHtmlToExcel(string inputPath, string outputPath);

        List<Dictionary<string, string>> ExtractTables(string path);

        Task ConvertHtmlToWordAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task ConvertHtmlToPdfAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task ConvertHtmlToExcelAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);
    }
}
