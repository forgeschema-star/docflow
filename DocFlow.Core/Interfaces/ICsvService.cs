using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace DocFlow.Core.Interfaces
{
    public interface ICsvService
    {
        void CreateCsv(string path, List<Dictionary<string, string>> data);

        List<Dictionary<string, string>> ReadCsv(string path);

        void ConvertCsvToExcel(string inputPath, string outputPath);

        void ConvertExcelToCsv(string inputPath, string outputPath);

        void ConvertCsvToPdf(string inputPath, string outputPath);

        Task ConvertCsvToExcelAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task ConvertExcelToCsvAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task ConvertCsvToPdfAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);
    }
}
