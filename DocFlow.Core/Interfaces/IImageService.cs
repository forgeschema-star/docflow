using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace DocFlow.Core.Interfaces
{
    public interface IImageService
    {
        string ReadImageText(string path);

        void ConvertImageToPdf(string inputPath, string outputPath);

        void ConvertImageToWord(string inputPath, string outputPath);

        void ConvertImageToExcel(string inputPath, string outputPath);

        Task<string> ReadImageTextAsync(string path, CancellationToken cancellationToken = default);

        Task ConvertImageToPdfAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task ConvertImageToWordAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task ConvertImageToExcelAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);
    }
}
