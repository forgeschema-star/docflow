using System.IO;
using System.Threading;
using System.Threading.Tasks;
using DocFlow.Core.Models;

namespace DocFlow.Core.Interfaces
{
    public interface IConversionService
    {
        ConversionResult Convert(DocumentType from, DocumentType to, string inputPath, string outputPath);

        ConversionResult Convert(DocumentType from, DocumentType to, Stream input, Stream output);

        ConversionResult Convert(DocumentType from, DocumentType to, byte[] inputBytes);

        Task<ConversionResult> ConvertAsync(DocumentType from, DocumentType to, string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task<ConversionResult> ConvertAsync(DocumentType from, DocumentType to, Stream input, Stream output, CancellationToken cancellationToken = default);

        Task<ConversionResult> ConvertAsync(DocumentType from, DocumentType to, byte[] inputBytes, CancellationToken cancellationToken = default);
    }
}
