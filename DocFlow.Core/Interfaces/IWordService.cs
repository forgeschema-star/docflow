using System.IO;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DocFlow.Core.Models;

namespace DocFlow.Core.Interfaces
{
    public interface IWordService
    {
        void CreateWord(string path, string content);

        void CreateWord(Stream output, string content);

        byte[] CreateWord(string content);

        Task CreateWordAsync(string path, string content, CancellationToken cancellationToken = default);

        Task CreateWordAsync(Stream output, string content, CancellationToken cancellationToken = default);

        Task<byte[]> CreateWordAsync(string content, CancellationToken cancellationToken = default);

        string ReadWord(string path);

        string ReadWord(Stream input);

        string ReadWord(byte[] inputBytes);

        void ReplacePlaceholders(string templatePath, string outputPath, IDictionary<string, string> placeholders);

        void ReplacePlaceholders(Stream templateStream, Stream outputStream, IDictionary<string, string> placeholders);

        byte[] ReplacePlaceholders(byte[] templateBytes, IDictionary<string, string> placeholders);

        Task<string> ReadWordAsync(string path, CancellationToken cancellationToken = default);

        Task<string> ReadWordAsync(Stream input, CancellationToken cancellationToken = default);

        Task<string> ReadWordAsync(byte[] inputBytes, CancellationToken cancellationToken = default);

        Task ReplacePlaceholdersAsync(string templatePath, string outputPath, IDictionary<string, string> placeholders, CancellationToken cancellationToken = default);

        Task ReplacePlaceholdersAsync(Stream templateStream, Stream outputStream, IDictionary<string, string> placeholders, CancellationToken cancellationToken = default);

        Task<byte[]> ReplacePlaceholdersAsync(byte[] templateBytes, IDictionary<string, string> placeholders, CancellationToken cancellationToken = default);

        void ConvertWordToPdf(string inputPath, string outputPath);

        void ConvertWordToPdf(Stream input, Stream output);

        byte[] ConvertWordToPdf(byte[] inputBytes);

        Task ConvertWordToPdfAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default);

        Task ConvertWordToPdfAsync(Stream input, Stream output, CancellationToken cancellationToken = default);

        Task<byte[]> ConvertWordToPdfAsync(byte[] inputBytes, CancellationToken cancellationToken = default);

        // ── Styled-block overloads ───────────────────────────────────────────

        void CreateWord(string path, IList<WordBlock> blocks);

        void CreateWord(Stream output, IList<WordBlock> blocks);

        byte[] CreateWord(IList<WordBlock> blocks);

        Task CreateWordAsync(string path, IList<WordBlock> blocks, CancellationToken cancellationToken = default);

        Task CreateWordAsync(Stream output, IList<WordBlock> blocks, CancellationToken cancellationToken = default);

        Task<byte[]> CreateWordAsync(IList<WordBlock> blocks, CancellationToken cancellationToken = default);
    }
}
