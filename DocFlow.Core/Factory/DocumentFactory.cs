using System.IO;
using DocFlow.Core.Helpers;
using DocFlow.Core.Models;

namespace DocFlow.Core.Factory
{
    public sealed class DocumentFactory
    {
        public DocumentRequest Create(DocumentType from, DocumentType to, string inputPath, string outputPath)
        {
            return new DocumentRequest
            {
                SourceType = from,
                TargetType = to,
                InputPath = inputPath,
                OutputPath = outputPath
            };
        }

        public DocumentRequest Create(DocumentType from, DocumentType to, Stream inputStream, Stream outputStream)
        {
            return new DocumentRequest
            {
                SourceType = from,
                TargetType = to,
                InputStream = inputStream,
                OutputStream = outputStream
            };
        }

        public DocumentRequest Create(DocumentType from, DocumentType to, byte[] inputBytes)
        {
            return new DocumentRequest
            {
                SourceType = from,
                TargetType = to,
                InputBytes = inputBytes
            };
        }

        public DocumentType ResolveDocumentType(string path)
        {
            return FileHelper.ResolveDocumentType(path);
        }
    }
}
