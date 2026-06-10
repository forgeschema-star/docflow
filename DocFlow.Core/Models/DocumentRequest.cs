using System.IO;

namespace DocFlow.Core.Models
{
    public sealed class DocumentRequest
    {
        public DocumentType SourceType { get; set; }

        public DocumentType TargetType { get; set; }

        public string InputPath { get; set; }

        public string OutputPath { get; set; }

        public Stream InputStream { get; set; }

        public Stream OutputStream { get; set; }

        public byte[] InputBytes { get; set; }
    }
}
