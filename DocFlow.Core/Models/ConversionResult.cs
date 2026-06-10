namespace DocFlow.Core.Models
{
    public sealed class ConversionResult
    {
        public bool Success { get; set; }

        public ConversionErrorCode ErrorCode { get; set; }

        public DocumentType SourceType { get; set; }

        public DocumentType TargetType { get; set; }

        public string Message { get; set; }

        public string OutputPath { get; set; }

        public byte[] OutputBytes { get; set; }
    }
}
