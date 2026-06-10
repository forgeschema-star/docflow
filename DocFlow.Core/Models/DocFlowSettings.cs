namespace DocFlow.Core.Models
{
    public sealed class DocFlowSettings
    {
        public string TempDirectory { get; set; }

        public string OcrDataPath { get; set; }

        public long MaxFileSizeBytes { get; set; }

        public bool LoggingEnabled { get; set; }

        public bool AllowOverwrite { get; set; }

        public static DocFlowSettings CreateDefault()
        {
            return new DocFlowSettings
            {
                TempDirectory = string.Empty,
                OcrDataPath = string.Empty,
                MaxFileSizeBytes = 25L * 1024L * 1024L,
                LoggingEnabled = true,
                AllowOverwrite = false
            };
        }
    }
}
