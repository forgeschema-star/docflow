namespace DocFlow.Core.Models
{
    public enum ConversionErrorCode
    {
        None = 0,
        ValidationFailed = 1,
        FileNotFound = 2,
        InvalidFileType = 3,
        FileTooLarge = 4,
        OutputAlreadyExists = 5,
        UnsupportedConversion = 6,
        ProcessingFailed = 7
    }
}
