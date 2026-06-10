using System;
using System.Collections.Generic;
using System.IO;
using DocFlow.Core.Models;

namespace DocFlow.Core.Helpers
{
    public static class FileHelper
    {
        private static readonly IDictionary<DocumentType, string[]> AllowedExtensions =
            new Dictionary<DocumentType, string[]>
            {
                { DocumentType.Word, new[] { ".docx" } },
                { DocumentType.Pdf, new[] { ".pdf" } },
                { DocumentType.Excel, new[] { ".xlsx" } },
                { DocumentType.Csv, new[] { ".csv" } },
                { DocumentType.Html, new[] { ".html", ".htm" } },
                { DocumentType.Image, new[] { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff" } }
            };

        public static void EnsurePath(string path, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Path cannot be null or whitespace.", parameterName);
            }
        }

        public static void EnsureInputFileExists(string path)
        {
            EnsurePath(path, nameof(path));

            if (!File.Exists(path))
            {
                throw new FileNotFoundException("The specified input file was not found.", path);
            }
        }

        public static void EnsureFileSize(string path, long maxFileSizeBytes)
        {
            if (maxFileSizeBytes <= 0)
            {
                return;
            }

            EnsureInputFileExists(path);
            var fileInfo = new FileInfo(path);
            if (fileInfo.Length > maxFileSizeBytes)
            {
                throw new InvalidOperationException(
                    string.Format("The file '{0}' exceeds the configured size limit.", path));
            }
        }

        public static void EnsureDocumentType(string path, DocumentType documentType)
        {
            EnsurePath(path, nameof(path));

            var extension = Path.GetExtension(path);
            if (string.IsNullOrWhiteSpace(extension))
            {
                throw new InvalidOperationException("The file extension could not be determined.");
            }

            if (!AllowedExtensions.ContainsKey(documentType))
            {
                throw new InvalidOperationException("Unsupported document type.");
            }

            var allowed = AllowedExtensions[documentType];
            foreach (var candidate in allowed)
            {
                if (string.Equals(extension, candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            throw new InvalidOperationException(
                string.Format("The file '{0}' is not a valid {1} document.", path, documentType));
        }

        public static void EnsureOutputDirectory(string path)
        {
            EnsurePath(path, nameof(path));

            var directory = Path.GetDirectoryName(Path.GetFullPath(path));
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidOperationException("The output directory could not be determined.");
            }

            Directory.CreateDirectory(directory);
        }

        public static void EnsureCanWriteOutput(string path, bool allowOverwrite)
        {
            EnsurePath(path, nameof(path));
            EnsureOutputDirectory(path);

            if (!allowOverwrite && File.Exists(path))
            {
                throw new InvalidOperationException("The output file already exists and overwrite is disabled.");
            }
        }

        public static DocumentType ResolveDocumentType(string path)
        {
            EnsurePath(path, nameof(path));

            var extension = Path.GetExtension(path);
            if (string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase))
            {
                return DocumentType.Word;
            }

            if (string.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                return DocumentType.Pdf;
            }

            if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return DocumentType.Excel;
            }

            if (string.Equals(extension, ".csv", StringComparison.OrdinalIgnoreCase))
            {
                return DocumentType.Csv;
            }

            if (string.Equals(extension, ".html", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".htm", StringComparison.OrdinalIgnoreCase))
            {
                return DocumentType.Html;
            }

            if (string.Equals(extension, ".png", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".jpg", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".jpeg", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".bmp", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".gif", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".tif", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".tiff", StringComparison.OrdinalIgnoreCase))
            {
                return DocumentType.Image;
            }

            return DocumentType.Unknown;
        }
    }
}
