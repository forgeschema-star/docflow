using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace DocFlow.Core.Helpers
{
    public static class StreamHelper
    {
        public static void EnsureReadable(Stream stream, string parameterName)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            if (!stream.CanRead)
            {
                throw new InvalidOperationException("The provided stream must be readable.");
            }
        }

        public static void EnsureWritable(Stream stream, string parameterName)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            if (!stream.CanWrite)
            {
                throw new InvalidOperationException("The provided stream must be writable.");
            }
        }

        public static MemoryStream EnsureSeekable(Stream stream)
        {
            EnsureReadable(stream, nameof(stream));

            if (stream is MemoryStream memoryStream && stream.CanSeek)
            {
                memoryStream.Position = 0;
                return new MemoryStream(memoryStream.ToArray());
            }

            if (stream.CanSeek)
            {
                stream.Position = 0;
            }

            var copy = new MemoryStream();
            stream.CopyTo(copy);
            copy.Position = 0;
            return copy;
        }

        public static byte[] ReadAllBytes(Stream stream)
        {
            using (var copy = EnsureSeekable(stream))
            {
                return copy.ToArray();
            }
        }

        public static async Task<byte[]> ReadAllBytesAsync(Stream stream, CancellationToken cancellationToken = default)
        {
            EnsureReadable(stream, nameof(stream));

            var copy = new MemoryStream();
            if (stream.CanSeek)
            {
                stream.Position = 0;
            }

            await stream.CopyToAsync(copy, 81920, cancellationToken).ConfigureAwait(false);
            return copy.ToArray();
        }

        public static MemoryStream ToMemoryStream(byte[] inputBytes, string parameterName)
        {
            if (inputBytes == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            if (inputBytes.Length == 0)
            {
                throw new ArgumentException("The provided byte array cannot be empty.", parameterName);
            }

            return new MemoryStream(inputBytes, writable: false);
        }
    }
}
