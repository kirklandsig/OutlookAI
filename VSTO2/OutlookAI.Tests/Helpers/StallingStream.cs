using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Tests.Helpers
{
    /// <summary>
    /// Returns <c>prefix</c>, then blocks forever. Like the .NET Framework
    /// connection stream behind HttpClient, a pending read ignores its
    /// CancellationToken; only disposing the stream ends it.
    /// </summary>
    internal sealed class StallingStream : Stream
    {
        private readonly byte[] _prefix;
        private readonly TaskCompletionSource<int> _stall =
            new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _position;

        public StallingStream(byte[] prefix)
        {
            _prefix = prefix ?? new byte[0];
        }

        public bool Disposed { get; private set; }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            if (_position < _prefix.Length)
            {
                var n = Math.Min(count, _prefix.Length - _position);
                Array.Copy(_prefix, _position, buffer, offset, n);
                _position += n;
                return Task.FromResult(n);
            }
            return _stall.Task;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return ReadAsync(buffer, offset, count, CancellationToken.None).GetAwaiter().GetResult();
        }

        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            _stall.TrySetException(new ObjectDisposedException(nameof(StallingStream)));
            base.Dispose(disposing);
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
