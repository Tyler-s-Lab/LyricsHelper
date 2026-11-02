using System;
using System.IO;

namespace LyricsHelper {
	internal sealed class DocumentLimitException(string message) : Exception(message);

	/// <summary>
	/// A budget shared by all decompressed streams of one document.
	/// </summary>
	internal sealed class ByteBudget(long bytes) {
		public long Remaining { get; private set; } = bytes;

		public void Consume(int count) {
			if (count <= 0)
				return;

			Remaining -= count;
			if (Remaining < 0)
				throw new DocumentLimitException("document exceeds the configured decompression budget");
		}
	}

	/// <summary>
	/// Read-only stream wrapper that enforces both a per-entry limit and a
	/// shared budget. It deliberately allows a one-byte probe at the limit so
	/// an exactly-sized document can still reach a clean EOF.
	/// </summary>
	internal sealed class BoundedReadStream(Stream inner, long limit, ByteBudget budget) : Stream {
		private readonly Stream _inner = inner;
		private readonly long _limit = limit;
		private readonly ByteBudget _budget = budget;
		private long _read;

		public override bool CanRead => true;
		public override bool CanSeek => false;
		public override bool CanWrite => false;
		public override long Length => throw new NotSupportedException();
		public override long Position {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		private int ReadCore(Span<byte> buffer) {
			long remaining = _limit - _read;
			if (remaining < 0)
				throw new DocumentLimitException("document part exceeds the configured size limit");

			int max = (int)Math.Min(buffer.Length, remaining + 1);
			int read = _inner.Read(buffer[..max]);
			if (read > remaining)
				throw new DocumentLimitException("document part exceeds the configured size limit");
			if (read > 0) {
				_read += read;
				_budget.Consume(read);
			}
			return read;
		}

		public override int Read(byte[] buffer, int offset, int count) {
			return ReadCore(buffer.AsSpan(offset, count));
		}

		public override int Read(Span<byte> buffer) {
			return ReadCore(buffer);
		}

		public override void Flush() { }
		public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
		public override void SetLength(long value) => throw new NotSupportedException();
		public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

		protected override void Dispose(bool disposing) {
			if (disposing)
				_inner.Dispose();
			base.Dispose(disposing);
		}
	}
}
