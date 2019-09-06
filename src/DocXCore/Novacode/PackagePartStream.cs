using System.IO;

namespace Novacode
{
	public class PackagePartStream : Stream
	{
		private static readonly object lockObject = new object();

		private readonly Stream stream;

		public override bool CanRead => stream.CanRead;

		public override bool CanSeek => stream.CanSeek;

		public override bool CanWrite => stream.CanWrite;

		public override long Length => stream.Length;

		public override long Position
		{
			get
			{
				return stream.Position;
			}
			set
			{
				stream.Position = value;
			}
		}

		public PackagePartStream(Stream stream)
		{
			this.stream = stream;
		}

		public override long Seek(long offset, SeekOrigin origin)
		{
			return stream.Seek(offset, origin);
		}

		public override void SetLength(long value)
		{
			stream.SetLength(value);
		}

		public override int Read(byte[] buffer, int offset, int count)
		{
			return stream.Read(buffer, offset, count);
		}

		public override void Write(byte[] buffer, int offset, int count)
		{
			lock (lockObject)
			{
				stream.Write(buffer, offset, count);
			}
		}

		public override void Flush()
		{
			lock (lockObject)
			{
				stream.Flush();
			}
		}

		public override void Close()
		{
			stream.Close();
		}

		protected override void Dispose(bool disposing)
		{
			stream.Flush();
			stream.Dispose();
		}
	}
}
