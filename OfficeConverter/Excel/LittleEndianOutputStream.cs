using System;
using System.IO;
using OfficeConverter.Excel.Interfaces;

namespace OfficeConverter.Excel
{
    /// <summary>
    ///     Wraps an <see cref="T:System.IO.Stream" /> providing <see cref="T:NPOI.Util.ILittleEndianOutput" />
    /// </summary>
    /// <remarks>@author Josh Micich</remarks>
    internal class LittleEndianOutputStream : ILittleEndianOutput, IDisposable
    {
        private Stream _output;

        public LittleEndianOutputStream(Stream out1)
        {
            _output = out1;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) return;
            if (null == _output) return;
            _output.Dispose();
            _output = null;
        }

        public void WriteByte(int v)
        {
            _output.WriteByte((byte) v);
        }

        public void WriteDouble(double v)
        {
            WriteLong(BitConverter.DoubleToInt64Bits(v));
        }

        public void WriteInt(int v)
        {
            var b3 = (v >> 24) & 0xFF;
            var b2 = (v >> 16) & 0xFF;
            var b1 = (v >> 8) & 0xFF;
            var b0 = (v >> 0) & 0xFF;
            _output.WriteByte((byte) b0);
            _output.WriteByte((byte) b1);
            _output.WriteByte((byte) b2);
            _output.WriteByte((byte) b3);
        }

        public void WriteLong(long v)
        {
            WriteInt((int) (v >> 0));
            WriteInt((int) (v >> 32));
        }

        public void WriteShort(int v)
        {
            var b1 = (v >> 8) & 0xFF;
            var b0 = (v >> 0) & 0xFF;
            _output.WriteByte((byte) b0);
            _output.WriteByte((byte) b1);
        }

        public void Write(byte[] b)
        {
            // suppress IOException for interface method
            _output.Write(b, 0, b.Length);
        }

        public void Write(byte[] b, int off, int len)
        {
            // suppress IOException for interface method
            _output.Write(b, off, len);
        }

        public void Flush()
        {
            _output.Flush();
        }
    }
}