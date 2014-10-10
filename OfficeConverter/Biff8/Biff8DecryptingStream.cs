using System;
using System.IO;
using OfficeConverter.Biff8.Interfaces;
using OfficeConverter.Exceptions;

namespace OfficeConverter.Biff8
{
    /// <summary>
    /// Used to return a decrypted Biff8 stream
    /// </summary>
    internal class Biff8DecryptingStream : IBiffHeaderInput, ILittleEndianInput
    {
        #region Fields
        private readonly ILittleEndianInput _littleEndianInput;
        private readonly Biff8RC4 _rc4;
        #endregion

        #region Constructor
        public Biff8DecryptingStream(Stream inputStream, int initialOffSet, Biff8EncryptionKey key)
        {
            _rc4 = new Biff8RC4(initialOffSet, key);
            // ReSharper disable once SuspiciousTypeConversion.Global
            var input = inputStream as ILittleEndianInput;
            _littleEndianInput = input ?? new LittleEndianInputStream(inputStream);
        }
        #endregion

        #region Available
        public int Available()
        {
            return _littleEndianInput.Available();
        }
        #endregion

        #region ReadRecordSid
        /// <summary>
        /// Returns an unsigned short value without decrypting
        /// </summary>
        /// <returns></returns>
        public int ReadRecordSid()
        {
            var sid = _littleEndianInput.ReadUShort();
            _rc4.SkipTwoBytes();
            _rc4.StartRecord(sid);
            return sid;
        }
        #endregion

        #region ReadDataSize
        /// <summary>
        /// Returns an unsigned short value without decrypting
        /// </summary>
        /// <returns></returns>
        public int ReadDataSize()
        {
            var dataSize = _littleEndianInput.ReadUShort();
            _rc4.SkipTwoBytes();
            return dataSize;
        }
        #endregion

        #region ReadDouble
        /// <summary>
        /// Returns a double from the stream
        /// </summary>
        /// <returns></returns>
        public double ReadDouble()
        {
            var valueLongBits = ReadLong();
            var result = BitConverter.Int64BitsToDouble(valueLongBits);
            if (Double.IsNaN(result))
                throw new OCFileIsCorrupt("Did not expect to read NaN");

            return result;
        }
        #endregion

        #region ReadFully
        public void ReadFully(byte[] buffer)
        {
            ReadFully(buffer, 0, buffer.Length);
        }

        /// <summary>
        /// Returns a full array from the stream
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="offset"></param>
        /// <param name="length"></param>
        public void ReadFully(byte[] buffer, int offset, int length)
        {
            _littleEndianInput.ReadFully(buffer, offset, length);
            _rc4.Xor(buffer, offset, length);
        }
        #endregion

        #region ReadUByte
        /// <summary>
        /// Returns an unsigned byte from the stream
        /// </summary>
        /// <returns></returns>
        public int ReadUByte()
        {
            return _rc4.XorByte(_littleEndianInput.ReadUByte());
        }
        #endregion

        #region ReadByte
        /// <summary>
        /// Returns a byte from the stream
        /// </summary>
        /// <returns></returns>
        public int ReadByte()
        {
            return _rc4.XorByte(_littleEndianInput.ReadUByte());
        }
        #endregion

        #region ReadUShort
        /// <summary>
        /// Returns a short from the stream
        /// </summary>
        /// <returns></returns>
        public int ReadUShort()
        {
            return _rc4.Xorshort(_littleEndianInput.ReadUShort());
        }
        #endregion

        #region ReadShort
        /// <summary>
        /// Returns a short from the stream
        /// </summary>
        /// <returns></returns>
        public short ReadShort()
        {
            return (short) _rc4.Xorshort(_littleEndianInput.ReadUShort());
        }
        #endregion

        #region ReadInt
        /// <summary>
        /// Returns an integer from the stream
        /// </summary>
        /// <returns></returns>
        public int ReadInt()
        {
            return _rc4.XorInt(_littleEndianInput.ReadInt());
        }
        #endregion

        #region ReadLong
        /// <summary>
        /// Returns a long from the stream
        /// </summary>
        /// <returns></returns>
        public long ReadLong()
        {
            return _rc4.XorLong(_littleEndianInput.ReadLong());
        }
        #endregion
    }
}

