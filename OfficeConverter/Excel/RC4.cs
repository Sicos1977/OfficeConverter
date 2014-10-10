using System.Collections.Generic;

namespace OfficeConverter.Excel
{
    // ReSharper disable once InconsistentNaming
    internal class RC4
    {
        private readonly byte[] _bytes = new byte[256];
        private int _i;
        private int _j;

        public RC4(IList<byte> key)
        {
            var keyLength = key.Count;

            for (var i = 0; i < 256; i++)
                _bytes[i] = (byte) i;

            for (int i = 0, j = 0; i < 256; i++)
            {
                j = (j + key[i%keyLength] + _bytes[i]) & 255;
                var temp = _bytes[i];
                _bytes[i] = _bytes[j];
                _bytes[j] = temp;
            }

            _i = 0;
            _j = 0;
        }

        public byte Output()
        {
            _i = (_i + 1) & 255;
            _j = (_j + _bytes[_i]) & 255;

            var temp = _bytes[_i];
            _bytes[_i] = _bytes[_j];
            _bytes[_j] = temp;

            return _bytes[(_bytes[_i] + _bytes[_j]) & 255];
        }

        public void Encrypt(byte[] in1)
        {
            for (var i = 0; i < in1.Length; i++)
            {
                in1[i] = (byte) (in1[i] ^ Output());
            }
        }

        public void Encrypt(byte[] in1, int offSet, int len)
        {
            var end = offSet + len;
            for (var i = offSet; i < end; i++)
                in1[i] = (byte) (in1[i] ^ Output());
        }
    }
}