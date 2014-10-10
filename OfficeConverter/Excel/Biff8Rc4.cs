using System;

namespace OfficeConverter.Excel
{
    /**
     * Used for both encrypting and decrypting BIFF8 streams. The internal
     * {@link RC4} instance is renewed (re-keyed) every 1024 bytes.
     *
     * @author Josh Micich
     */
    // ReSharper disable once InconsistentNaming
    internal class Biff8RC4
    {

        private const int Rc4RekeyingInterval = 1024;

        private RC4 _rc4;
        /**
         * This field is used to keep track of when to change the {@link RC4}
         * instance. The change occurs every 1024 bytes. Every byte passed over is
         * counted.
         */
        private int _streamPos;
        private int _nextRc4BlockStart;
        private int _currentKeyIndex;
        private bool _shouldSkipEncryptionOnCurrentRecord;

        private readonly Biff8EncryptionKey _key;

        public Biff8RC4(int initialOffset, Biff8EncryptionKey key)
        {
            if (initialOffset >= Rc4RekeyingInterval)
            {
                throw new Exception("InitialOffset (" + initialOffset + ")>"
                        + Rc4RekeyingInterval + " not supported yet");
            }

            _key = key;
            _streamPos = 0;
            RekeyForNextBlock();
            _streamPos = initialOffset;
            for (var i = initialOffset; i > 0; i--)
                _rc4.Output();
            
            _shouldSkipEncryptionOnCurrentRecord = false;
        }

        private void RekeyForNextBlock()
        {
            _currentKeyIndex = _streamPos / Rc4RekeyingInterval;
            _rc4 = _key.CreateRC4(_currentKeyIndex);
            _nextRc4BlockStart = (_currentKeyIndex + 1) * Rc4RekeyingInterval;
        }

        private int GetNextRC4Byte()
        {
            if (_streamPos >= _nextRc4BlockStart)
            {
                RekeyForNextBlock();
            }
            byte mask = _rc4.Output();
            _streamPos++;
            if (_shouldSkipEncryptionOnCurrentRecord)
            {
                return 0;
            }
            return mask & 0xFF;
        }

        public void StartRecord(int currentSid)
        {
            _shouldSkipEncryptionOnCurrentRecord = IsNeverEncryptedRecord(currentSid);
        }

        /**
         * TODO: Additionally, the lbPlyPos (position_of_BOF) field of the BoundSheet8 record MUST NOT be encrypted.
         *
         * @return <c>true</c> if record type specified by <c>sid</c> is never encrypted
         */
        private static bool IsNeverEncryptedRecord(int sid)
        {
            switch (sid)
            {
                case 0x809:
                case 0xe1:
                case 0x2F:
                    return true;

                default:
                    return false;
            }
        }

        /**
         * Used when BIFF header fields (sid, size) are being Read. The internal
         * {@link RC4} instance must step even when unencrypted bytes are read
         */
        public void SkipTwoBytes()
        {
            GetNextRC4Byte();
            GetNextRC4Byte();
        }

        public void Xor(byte[] buf, int pOffSet, int pLen)
        {
            var nLeftInBlock = _nextRc4BlockStart - _streamPos;
            if (pLen <= nLeftInBlock)
            {
                // simple case - this read does not cross key blocks
                _rc4.Encrypt(buf, pOffSet, pLen);
                _streamPos += pLen;
                return;
            }

            var offset = pOffSet;
            var len = pLen;

            // start by using the rest of the current block
            if (len > nLeftInBlock)
            {
                if (nLeftInBlock > 0)
                {
                    _rc4.Encrypt(buf, offset, nLeftInBlock);
                    _streamPos += nLeftInBlock;
                    offset += nLeftInBlock;
                    len -= nLeftInBlock;
                }
                RekeyForNextBlock();
            }
            // all full blocks following
            while (len > Rc4RekeyingInterval)
            {
                _rc4.Encrypt(buf, offset, Rc4RekeyingInterval);
                _streamPos += Rc4RekeyingInterval;
                offset += Rc4RekeyingInterval;
                len -= Rc4RekeyingInterval;
                RekeyForNextBlock();
            }
            // finish with incomplete block
            _rc4.Encrypt(buf, offset, len);
            _streamPos += len;
        }

        public int XorByte(int rawVal)
        {
            int mask = GetNextRC4Byte();
            return (byte)(rawVal ^ mask);
        }

        public int Xorshort(int rawVal)
        {
            int b0 = GetNextRC4Byte();
            int b1 = GetNextRC4Byte();
            int mask = (b1 << 8) + (b0 << 0);
            return rawVal ^ mask;
        }

        public int XorInt(int rawVal)
        {
            var b0 = GetNextRC4Byte();
            var b1 = GetNextRC4Byte();
            var b2 = GetNextRC4Byte();
            var b3 = GetNextRC4Byte();
            var mask = (b3 << 24) + (b2 << 16) + (b1 << 8) + (b0 << 0);
            return rawVal ^ mask;
        }

        public long XorLong(long rawVal)
        {
            var b0 = GetNextRC4Byte();
            var b1 = GetNextRC4Byte();
            var b2 = GetNextRC4Byte();
            var b3 = GetNextRC4Byte();
            var b4 = GetNextRC4Byte();
            var b5 = GetNextRC4Byte();
            var b6 = GetNextRC4Byte();
            var b7 = GetNextRC4Byte();
            var mask =
                (((long) b7) << 56)
                + (((long) b6) << 48)
                + (((long) b5) << 40)
                + (((long) b4) << 32)
                + (((long) b3) << 24)
                + (b2 << 16)
                + (b1 << 8)
                + (b0 << 0);
            return rawVal ^ mask;
        }
    }
}

