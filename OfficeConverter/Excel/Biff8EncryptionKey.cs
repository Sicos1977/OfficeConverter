using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;

namespace OfficeConverter.Excel
{
    internal class Biff8EncryptionKey
    {
        // these two constants coincidentally have the same value
        private const int KeyDigestLength = 5;
        private const int PasswordHashNumberOfBytesUsed = 5;

        private readonly byte[] _keyDigest;

        /**
         * Create using the default password and a specified docId
         * @param docId 16 bytes
         */
        public static Biff8EncryptionKey Create(byte[] docId)
        {
            return new Biff8EncryptionKey(CreateKeyDigest("VelvetSweatshop", docId));
        }
        public static Biff8EncryptionKey Create(String password, byte[] docIdData)
        {
            return new Biff8EncryptionKey(CreateKeyDigest(password, docIdData));
        }

        internal Biff8EncryptionKey(byte[] keyDigest)
        {
            if (keyDigest.Length != KeyDigestLength)
            {
                // TODO: Fixen
                //throw new ArgumentException("Expected 5 byte key digest, but got " + HexDump.ToHex(keyDigest));
            }
            _keyDigest = keyDigest;
        }

        internal static byte[] CreateKeyDigest(String password, byte[] docIdData)
        {
            Check16Bytes(docIdData, "docId");
            var nChars = Math.Min(password.Length, 16);
            var passwordData = new byte[nChars * 2];
            for (var i = 0; i < nChars; i++)
            {
                var chr = password[i];
                passwordData[i * 2 + 0] = (byte)((chr << 0) & 0xFF);
                passwordData[i * 2 + 1] = (byte)((chr << 8) & 0xFF);
            }

            using (MD5 md5 = new MD5CryptoServiceProvider())
            {
                var passwordHash = md5.ComputeHash(passwordData);

                md5.Initialize();

                var data = new byte[PasswordHashNumberOfBytesUsed * 16 + docIdData.Length * 16];

                var offset = 0;
                for (var i = 0; i < 16; i++)
                {
                    Array.Copy(passwordHash, 0, data, offset, PasswordHashNumberOfBytesUsed);
                    offset += PasswordHashNumberOfBytesUsed;// passwordHash.Length;
                    Array.Copy(docIdData, 0, data, offset, docIdData.Length);
                    offset += docIdData.Length;
                }
                var kd = md5.ComputeHash(data);
                var result = new byte[KeyDigestLength];
                Array.Copy(kd, 0, result, 0, KeyDigestLength);
                md5.Clear();

                return result;
            }
        }

        /**
         * @return <c>true</c> if the keyDigest is compatible with the specified saltData and saltHash
         */
        public bool Validate(byte[] saltData, byte[] saltHash)
        {
            Check16Bytes(saltData, "saltData");
            Check16Bytes(saltHash, "saltHash");

            // validation uses the RC4 for block zero
            var rc4 = CreateRC4(0);
            var saltDataPrime = new byte[saltData.Length];
            Array.Copy(saltData, saltDataPrime, saltData.Length);
            rc4.Encrypt(saltDataPrime);

            var saltHashPrime = new byte[saltHash.Length];
            Array.Copy(saltHash, saltHashPrime, saltHash.Length);
            rc4.Encrypt(saltHashPrime);

            using (MD5 md5 = new MD5CryptoServiceProvider())
            {
                var finalSaltResult = md5.ComputeHash(saltDataPrime);
                return Arrays.Equals(saltHashPrime, finalSaltResult);
            }
        }

        private static byte[] Xor(IList<byte> a, IList<byte> b)
        {
            if (b == null) throw new ArgumentNullException("b");
            var c = new byte[a.Count];
            for (var i = 0; i < c.Length; i++)
                c[i] = (byte)(a[i] ^ b[i]);
            return c;
        }

        private static void Check16Bytes(ICollection<byte> data, string argName)
        {
            if (data.Count != 16)
            {
                // TODO: Fixen
                //throw new ArgumentException("Expected 16 byte " + argName + ", but got " + HexDump.ToHex(data));
            }
        }

        //private static ConcatBytes()

        /**
         * The {@link RC4} instance needs to be Changed every 1024 bytes.
         * @param keyBlockNo used to seed the newly Created {@link RC4}
         */
        internal RC4 CreateRC4(int keyBlockNo)
        {
            using (MD5 md5 = new MD5CryptoServiceProvider())
            {
                using (var baos = new MemoryStream(4))
                {
                    new LittleEndianOutputStream(baos).WriteInt(keyBlockNo);
                    var baosToArray = baos.ToArray();
                    var data = new byte[baosToArray.Length + _keyDigest.Length];
                    Array.Copy(_keyDigest, 0, data, 0, _keyDigest.Length);
                    Array.Copy(baosToArray, 0, data, _keyDigest.Length, baosToArray.Length);

                    var digest = md5.ComputeHash(data);
                    return new RC4(digest);
                }
            }
        }


        /**
         * Stores the BIFF8 encryption/decryption password for the current thread.  This has been done
         * using a {@link ThreadLocal} in order to avoid further overloading the various public APIs
         * (e.g. {@link HSSFWorkbook}) that need this functionality.
         */
        [ThreadStatic]
        private static String _userPasswordTls;

        /**
         * @return the BIFF8 encryption/decryption password for the current thread.
         * <code>null</code> if it is currently unSet.
         */
        public static String CurrentUserPassword
        {
            get
            {
                return _userPasswordTls;
            }
            set
            {
                _userPasswordTls = value;
            }
        }
    }

}
