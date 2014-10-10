using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeConverter.Excel.Interfaces
{
    internal interface IBiffHeaderInput
    {
        /**
         * Read an unsigned short from the stream without decrypting
         */
        int ReadRecordSID();
        /**
         * Read an unsigned short from the stream without decrypting
         */
        int ReadDataSize();

        int Available();
    }
}
