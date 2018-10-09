using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using OfficeConverter;

namespace OfficeConverterService
{
    [ServiceBehavior(ConcurrencyMode = ConcurrencyMode.Multiple)]
    public class Converter : Interfaces.IConverter
    {
        #region Implementation of IConverter
        public void ConvertFile(string inputFile, string outputFile)
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}
