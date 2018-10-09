using System.ServiceModel;

namespace OfficeConverterService.Interfaces
{
    [ServiceContract(Namespace = "http://www.achmea.nl/DocumentServices/OfficeConverter/Converter")]

    public interface IConverter
    {
        /// <summary>
        /// Converts the given <paramref name="inputFile"/> to the given <paramref name="outputFile"/>
        /// </summary>
        /// <param name="inputFile">The office file</param>
        /// <param name="outputFile">The output pdf file</param>
        [OperationContract]
        void ConvertFile(string inputFile, string outputFile);
    }
}
