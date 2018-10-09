#if !(DEBUG)
using System.ServiceProcess;
#endif

namespace OfficeConverterService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
#if(DEBUG)
            using (var service = new Service1())
            {
                service.DebugOnStart();
                System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);
            }
#else
            var servicesToRun = new ServiceBase[] {new Service1()};
            ServiceBase.Run(servicesToRun);
#endif
        }
    }
}
