using System;
using System.ServiceProcess;
using System.Threading;

namespace OfficeConverterService
{
    public partial class Service1 : ServiceBase
    {
        #region Fields
        /// <summary>
        /// The <see cref="ConverterWorker"/> object
        /// </summary>
        private ConverterWorker _converterWorker;

        /// <summary>
        /// The thread on wich the <see cref="ConverterWorker"/> is started
        /// </summary>
        private Thread _convertedWorkerTread;
        #endregion

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                //_systemLogs.Insert(Server.BusinessLogic.ServiceName.Recognition, "Service gestart");

                //var applicationSettings = ApplicationSettings.Instance;
                //var serviceName = Server.BusinessLogic.ServiceName.Recognition.ToString();

                // Ophalen van de service instellingen
                //var recognitionBaseAddress = applicationSettings.GetValueAsString(serviceName, "RecognitionBaseAddress");
                //var enableMetadata = applicationSettings.GetValueAsBool(serviceName, "EnableMetadata");
                //var maxMessageSize = applicationSettings.GetValueAsLong(serviceName, "MaxMessageSize");
                //var changeTimeout= applicationSettings.GetValueAsInt(serviceName, "ChangeTimeout");

                var recognitionBaseAddress = "http://localhost/converterworker:45001";
                var enableMetadata = true;
                var maxMessageSize = 100000;
                var changeTimeout = 60;

                _converterWorker = new ConverterWorker(recognitionBaseAddress, 
                    enableMetadata, 
                    maxMessageSize,
                    changeTimeout);

                _convertedWorkerTread = new Thread(_converterWorker.Start);
                _converterWorker.Start();
            }
            catch (Exception exception)
            {
                //_errorLogs.Insert(exception);
                throw;
            }
        }

#if (DEBUG)
        #region DebugOnStart
        /// <summary>
        /// Starts the service in debugging mode from Visual Studio
        /// </summary>
        public void DebugOnStart()
        {
            OnStart(null);
        }
        #endregion
#endif

        /// <summary>
        /// Stops the service
        /// </summary>
        protected override void OnStop()
        {
            try
            {
                _converterWorker.Stop();
                _convertedWorkerTread.Join();
                //_systemLogs.Insert(Server.BusinessLogic.ServiceName.Recognition, "Service gestopt");
            }
            catch (Exception exception)
            {
                //_errorLogs.Insert(exception);
            }
        }
    }
}
