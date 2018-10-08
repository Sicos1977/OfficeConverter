using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeConverterService
{
    /// <summary>
    /// Deze classe start de verschillende onderdelen van de Email2Storage recognition service.
    /// Tot op heden is dit alleen de FileReader classe
    /// </summary>
    internal class RecognitionWorker
    {
        #region Fields
        /// <summary>
        /// Host for the <see cref="Converter"/> service
        /// </summary>
        private ServiceHost _converterServiceHost;

        /// <summary>
        /// The <see cref="_converterServiceHost"/> base address
        /// </summary>
        private readonly Uri _converterBaseAddress;

        /// <summary>
        /// When set to <c>true</c> then metadata discovery is enabled
        /// </summary>
        private readonly bool _enableMetadata;

        /// <summary>
        /// The max message size
        /// </summary>
        private readonly long _maxMessageSize;

        /// <summary>
        /// Interval in secondens before the service checks if something has changed inside the database 
        /// </summary>
        private readonly int _changeTimeout;

        /// <summary>
        /// When set to <c>true</c> <see cref="CheckForChanges"/> will keep running
        /// </summary>
        private volatile bool _runWorker;

        /// <summary>
        /// A list to keep track of all our started taks
        /// </summary>
        private readonly List<Task> _runningTasks = new List<Task>();

        /// <summary>
        /// Used to put the service to sleep for some time
        /// </summary>
        private readonly ManualResetEvent _sleepServiceManualResetEvent = new ManualResetEvent(false);
        #endregion

        #region Constructor
        /// <summary>
        /// Maakt dit object en vult alle benodigde properties
        /// </summary>
        /// <param name="converterBaseAddress">The <see cref="_converterServiceHost"/> base address</param>
        /// <param name="enableMetadata">When set to <c>true</c> then metadata discovery is enabled</param>
        /// <param name="maxMessageSize">The max message size</param>
        /// <param name="changeTimeout">Interval in secondens before the service checks if something has changed inside the database </param>
        internal RecognitionWorker(string converterBaseAddress, 
                                   bool enableMetadata, 
                                   long maxMessageSize,
                                   int changeTimeout)
        {
            _converterBaseAddress = new Uri(converterBaseAddress);
            _enableMetadata = enableMetadata;
            _maxMessageSize = maxMessageSize;
            _changeTimeout = changeTimeout;
        }
        #endregion

        #region Start
        /// <summary>
        /// Start the different parts of the service
        /// </summary>
        internal void Start()
        {
            try
            {
                _converterServiceHost = new ServiceHost(typeof(Converter), _converterBaseAddress);
                var serviceDebugBehavior = _converterServiceHost.Description.Behaviors.Find<ServiceDebugBehavior>();
                if (serviceDebugBehavior == null)
                {
                    _converterServiceHost.Description.Behaviors.Add(new ServiceDebugBehavior()
                    {
                        IncludeExceptionDetailInFaults = true
                    });
                }
                else
                {
                    if (!serviceDebugBehavior.IncludeExceptionDetailInFaults)
                        serviceDebugBehavior.IncludeExceptionDetailInFaults = true;
                }

                var basicHttpBinding = new BasicHttpBinding
                {
                    MaxReceivedMessageSize = _maxMessageSize
                };

                const string endpointName = "basicHttpBindingEndpoint";

                _converterServiceHost.AddServiceEndpoint(typeof(Interfaces.IConverter), basicHttpBinding, endpointName);

                if (_enableMetadata)
                {
                    var serviceMetadataBehavior = new ServiceMetadataBehavior
                    {
                        HttpGetEnabled = true,
                        MetadataExporter = {PolicyVersion = PolicyVersion.Policy15}
                    };

                    _converterServiceHost.Description.Behaviors.Add(serviceMetadataBehavior);
                }

                _converterServiceHost.Open();

                //foreach (var baseAddress in _converterServiceHost.BaseAddresses)
                //    _systemLogs.Insert(ServiceName.Recognition,
                //        "Recognition gestart op basis adres '" + baseAddress.AbsoluteUri + "/" + endpointName +
                //        "', maximale bericht grootte " + FileManager.GetFileSizeString(_maxMessageSize));

                _runWorker = true;
                _runningTasks.Add(Task.Factory.StartNew(CheckForChanges, TaskCreationOptions.LongRunning));
            }
            catch (Exception exception)
            {
                //_errorLogs.Insert(exception);

                if (_converterServiceHost.State == CommunicationState.Opened)
                    _converterServiceHost.Close();
            }
        }
        #endregion

        #region Stop
        /// <summary>
        /// Stops the different parts of the service
        /// </summary>
        internal void Stop()
        {
            try
            {
                if (_converterServiceHost.State == CommunicationState.Opened)
                    _converterServiceHost.Close();

                _runWorker = false;
                _sleepServiceManualResetEvent.Set();

                // Wachten totdat alle tasks zijn gestopt
                foreach (var runningTask in _runningTasks)
                    runningTask.Wait();

                //_systemLogs.Insert(ServiceName.Recognition, "Recognition gestopt");
            }
            catch (Exception exception)
            {
                //_errorLogs.Insert(exception);
            }
        }
        #endregion

        #region CheckForChanges
        /// <summary>
        /// Checks if there are changes made in the database
        /// </summary>
        private void CheckForChanges()
        {
            while (_runWorker)
            {
                try
                {
                    if (_runWorker)
                        _sleepServiceManualResetEvent.WaitOne(_changeTimeout * 1000);

                    //if (!Lists.Instance.Changed(out var name)) continue;
                    //Lists.Instance.Refresh();
                    //_systemLogs.Insert(ServiceName.Recognition,
                    //    !string.IsNullOrEmpty(name)
                    //        ? $"Lijsten opnieuw ingeladen omdat de lijst '{name}' is aangepast"
                    //        : "Lijsten opnieuw ingeladen omdat 1 of meerdere lijsten zijn verwijdert/aangepast");
                }
                catch (Exception exception)
                {
                    //_errorLogs.Insert(exception);
                    //if (_runWorker)
                    //    _sleepServiceManualResetEvent.WaitOne(_changeTimeout * 1000);
                }
            }
        }
        #endregion
    }
}