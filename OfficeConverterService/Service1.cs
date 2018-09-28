using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace OfficeConverterService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
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

        protected override void OnStop()
        {
        }
    }
}
