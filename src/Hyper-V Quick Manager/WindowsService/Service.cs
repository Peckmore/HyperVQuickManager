using HyperVQuickManager.Properties;
using System;
using System.Diagnostics;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.ServiceProcess;

namespace HyperVQuickManager.WindowsService
{
    public partial class Service : ServiceBase
    {
        #region Fields

        ServiceHost _serviceHost;

        #endregion

        #region Construction

        public Service()
        {
            InitializeComponent();

            // Set the service name
            ServiceName = Program.ServiceName;
        }

        #endregion

        #region Methods

        #region Protected

        protected override void OnStart(string[] args)
        {
            try
            {
                // We instantiate a ServiceHost, passing in the type of our VmManager class which will be hosted via WCF,
                // specifying that we want to use Named Pipes for communication.
                _serviceHost = new ServiceHost(typeof(VmManager), new Uri("net.pipe://localhost"));
                _serviceHost.AddServiceEndpoint(typeof(IVmManager), new NetNamedPipeBinding(), Program.NamedPipeName);

#if DEBUG
                // If we are running in debug mode then we want to include the full exception details in all
                // faults that occur to help us track down the issue.

                var debugBehavior = _serviceHost.Description.Behaviors.Find<ServiceDebugBehavior>();
                if (debugBehavior == null)
                    _serviceHost.Description.Behaviors.Add(new ServiceDebugBehavior { IncludeExceptionDetailInFaults = true });
                else if (!debugBehavior.IncludeExceptionDetailInFaults)
                    debugBehavior.IncludeExceptionDetailInFaults = true;
#endif

                // Finally, with all prep done we can open the ServiceHost.
                _serviceHost.Open();
            }
            catch (Exception ex)
            {
                Program.GetEventLog().WriteEntry(string.Format(Strings.Error_ServiceFailedToStart, ex.Message), EventLogEntryType.Error, (int)EventLogEntryId.FailedToStart);
                throw;
            }
        }
        protected override void OnStop() => _serviceHost.Close(); // In order to stop the service the only thing we need to do is close our ServiceHost.

        #endregion

        #region Public Static

#if DEBUG
        /// <summary>
        /// A debug method that creates an instance of the Windows Service as an object, starts it and returns it.
        /// </summary>
        public static Service Start()
        {
            var svc = new Service();
            svc.OnStart(new string[0]);
            return svc;
        }
#endif

        #endregion

        #endregion
    }
}