using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Management;
using System.ServiceModel;

namespace HyperVQuickManager
{
    [ServiceBehavior(InstanceContextMode = InstanceContextMode.Single, ConcurrencyMode = ConcurrencyMode.Multiple)]
    internal class VmManager : IVmManager, IDisposable
    {
        #region Constants

        private const int WmiInterval = 2; // This is the interval WMI will use to check for events.

        #endregion

        #region Fields

        private readonly List<IVmManagerCallback> _callbackChannels = new List<IVmManagerCallback>();
        private ManagementEventWatcher _wmiEventWatcher;

        #endregion

        #region Events

        public event EventHandler<VmOverview> StateChanged;

        #endregion

        #region Construction & Finalization

        public VmManager()
        {
            // Create a ManagementScope object and connect to it. This object defines
            // the WMI path we are going to use for our WMI Event Watcher.
            var scope = new ManagementScope(RootWmiPath);
            scope.Connect();

            // Set up our WMI Event Watcher to monitor for changes in state to any Hyper-V
            // virtual machine. The watcher will monitor for changes every X seconds, where
            // X is a constant defined at the start of the class.
            var query = new WqlEventQuery($"SELECT * FROM __InstanceOperationEvent WITHIN {WmiInterval} WHERE TargetInstance ISA 'Msvm_ComputerSystem' AND TargetInstance.EnabledState <> PreviousInstance.EnabledState");
            _wmiEventWatcher = new ManagementEventWatcher(query) { Scope = scope };
            _wmiEventWatcher.EventArrived += WmiEventWatcher_EventArrived;
            _wmiEventWatcher.Start();
        }
        ~VmManager() => Dispose(false);

        #endregion

        #region Properties

        [SuppressMessage("ReSharper", "MemberCanBeMadeStatic.Local")]
        private string RootWmiPath => Program.UseV2Api ? "ROOT\\virtualization\\v2" : "ROOT\\virtualization"; // This is the root WMI path we will use, which varies depending on whether we use the v1 or v2 Hyper-V WMI provider.

        #endregion

        #region Methods

        #region Event Handlers

        /// <summary>
        /// This method handles events raised by our WMI Event Watcher. Events are raised when there is a change to an object specified in our WMI query.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void WmiEventWatcher_EventArrived(object sender, EventArrivedEventArgs e)
        {
            // If the WMI Event Watcher has fired an event, check that the args
            // contain a ManagementBaseObject.
            if (!(e.NewEvent["TargetInstance"] is ManagementBaseObject manObj))
                return;

            // If we have a ManagementBaseObject create a new VmOverview instance
            // using the virtual machine name and state taken from this object.
            var args = new VmOverview { Name = manObj["ElementName"] as string, State = (VmState)(ushort)manObj["EnabledState"] };

            // Iterate through every callback channel and fire the StateChanged
            // method.
            foreach (var channel in _callbackChannels)
                channel.StateChanged(args);

            // Raise the VmStateChanged event for all subscribed handlers.
            StateChanged?.Invoke(this, args);
        }

        #endregion

        #region Private

        private void Dispose(bool disposing)
        {
            // Check whether we are disposing or finalizing.
            if (disposing)
            {
                // If we are disposing check to see if we have instantiated the
                // WMI Event Watcher object.
                if (_wmiEventWatcher != null)
                {
                    // If the object has been created then stop the watcher, unsubscribe
                    // from any events, dispose of the object and set our variable to null.
                    _wmiEventWatcher.Stop();
                    _wmiEventWatcher.EventArrived -= WmiEventWatcher_EventArrived;
                    _wmiEventWatcher.Dispose();
                    _wmiEventWatcher = null;
                }

                // Clear out our list of callback channels.
                _callbackChannels.Clear();
            }
        }
        private IEnumerable<ManagementObject> GetVmInternal(string name = null)
        {
            // Create our WMI query string to get one virtual machine, or all virtual
            // machines, from the list of machines configured on this system. If the
            // 'name' parameter is null then we get all virtual machines, otherwise we
            // look for a machine with a matching name.
            var queryString = $"SELECT * FROM Msvm_ComputerSystem WHERE Name!='{Environment.MachineName}'";
            if (name != null)
                queryString += $" AND ElementName='{name}'";

            // Create a ManagementScope object which defines the WMI path we are going
            // to use for our WMI query.
            var manScope = new ManagementScope(RootWmiPath);

            // Create a WMI query object using the querystring we defined previously.
            var queryObj = new ObjectQuery(queryString);

            // Create a searcher object, passing in our search parameters as previously
            // defined.
            var vmSearcher = new ManagementObjectSearcher(manScope, queryObj);

            // Run the search, and cast the results to a list of ManagementObject.
            return vmSearcher.Get().Cast<ManagementObject>().ToList();
        }

        #endregion

        #region Public

        public void Dispose()
        {
            // Trigger our internal dispose method, indicating that we are calling it
            // from Dispose rather than finalization.
            Dispose(true);

            // Tell the GC to suppress finalization for this object as it has already
            // been disposed.
            GC.SuppressFinalize(this);
        }
        public IList<VmOverview> GetVm(string name = null) => GetVmInternal(name).Select(vm => new VmOverview { Name = vm["ElementName"] as string, State = (VmState)Convert.ToInt32(vm["EnabledState"]) }).ToList(); // Call our internal method GetVmInternal and cast the results to a list of VmOverview.
        public StateChangeResponse RequestVmStateChange(string name, VmState vmState)
        {
            // Get the WMI Management Object for the virtual machine we are interested in.
            var vm = GetVmInternal(name).FirstOrDefault();

            // Check to see whether we found a matching virtual machine.
            if (vm != null)
            {
                // Get the parameters for the 'RequestStateChange' method.
                var inParams = vm.GetMethodParameters("RequestStateChange");

                // Set the 'RequestedState' parameter to the desired state.
                inParams["RequestedState"] = vmState;

                // Fire off the request to change the state.
                var outParams = vm.InvokeMethod("RequestStateChange", inParams, null);

                // Return the result of the method call.
                if (outParams != null)
                    return (StateChangeResponse)outParams["ReturnValue"];
            }

            // If we couldn't find a matching virtual machine then return a
            // 'Failed' response.
            return StateChangeResponse.Failed;
        }
        public StateChangeResponse ShutdownVm(string name)
        {
            // Get the WMI Management Object for the virtual machine we are interested in.
            var vm = GetVmInternal(name).FirstOrDefault();

            // If we found a matching virtual machine get the Msvm_ShutdownComponent for it.
            var shutdowncomponent = vm?.GetRelated("Msvm_ShutdownComponent").Cast<ManagementObject>().FirstOrDefault();

            // Check to see whether we found a Shutdown Component
            if (shutdowncomponent != null)
            {
                // Get the parameters for the InitiateShudown method
                var inParams = shutdowncomponent.GetMethodParameters("InitiateShutdown");

                // Set the 'Force' and 'Reason' parameters.
                inParams["Force"] = true;
                inParams["Reason"] = Program.ApplicationName;

                // Invoke the method, passing in the parameters we just set.
                var outParams = shutdowncomponent.InvokeMethod("InitiateShutdown", inParams, null);

                // Return the result of the method call.
                if (outParams != null)
                    return (StateChangeResponse)outParams["ReturnValue"];
            }

            // If we couldn't find a matching virtual machine then return a
            // 'Failed' response.
            return StateChangeResponse.Failed;
        }
        public void Subscribe()
        {
            // Clients should call this method to "subscribe"/register themselves with
            // the server in order to receive callbacks when virtual machine state
            // changes occur.

            // Grab the callback channel for the client calling this method.
            var channel = OperationContext.Current.GetCallbackChannel<IVmManagerCallback>();

            // If we don't already have a reference to this client then add them to our list.
            if (!_callbackChannels.Contains(channel))
                _callbackChannels.Add(channel);
        }

        #endregion

        #endregion
    }
}