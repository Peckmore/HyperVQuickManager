using System.Collections.Generic;
using System.ServiceModel;

namespace HyperVQuickManager
{
    [ServiceContract(CallbackContract = typeof(IVmManagerCallback))]

    public interface IVmManager
    {
        [OperationContract]
        IList<VmOverview> GetVm(string name = null);
        [OperationContract]
        StateChangeResponse RequestVmStateChange(string name, VmState vmState);
        [OperationContract]
        StateChangeResponse ShutdownVm(string name);
        [OperationContract]
        void Subscribe();
    }
}