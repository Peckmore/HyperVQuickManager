using System.ServiceModel;

namespace HyperVQuickManager
{
    public interface IVmManagerCallback
    {
        [OperationContract(IsOneWay = true)]
        void StateChanged(VmOverview overview);

    }
}