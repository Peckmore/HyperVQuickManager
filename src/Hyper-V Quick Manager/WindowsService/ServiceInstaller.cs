using HyperVQuickManager.Properties;
using System.ComponentModel;
using System.Configuration.Install;

namespace HyperVQuickManager.WindowsService
{
    [RunInstaller(true)]
    public partial class ServiceInstaller : Installer
    {
        public ServiceInstaller()
        {
            InitializeComponent();

            serviceInstaller.Description = string.Format(Strings.Service_Description, Program.ApplicationName);
            serviceInstaller.DisplayName = Program.ServiceDisplayName;
            serviceInstaller.ServiceName = Program.ServiceName;
        }
    }
}