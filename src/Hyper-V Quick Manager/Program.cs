/*
 * Project inspiration - https://blogs.msdn.microsoft.com/jorman/2010/01/24/hyper-v-manager/
 * Service installer code example - https://stackoverflow.com/a/1195621/
 * Hyper-V WMI Provider API - https://docs.microsoft.com/en-gb/previous-versions/windows/desktop/virtual/windows-virtualization-portal
 * Hyper-V WMI Provider (v2) API - https://docs.microsoft.com/en-gb/windows/desktop/HyperV_v2/windows-virtualization-portal
 */

using HyperVQuickManager.Properties;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections;
using System.Configuration.Install;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.ServiceProcess;
using System.Windows.Forms;
using HyperVQuickManager.WindowsService;

namespace HyperVQuickManager
{
    internal static class Program
    {
        #region Properties

        /// <summary>
        /// The display name of the application.
        /// </summary>
        internal static string ApplicationName => (Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false).FirstOrDefault() as AssemblyTitleAttribute)?.Title;
        /// <summary>
        /// The name of the Named Pipe that is used by the application and Windows Service for communications.
        /// </summary>
        internal static string NamedPipeName => "HyperVQuickManager";
        /// <summary>
        /// The Service Display Name that will be used if the application is installed as a Windows Service.
        /// </summary>
        internal static string ServiceDisplayName => $"{ApplicationName} Service";
        /// <summary>
        /// The Service Name that will be used if the application is installed as a Windows Service.
        /// </summary>
        internal static string ServiceName => "vmqms";
        /// <summary>
        /// Specifies whether the program is running as a desktop application of a Windows Service.
        /// </summary>
        internal static bool UserInteractive { get; private set; }
        /// <summary>
        /// A property indicating whether we are going to use the v1 or v2 Hyper-V WMI Provider. This is set to true for any version of Windows greater than 6.2 (i.e., Windows 8.1 and Windows Server 2012 R2 onwards).
        /// </summary>
        internal static bool UseV2Api => Environment.OSVersion.Version > new Version(6, 2);

        #endregion

        #region Methods

        #region Event Handlers

        private static void TaskDialog_Opened(object sender, EventArgs e)
        {
            // This is to fix a bug in the Windows API Code Pack.
            // https://stackoverflow.com/a/22576707/2678851
            if (sender is TaskDialog taskDialog)
            {
                taskDialog.Icon = taskDialog.Icon;
                taskDialog.InstructionText = taskDialog.InstructionText;
            }
        }

        #endregion

        #region Private

        private static AssemblyInstaller GetServiceInstaller() => new AssemblyInstaller(typeof(Service).Assembly, null) { UseNewContext = true };
        private static void InstallService()
        {
            if (IsServiceInstalled())
                return;

            using (var installer = GetServiceInstaller())
            {
                var state = new Hashtable();
                try
                {
                    installer.Install(state);
                    installer.Commit(state);
                }
                catch
                {
                    installer.Rollback(state);
                }
            }
        }
        [STAThread]
        private static void Main(string[] args)
        {
            // In order to keep the program compact and easily distributable
            // everything is inside a single executable (both the tray icon
            // and the Windows Service). As a result we need to detect in
            // which context the application is running to determine what we
            // need to do. If running as a Windows Service Environment.UserInteractive
            // will return false.
            UserInteractive = Environment.UserInteractive;

            // Check the host OS to determine whether we support it.
            if (Environment.OSVersion.Version.Major < 6)
                // This is an OS before Windows Vista, which is not supported.
                UnsupportedOs();
            else if (Environment.OSVersion.Version.Major == 6)
            {
                // If the major version is 6 then we are looking at a version of Windows anywhere from
                // Windows Vista/Server 2008 to Windows 8.1/Server 2012 R2.

                // If the minor version is less than 2 then we are either on Windows Vista/Server 2008 (0) or
                // Windows 7/Server 2008 R2 (1). In both of these cases, Hyper-V was not available for
                // desktop versions of Windows, only server variants, so we need to check whether we are
                // running on a server OS. If not, we let the user know that their OS is not supported.
                if (Environment.OSVersion.Version.Minor < 2 && !NativeMethods.IsOS(29)) // 29 = OS_ANYSERVER
                    UnsupportedOs();
            }

            // If we get this far then we have passed the OS check so we can now proceed
            // to run the application.
            if (UserInteractive)
            {
                // We're running on the desktop, so run as a tray icon app.

                if (args.Length == 0)
                {
                    // No arguments have been passed in, so we just run the tray icon application as normal.

                    // Check if we are running as admin as we need admin rights to query Hyper-V using WMI. If
                    // we are not running as admin check if the Hyper-V Quick Manager service is installed. If it
                    // is then we are ok to run as non-admin as the service will do the heavy-lifting for us and
                    // we just use WCF to grab the data from it and send commands to it.
                    if (!NativeMethods.IsUserAnAdmin() && !IsServiceInstalled())
                        RequiresAdmin();

                    // Run our tray icon.
                    Application.Run(new MainContext());
                }
                else if (args.Length == 1)
                {
                    // A single argument has been passed in, which is supported. We check to see if the argument
                    // is valid and either take appropriate action, or return an error to the user if it is not
                    // valid.

                    switch (args[0])
                    {
                        case "/i":
                        case "-i":
                        case "--install":
                            if (!NativeMethods.IsUserAnAdmin())
                                RequiresAdminToInstall();

                            InstallService();
                            StartService();
                            break;
                        case "/u":
                        case "-u":
                        case "--uninstall":
                            if (!NativeMethods.IsUserAnAdmin())
                                RequiresAdminToInstall();

                            StopService();
                            UninstallService();
                            break;
                        default:
                            UnsupportedArguments();
                            break;
                    }
                }
                else
                    // If more than one argument has been passed in then we return an error to the user as this
                    // is not supported.
                    UnsupportedArguments();
            }
            else
            {
                // We're running as a Windows Service, so start the service instance.

                // Check if we are running as admin as we need admin rights to query Hyper-V using WMI.
                if (!NativeMethods.IsUserAnAdmin())
                    RequiresAdmin();

                // Start the service
                ServiceBase.Run(new Service());
            }
        }
        private static void RequiresAdmin()
        {
            if (UserInteractive)
                // We're running on the desktop so show the user a message box to indicate that
                // we need admin rights.
                ShowDialog(string.Format(Strings.Error_RequiresAdminDesktop_Text, ApplicationName), string.Format(Strings.Error_RequiresAdminDesktop_InstructionText, ApplicationName), TaskDialogStandardButtons.Ok, TaskDialogStandardIcon.Error);
            else
                // Log a message in the event log to indicate that we need admin rights.
                GetEventLog().WriteEntry(string.Format(Strings.Error_RequiresAdminService, ServiceDisplayName), EventLogEntryType.Error, (int)EventLogEntryId.RequiresAdmin);

#if !DEBUG
            // Exit the application/service.
            Environment.Exit(1);
#endif
        }
        private static void RequiresAdminToInstall()
        {
            if (UserInteractive)
                // We're running on the desktop so show the user a message box to indicate that
                // we need admin rights to install the service.
                ShowDialog(string.Format(Strings.Error_RequiresAdminToInstall_Text, ApplicationName), string.Format(Strings.Error_RequiresAdminToInstall_InstructionText, ApplicationName), TaskDialogStandardButtons.Ok, TaskDialogStandardIcon.Error);

#if !DEBUG
            // Exit the application/service.
            Environment.Exit(1);
#endif
        }
        private static void StartService()
        {
            if (!IsServiceInstalled())
                return;

            using (ServiceController controller = new ServiceController(ServiceName))
            {
                if (controller.Status != ServiceControllerStatus.Running)
                {
                    controller.Start();
                    controller.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                }
            }
        }
        private static void StopService()
        {
            if (!IsServiceInstalled())
                return;

            using (ServiceController controller = new ServiceController(ServiceName))
            {
                if (controller.Status != ServiceControllerStatus.Stopped)
                {
                    controller.Stop();
                    controller.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(10));
                }
            }
        }
        private static void UninstallService()
        {
            if (!IsServiceInstalled())
                return;

            using (var installer = GetServiceInstaller())
            {
                var state = new Hashtable();
                installer.Uninstall(state);
            }
        }
        private static void UnsupportedArguments()
        {
            if (UserInteractive)
                // We're running on the desktop so show the user a message box to indicate this
                // is an unsupported OS.
                ShowDialog(string.Format(Strings.Error_Unsupported_Text, ApplicationName), string.Format(Strings.Error_Unsupported_InstructionText, ApplicationName), TaskDialogStandardButtons.Ok, TaskDialogStandardIcon.Error);

#if !DEBUG
            // Exit the application/service.
            Environment.Exit(1);
#endif
        }
        private static void UnsupportedOs()
        {
            if (UserInteractive)
                // We're running on the desktop so show the user a message box to indicate this
                // is an unsupported OS.
                ShowDialog(string.Format(Strings.Error_Unsupported_Text, ApplicationName), string.Format(Strings.Error_Unsupported_InstructionText, ApplicationName), TaskDialogStandardButtons.Ok, TaskDialogStandardIcon.Error);
            else
                // Log a message in the event log to indicate that the service cannot run on this OS.
                GetEventLog().WriteEntry(string.Format(Strings.Error_Unsupported_Text, ServiceDisplayName), EventLogEntryType.Error, (int)EventLogEntryId.UnsupportedOs);

#if !DEBUG
            // Exit the application/service.
            Environment.Exit(1);
#endif
        }

        #endregion

        #region Internal

        internal static EventLog GetEventLog()
        {
            // We're running as a service so initialise our event log class so that we can write
            // to the event log.
            var eventLog = new EventLog
            {
                Source = ServiceDisplayName,
                Log = "Application"
            };

            // Add our event log source if it doesn't already exist.
            if (!EventLog.SourceExists(eventLog.Source))
                EventLog.CreateEventSource(eventLog.Source, eventLog.Log);

            return eventLog;
        }
        internal static bool IsServiceInstalled() => ServiceController.GetServices("localhost").Any(s => s.ServiceName == ServiceName);
        internal static bool IsServiceRunning()
        {
            using (var controller = new ServiceController(ServiceName))
            {
                if (!IsServiceInstalled())
                    return false;

                return controller.Status == ServiceControllerStatus.Running;
            }
        }
        internal static TaskDialogResult ShowDialog(string text, string instructionText, TaskDialogStandardButtons buttons, TaskDialogStandardIcon icon)
        {
            // Create our task dialog instance - we use 'using' to take care
            // of disposing of the object when we're done with it.
            using (var taskDialog = new TaskDialog())
            {
                // Get the AssemblyTitle attribute of the project to automatically
                // give the dialog a caption matching the program name.
                taskDialog.Caption = ApplicationName;

                // Set the remaining properties on the dialog.
                taskDialog.Icon = icon;
                taskDialog.InstructionText = instructionText;
                taskDialog.StandardButtons = buttons;
                taskDialog.StartupLocation = TaskDialogStartupLocation.CenterOwner;
                taskDialog.Text = text;

                // Add a handler to the dialog opening event - this is to fix
                // a bug with the task dialog implementation.
                taskDialog.Opened += TaskDialog_Opened;
                try
                {
                    // Show the dialog and return the result.
                    return taskDialog.Show();
                }
                finally
                {
                    // Remove our handler for the opening event.
                    taskDialog.Opened -= TaskDialog_Opened;
                }
            }
        }

        #endregion

        #endregion
    }
}