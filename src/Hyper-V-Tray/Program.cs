using Microsoft.Toolkit.Uwp.Notifications;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Windows.Forms;

namespace HyperVTray
{
    internal static class Program
    {
        #region Constants

        private const string ApplicationName = @"Hyper-V Tray";
        private const int BalloonTipTimeout = 2500;
        private static readonly ContextMenu ContextMenu;
        private static readonly NotifyIcon NotifyIcon;
        private static readonly ManagementEventWatcher WmiEventWatcher;
        private static readonly ManagementScope WmiManagementScope;
        private const int WmiRefreshInterval = 2;

        #endregion

        #region Fields

        private static string? _hyperVInstallFolder;
        private static string? _vmConnectPath;
        private static string? _vmManagerPath;

        #endregion

        #region Construction

        static Program()
        {
            ContextMenu = new ContextMenu();

            NotifyIcon = new NotifyIcon();
            NotifyIcon.DoubleClick += NotifyIcon_DoubleClick;
            NotifyIcon.MouseClick += NotifyIcon_MouseClick;

            // This is the root WMI path we will use, which varies depenUseV2WmiProviderwe use the v1 or v2 Hyper-V WMI provider.
            var rootWmiPath = UseV2WmiProvider ? "ROOT\\virtualization\\v2" : "ROOT\\virtualization";

            // Create a ManagementScope object and connect to it. This object defines
            // the WMI path we are going to use for our WMI Event Watcher.
            WmiManagementScope = new ManagementScope(rootWmiPath);

            // Set up our WMI Event Watcher to monitor for changes in state to any Hyper-V
            // virtual machine. The watcher will monitor for changes every X seconds, where
            // X is a constant defined at the start of the class.
            var query = new WqlEventQuery($"SELECT * FROM __InstanceOperationEvent WITHIN {WmiRefreshInterval} " +
                                          $"WHERE TargetInstance ISA 'Msvm_ComputerSystem' " +
                                          $"AND TargetInstance.EnabledState <> PreviousInstance.EnabledState");
            WmiEventWatcher = new ManagementEventWatcher(query) { Scope = WmiManagementScope };
            WmiEventWatcher.EventArrived += WmiEventWatcher_EventArrived;

            SystemEvents.DisplaySettingsChanged += SystemEvents_DisplaySettingsChanged;
        }

        #endregion

        #region Properties

        /// <summary>
        /// A property indicating whether we are going to use the v1 or v2 Hyper-V WMI Provider. This is set to true for any version of
        /// Windows greater than 6.2 (i.e., Windows 8.1 and Windows Server 2012 R2 onwards).
        /// </summary>
        internal static bool UseV2WmiProvider => Environment.OSVersion.Version > new Version(6, 2);

        #endregion

        #region Methods

        #region Event Handlers

        private static void ExitMenuItem_Click(object? sender, EventArgs e)
        {
            NotifyIcon.Visible = false;
            Application.Exit();
        }
        private static void NotifyIcon_DoubleClick(object? sender, EventArgs e)
        {
            OpenHyperVManager();
        }
        private static void NotifyIcon_MouseClick(object? sender, MouseEventArgs e)
        {
            GenerateContextMenu();
            NotifyIcon.ShowContextMenu(ContextMenu);
        }
        private static void PauseMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!RequestVmStateChange(menuItem.Parent.Name, VmState.Quiesce))
                {
                    ShowError(string.Format(ResourceHelper.Message_PauseVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void ResetMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!RequestVmStateChange(menuItem.Parent.Name, VmState.Reset))
                {
                    ShowError(ResourceHelper.Message_ResetVMFailed, menuItem.Parent.Name);
                }
            }
        }
        private static void ResumeMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!RequestVmStateChange(menuItem.Parent.Name, VmState.Enabled))
                {
                    ShowError(string.Format(ResourceHelper.Message_ResumeVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void SaveMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!RequestVmStateChange(menuItem.Parent.Name, VmState.Offline))
                {
                    ShowError(string.Format(ResourceHelper.Message_SaveStateVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void ShutDownMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!RequestVmShutdown(menuItem.Parent.Name))
                {
                    ShowError(string.Format(ResourceHelper.Message_ShutDownVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void StartMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!RequestVmStateChange(menuItem.Parent.Name, VmState.Enabled))
                {
                    ShowError(string.Format(ResourceHelper.Message_StartVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void SystemEvents_DisplaySettingsChanged(object? sender, EventArgs e)
        {
            Debug.WriteLine("Display settings change detected.");

            SetTrayIcon();
        }
        private static void TurnOffMenuItem_Click(object? sender, EventArgs e)
        {
            if (sender is MenuItem menuItem)
            {
                if (!RequestVmStateChange(menuItem.Parent.Name, VmState.Disabled))
                {
                    ShowError(string.Format(ResourceHelper.Message_PowerOffVMFailed, menuItem.Parent.Name));
                }
            }
        }
        private static void WmiEventWatcher_EventArrived(object? sender, EventArrivedEventArgs e)
        {
            // If the WMI Event Watcher has fired an event, check that the args contain a ManagementBaseObject.
            if (e.NewEvent["TargetInstance"] is ManagementBaseObject virtualMachine)
            {
                // Filter out "noisy" states, to just show important ones.
                var vmState = (VmState)(ushort)virtualMachine["EnabledState"];
                var vmName = virtualMachine["ElementName"].ToString();
                switch (vmState)
                {
                    case VmState.Enabled:
                    case VmState.Disabled:
                    case VmState.Offline:
                    case VmState.Quiesce:
                    case VmState.Reset:
                        ShowToast(vmName, vmState, false);
                        break;

                    case VmState.RunningCritical:
                    case VmState.OffCritical:
                    case VmState.StoppingCritical:
                    case VmState.SavedCritical:
                    case VmState.PausedCritical:
                    case VmState.StartingCritical:
                    case VmState.ResetCritical:
                    case VmState.SavingCritical:
                    case VmState.PausingCritical:
                    case VmState.ResumingCritical:
                    case VmState.FastSavedCritical:
                    case VmState.FastSavingCritical:
                        ShowToast(vmName, vmState, true);
                        break;
                }
            }
        }

        #endregion
            
        #region Private Static

        private static void ConnectToVm(string virtualMachineName)
        {
            var processInfo = new ProcessStartInfo
            {
                FileName = _vmConnectPath!,
                Arguments = $"localhost \"{virtualMachineName}\"",
            };
            Process.Start(processInfo);
        }
        private static void GenerateContextMenu()
        {
            // Clear the context menu.
            ContextMenu.MenuItems.Clear();

            // Get all VMs.
            var virtualMachines = GetVirtualMachines();

            // Create a menu entry for each VM.
            foreach (var virtualMachine in virtualMachines)
            {
                // Get the VM name.
                var virtualMachineName = virtualMachine["ElementName"].ToString();

                if (virtualMachineName != null)
                {
                    // Get the VM state.
                    var virtualMachineStatus = (VmState)Convert.ToInt32(virtualMachine["EnabledState"]);

                    // Generate the menu entry title for the VM.
                    var virtualMachineMenuTitle = virtualMachineName;
                    if (virtualMachineStatus != VmState.Disabled) // Stopped
                    {
                        virtualMachineMenuTitle += $" [{VmStateToString(virtualMachineStatus)}]";
                    }

                    // Create VM menu item.
                    var virtualMachineMenu = new MenuItem(virtualMachineMenuTitle) { Name = virtualMachineName };
                    var canResume = IsVmPaused(virtualMachineStatus); // Paused
                    var canStart = IsVmOff(virtualMachineStatus) || IsVmSaved(virtualMachineStatus); // Stopped or Saved

                    // Now generate control menu items for VM.

                    // Connect
                    if (_vmConnectPath != null)
                    {
                        var connectMenuItem = new MenuItem(ResourceHelper.Command_Connect);
                        connectMenuItem.Click += (_, _) => ConnectToVm(virtualMachineName);
                        virtualMachineMenu.MenuItems.Add(connectMenuItem);
                        virtualMachineMenu.MenuItems.Add(new MenuItem("-"));
                    }

                    if (canStart)
                    {
                        // Start
                        var startMenuItem = new MenuItem(ResourceHelper.Command_Start);
                        startMenuItem.Click += StartMenuItem_Click;
                        virtualMachineMenu.MenuItems.Add(startMenuItem);
                    }
                    else
                    {
                        // Turn Off
                        var stopMenuItem = new MenuItem(ResourceHelper.Command_TurnOff);
                        stopMenuItem.Click += TurnOffMenuItem_Click;
                        virtualMachineMenu.MenuItems.Add(stopMenuItem);

                        // Shut Down
                        if (!canResume)
                        {
                            var shutMenuDownItem = new MenuItem(ResourceHelper.Command_ShutDown);
                            shutMenuDownItem.Click += ShutDownMenuItem_Click;
                            virtualMachineMenu.MenuItems.Add(shutMenuDownItem);
                        }

                        // Save
                        var saveMenuStateItem = new MenuItem(ResourceHelper.Command_Save);
                        saveMenuStateItem.Click += SaveMenuItem_Click;
                        virtualMachineMenu.MenuItems.Add(saveMenuStateItem);

                        virtualMachineMenu.MenuItems.Add(new MenuItem("-"));
                        if (canResume)
                        {
                            // Resume
                            var resumeMenuItem = new MenuItem(ResourceHelper.Command_Resume);
                            resumeMenuItem.Click += ResumeMenuItem_Click;
                            virtualMachineMenu.MenuItems.Add(resumeMenuItem);
                        }
                        else
                        {
                            // Pause
                            var pauseMenuItem = new MenuItem(ResourceHelper.Command_Pause);
                            pauseMenuItem.Click += PauseMenuItem_Click;
                            virtualMachineMenu.MenuItems.Add(pauseMenuItem);
                        }

                        // Reset
                        var resetMenuItem = new MenuItem(ResourceHelper.Command_Reset);
                        resetMenuItem.Click += ResetMenuItem_Click;
                        virtualMachineMenu.MenuItems.Add(resetMenuItem);
                    }

                    // Add VM menu item to root menu.
                    ContextMenu.MenuItems.Add(virtualMachineMenu);
                }
            }

            if (virtualMachines.Any())
            {
                ContextMenu.MenuItems.Add(new MenuItem("-"));

                // Create a root menu item for the VM.
                var vmItem = new MenuItem("Strings.Menu_AllVirtualMachines");

                var subItems = new List<MenuItem>();
                var isOff = virtualMachines.Any(vm => IsVmOff((VmState)Convert.ToInt32(vm["EnabledState"])));
                var isPaused = virtualMachines.Any(vm => IsVmPaused((VmState)Convert.ToInt32(vm["EnabledState"])));
                var isRunning = virtualMachines.Any(vm => IsVmRunning((VmState)Convert.ToInt32(vm["EnabledState"])));
                var isSaved = virtualMachines.Any(vm => IsVmSaved((VmState)Convert.ToInt32(vm["EnabledState"])));

                // Start
                if (isOff || isSaved)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_Start)); //MenuItemStartAll_Click));
                }

                // Turn Off
                if (isRunning || isPaused)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_TurnOff)); //MenuItemStopAll_Click));
                }

                // Shut Down
                if (isRunning)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_ShutDown)); //MenuItemShutdownAll_Click));
                }

                // Save
                if (isRunning || isPaused)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_Save)); //MenuItemSaveAll_Click));
                }

                if (subItems.Any() && isRunning || isPaused)
                {
                    subItems.Add(new MenuItem("-"));
                }

                // Resume
                if (isPaused)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_Resume)); //MenuItemResumeAll_Click));
                }

                // Pause
                if (isRunning)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_Pause)); //MenuItemPauseAll_Click));
                }

                // Reset
                if (isRunning || isPaused)
                {
                    subItems.Add(new MenuItem(ResourceHelper.Command_Reset)); //MenuItemResetAll_Click));
                }

                vmItem.MenuItems.AddRange(subItems.ToArray());

                // Add the VM to the context menu.
                ContextMenu.MenuItems.Add(vmItem);
            }

            // Add `Hyper-V Manager` menu item.
            if (_vmManagerPath != null)
            {
                ContextMenu.MenuItems.Add(new MenuItem("-"));
                var managerItem = new MenuItem("Hyper-V Manager");
                managerItem.Click += (_, _) => OpenHyperVManager();
                ContextMenu.MenuItems.Add(managerItem);
            }

            // Add `Exit` menu item.
            ContextMenu.MenuItems.Add(new MenuItem("-"));
            var exitItem = new MenuItem("Exit");
            exitItem.Click += ExitMenuItem_Click;
            ContextMenu.MenuItems.Add(exitItem);
        
        }
        private static IEnumerable<ManagementObject> GetVirtualMachines(string? name = null)
        {
            // Create our WMI query string to get one virtual machine, or all virtual machines, from the list of machines configured
            // on this system. If the 'name' parameter is null then we get all virtual machines, otherwise we look for a machine with
            // a matching name.
            var queryString = $"SELECT * FROM Msvm_ComputerSystem WHERE Caption LIKE 'Virtual Machine'"; // WHERE Name!='{Environment.MachineName}'";
            if (name != null)
            {
                queryString += $" AND ElementName='{name}'";
            }

            // Create a WMI query object using the querystring we defined previously.
            var queryObj = new ObjectQuery(queryString);

            // Create a searcher object, passing in our search parameters as previously defined.
            var vmSearcher = new ManagementObjectSearcher(WmiManagementScope, queryObj);

            // Run the search, and cast the results to a list of ManagementObject.
            return vmSearcher.Get().Cast<ManagementObject>().ToList();
        }
        private static bool IsVmCritical(VmState state)
        {
            // Determine if the VmState enum value is for a critical state.
            return state switch
            {
                VmState.RunningCritical or VmState.OffCritical or VmState.PausedCritical or VmState.SavedCritical
                or VmState.FastSavedCritical or VmState.StartingCritical or VmState.SavingCritical or VmState.FastSavingCritical
                or VmState.StoppingCritical or VmState.PausingCritical or VmState.ResumingCritical or VmState.ResetCritical => true,
                _ => false,
            };
        }
        private static bool IsVmOff(VmState state)
        {
            // Determine if the VmState enum value is for an off state.
            return state switch
            {
                VmState.Disabled or VmState.OffCritical => true,
                _ => false,
            };
        }
        private static bool IsVmPaused(VmState state)
        {
            // Determine if the VmState enum value is for a paused state.
            return state switch
            {
                VmState.Paused or VmState.Quiesce or VmState.PausedCritical => true,
                _ => false,
            };
        }
        private static bool IsVmRunning(VmState state)
        {
            // Determine if the VmState enum value is for a running state.
            return state switch
            {
                VmState.Enabled or VmState.RunningCritical => true,
                _ => false,
            };
        }
        private static bool IsVmSaved(VmState state)
        {
            // Determine if the VmState enum value is for a saved state.
            return state switch
            {
                VmState.Suspended or VmState.Offline or VmState.SavedCritical or VmState.FastSaved or VmState.FastSavedCritical => true,
                _ => false,
            };
        }
        [STAThread]
        private static void Main()
        {
            // Application setup.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Detect Hyper-V components.
            _hyperVInstallFolder = @$"{Environment.GetEnvironmentVariable("ProgramFiles")}\Hyper-V\";
            if (!Directory.Exists(_hyperVInstallFolder))
            {
                ShowError("Hyper-V Tools not installed.");
                Application.Exit();
                return;
            }
            _vmConnectPath = $@"{Environment.GetEnvironmentVariable("SYSTEMROOT")}\System32\vmconnect.exe";
            if (!File.Exists(_vmConnectPath))
            {
                _vmConnectPath = null;
            }
            _vmManagerPath = $@"{Environment.GetEnvironmentVariable("SYSTEMROOT")}\System32\virtmgmt.msc";
            if (!File.Exists(_vmManagerPath))
            {
                _vmManagerPath = null;
            }

            // Initialize our ResourceHelper.
            ResourceHelper.Initialize(_hyperVInstallFolder);

            // Show our tray icon.
            NotifyIcon.Visible = true;
            SetTrayIcon();

            WmiManagementScope.Connect();
            WmiEventWatcher.Start();

            // Run the application.
            Application.Run();
        }
        private static void OpenHyperVManager()
        {
            var processInfo = new ProcessStartInfo
            {
                FileName = @$"{Environment.GetEnvironmentVariable("SYSTEMROOT")}\System32\mmc.exe",
                Arguments = _vmManagerPath!,
                WorkingDirectory = _hyperVInstallFolder,
                UseShellExecute = true,
                Verb = @"runas"
            };
            Process.Start(processInfo);
        }
        private static bool RequestVmShutdown(string virtualMachineName)
        {
            // Get the WMI Management Object for the virtual machine we are interested in.
            var virtualMachine = GetVirtualMachines(virtualMachineName).FirstOrDefault();

            // Check to see whether we found a matching virtual machine.
            if (virtualMachine != null)
            {
                // If we found a matching virtual machine get the `Msvm_ShutdownComponent` for it.
                var shutdowncomponent = virtualMachine.GetRelated("Msvm_ShutdownComponent").Cast<ManagementObject>().FirstOrDefault();

                // Check to see whether we found a Shutdown Component
                if (shutdowncomponent != null)
                {
                    // Get the parameters for the InitiateShudown method
                    var inParameters = shutdowncomponent.GetMethodParameters("InitiateShutdown");

                    // Set the 'Force' and 'Reason' parameters.
                    inParameters["Force"] = true;
                    inParameters["Reason"] = ApplicationName;

                    // Invoke the method, passing in the parameters we just set.
                    var outParameters = shutdowncomponent.InvokeMethod("InitiateShutdown", inParameters, null);

                    // Return the result of the method call.
                    if (outParameters != null)
                    {
                        return (StateChangeResponse)outParameters["ReturnValue"] is StateChangeResponse.CompletedwithNoError
                                                                                 or StateChangeResponse.MethodParametersCheckedTransitionStarted;
                    }
                }
            }

            // If we couldn't find a matching virtual machine then return `false`.
            return false;
        }
        private static bool RequestVmStateChange(string virtualMachineName, VmState state)
        {
            // Get the WMI Management Object for the virtual machine we are interested in.
            var virtualMachine = GetVirtualMachines(virtualMachineName).FirstOrDefault();

            // Check to see whether we found a matching virtual machine.
            if (virtualMachine != null)
            {
                // Get the parameters for the 'RequestStateChange' method.
                var inParameters = virtualMachine.GetMethodParameters("RequestStateChange");

                // Filter out the request as we only support a subset requesting a subset of all states.
                if (state is VmState.Enabled  // Running
                          or VmState.Disabled // Stopped
                          or VmState.Offline  // Saved
                          or VmState.Quiesce  // Paused
                          or VmState.Reset)
                {
                    // Set the 'RequestedState' parameter to the desired state.
                    inParameters["RequestedState"] = (ushort)state;

                    // Fire off the request to change the state.
                    var outParameters = virtualMachine.InvokeMethod("RequestStateChange", inParameters, null);

                    // Return the result of the method call.
                    if (outParameters != null)
                    {
                        return (StateChangeResponse)outParameters["ReturnValue"] is StateChangeResponse.CompletedwithNoError
                                                                                 or StateChangeResponse.MethodParametersCheckedTransitionStarted;
                    }
                }
            }

            // If we couldn't find a matching virtual machine then return `false`.
            return false;
        }
        private static void SetTrayIcon()
        {
            var iconWidth = PInvoke.GetTrayIconWidth(NotifyIcon.GetHandle());
            Debug.Assert(iconWidth > 0, "Icon width is 0");
            var iconSize = new Size(iconWidth, iconWidth);
            NotifyIcon.Icon = new Icon(ResourceHelper.Icon_HyperV, iconSize);

            Debug.WriteLine($"Icon size: {iconSize.Width}x{iconSize.Height}");
        }
        private static void ShowError(string heading, string text = "")
        {
            var taskDialogPage = new TaskDialogPage();
            taskDialogPage.AllowCancel = false;
            taskDialogPage.AllowMinimize = false;
            taskDialogPage.Buttons = [TaskDialogButton.Close];
            taskDialogPage.Caption = ApplicationName;
            taskDialogPage.Heading = heading;
            taskDialogPage.Icon = TaskDialogIcon.Error;
            taskDialogPage.Text = text;
            TaskDialog.ShowDialog(taskDialogPage);
        }
        private static void ShowToast(string virtualMachineName, VmState vmState, bool isCritical)
        {
            var status = VmStateToString(vmState);
            if (Environment.OSVersion.Version.Major >= 10)
            {
                var toast = new ToastContentBuilder().AddHeader(virtualMachineName, virtualMachineName, new ToastArguments());

                if (isCritical)
                { 
                    toast.AddText(ResourceHelper.Toast_CriticalState, AdaptiveTextStyle.Header);
                }

                toast.AddText(status)
                     .Show();
            }
            else
            {
                if (isCritical)
                {
                    status = $"{ResourceHelper.Toast_CriticalState}\n{status}";
                }
                NotifyIcon.ShowBalloonTip(BalloonTipTimeout, virtualMachineName, status, isCritical ? ToolTipIcon.Error : ToolTipIcon.Info);
            }
        }
        private static string VmStateToString(VmState state)
        {
            switch (state)
            {
                case VmState.Enabled:
                    return ResourceHelper.State_Running;
                case VmState.Disabled:
                    return ResourceHelper.State_Off;
                case VmState.Offline:
                case VmState.FastSaved:
                    return ResourceHelper.State_Saved;
                case VmState.Quiesce:
                    return ResourceHelper.State_Paused;
                default:
                    if ((int)state >= 32781 && (int)state <= 32792)
                    {
                        return ResourceHelper.State_Critical;
                    }

                    return state.ToString();
            }
        }

        #endregion

        #endregion
    }
}