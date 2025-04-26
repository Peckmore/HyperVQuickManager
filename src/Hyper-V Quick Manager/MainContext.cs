using HyperVQuickManager.Properties;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceModel;
using System.ServiceProcess;
using System.Threading;
using System.Windows.Forms;

namespace HyperVQuickManager
{
   [CallbackBehavior(ConcurrencyMode = ConcurrencyMode.Multiple)]
    public class MainContext : ApplicationContext, IVmManagerCallback
    {
        #region Fields

        private readonly Icon _baseIcon;
        private readonly IVmManager _hyperVManager;
        private readonly string _hyperVMscPath;
        private readonly string _mmcPath;
        private readonly ServiceController _serviceController;
        private readonly NotifyIcon _trayIcon;
        private readonly ContextMenu _trayIconContextMenu;
        private readonly string _vmConnectPath;

        #endregion

        #region Construction

        public MainContext()
        {
            if (Program.IsServiceInstalled())
            {
                // The application has been installed as a service so we want to use WCF
                // to communicate with it.
#if DEBUG
                //WindowsService.WindowsService.Start();
#endif

                var namedPipeFactory = new DuplexChannelFactory<IVmManager>(this, new NetNamedPipeBinding(), new EndpointAddress($"net.pipe://localhost/{Program.NamedPipeName}"));
               
                //ThreadPool.QueueUserWorkItem(obj => _hyperVManager = namedPipeFactory.CreateChannel());
                _hyperVManager = namedPipeFactory.CreateChannel();

                while (_hyperVManager == null)
                { Thread.Sleep(100); }

                _hyperVManager.Subscribe();
            }
            else
            {
                // There is no service installed so we are running in portable/standalone mode.


                    _hyperVManager = new VmManager();
                    ((VmManager)_hyperVManager).StateChanged += HyperVManager_VmStateChanged;
            }

            _hyperVMscPath = Path.Combine(Environment.SystemDirectory, "virtmgmt.msc");
            _mmcPath = Path.Combine(Environment.SystemDirectory, "mmc.exe");
            _vmConnectPath = "vmconnect.exe"; // Path.Combine(Environment.SystemDirectory, "vmconnect.exe");

            // Create the tray icon context menu
            _trayIconContextMenu = new ContextMenu();
            _trayIconContextMenu.Popup += TrayIconContextMenu_Opening;

            // Create the tray icon
            _trayIcon = new NotifyIcon { ContextMenu = _trayIconContextMenu };

            // Check the host OS to load the most appropriate tray icon.
            if (Environment.OSVersion.Version.Major == 6)
                _baseIcon = Environment.OSVersion.Version.Minor == 0 ? Resources.TrayIcon_WinVista : Resources.TrayIcon_Win7;
            else
                _baseIcon = Resources.TrayIcon_Win10;

            // Set the tray icon
            UpdateTrayIcon();

            _trayIcon.Click += TrayIcon_Click;
            _trayIcon.DoubleClick += TrayIcon_DoubleClick;

            // Show the tray icon
            _trayIcon.Visible = true;
        }

        #endregion

        #region Methods

        #region Event Handlers

        private void HyperVManager_VmStateChanged(object sender, VmOverview e) => StateChanged(e);
        private void MenuItemConnect_Click(object sender, EventArgs e) => ConnectToVm(((MenuItem)sender).Parent.Name);
        private void MenuItemConnectAndStart_Click(object sender, EventArgs e)
        {
            var name = ((MenuItem)sender).Parent.Name;
            InvokeSingleChangeState(StartVm, name);
            ConnectToVm(name);
        }
        private void MenuItemExit_Click(object sender, EventArgs e) => Application.Exit();
        private void MenuItemManager_Click(object sender, EventArgs e) => OpenHyperVManager();
        private void MenuItemPause_Click(object sender, EventArgs e) => InvokeSingleChangeState(PauseVm, ((MenuItem)sender).Parent.Name);
        private void MenuItemPauseAll_Click(object sender, EventArgs e) => InvokeMultipleChangeState(PauseVm, _hyperVManager.GetVm().Where(vm => IsRunning(vm.State)));
        private void MenuItemReset_Click(object sender, EventArgs e) => InvokeSingleChangeState(ResetVm, ((MenuItem)sender).Parent.Name);
        private void MenuItemResetAll_Click(object sender, EventArgs e) => InvokeMultipleChangeState(ResetVm, _hyperVManager.GetVm().Where(vm => IsRunning(vm.State) || IsPaused(vm.State)));
        private void MenuItemResumeAll_Click(object sender, EventArgs e) => InvokeMultipleChangeState(StartVm, _hyperVManager.GetVm().Where(vm => IsPaused(vm.State)));
        private void MenuItemSave_Click(object sender, EventArgs e) => InvokeSingleChangeState(SaveVm, ((MenuItem)sender).Parent.Name);
        private void MenuItemSaveAll_Click(object sender, EventArgs e) => InvokeMultipleChangeState(SaveVm, _hyperVManager.GetVm().Where(vm => IsRunning(vm.State) || IsPaused(vm.State)));
        private void MenuItemShutdown_Click(object sender, EventArgs e) => InvokeSingleChangeState(_hyperVManager.ShutdownVm, ((MenuItem)sender).Parent.Name);
        private void MenuItemShutdownAll_Click(object sender, EventArgs e) => InvokeMultipleChangeState(_hyperVManager.ShutdownVm, _hyperVManager.GetVm().Where(vm => IsRunning(vm.State)));
        private void MenuItemStart_Click(object sender, EventArgs e) => InvokeSingleChangeState(StartVm, ((MenuItem)sender).Parent.Name);
        private void MenuItemStartAll_Click(object sender, EventArgs e) => InvokeMultipleChangeState(StartVm, _hyperVManager.GetVm().Where(vm => IsOff(vm.State) || IsSaved(vm.State)));
        private void MenuItemStop_Click(object sender, EventArgs e) => InvokeSingleChangeState(TurnOffVm, ((MenuItem)sender).Parent.Name);
        private void MenuItemStopAll_Click(object sender, EventArgs e) => InvokeMultipleChangeState(TurnOffVm, _hyperVManager.GetVm().Where(vm => IsRunning(vm.State) || IsPaused(vm.State)));
        private void TrayIcon_Click(object sender, EventArgs e) => typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic)?.Invoke(_trayIcon, null);
        private void TrayIcon_DoubleClick(object sender, EventArgs e) => OpenHyperVManager();
        private void TrayIconContextMenu_Opening(object sender, EventArgs e) => PopulateMenu();

        #endregion

        #region Private

        private void ConnectToVm(string name) => Process.Start(new ProcessStartInfo(_vmConnectPath, $"localhost {name}") { Verb = "runas" });

        [SuppressMessage("ReSharper", "MemberCanBeMadeStatic.Local")]
        [SuppressMessage("ReSharper", "SwitchStatementMissingSomeCases")]
        private string GetRequestStateChangeResponseString(StateChangeResponse response)
        {
            switch (response)
            {
                case StateChangeResponse.CompletedwithNoError:
                    return Strings.StateChangeResponse_CompletedwithNoError;
                case StateChangeResponse.MethodParametersCheckedTransitionStarted:
                    return Strings.StateChangeResponse_MethodParametersCheckedTransitionStarted;
                case StateChangeResponse.Failed:
                    return Strings.StateChangeResponse_Failed;
                case StateChangeResponse.AccessDenied:
                    return Strings.StateChangeResponse_AccessDenied;
                case StateChangeResponse.NotSupported:
                    return Strings.StateChangeResponse_NotSupported;
                case StateChangeResponse.StatusIsUnknown:
                    return Strings.StateChangeResponse_StatusIsUnknown;
                case StateChangeResponse.Timeout:
                    return Strings.StateChangeResponse_Timeout;
                case StateChangeResponse.InvalidParameter:
                    return Strings.StateChangeResponse_InvalidParameter;
                case StateChangeResponse.SystemIsInUse:
                    return Strings.StateChangeResponse_SystemIsInUse;
                case StateChangeResponse.InvalidStateForThisOperation:
                    return Strings.StateChangeResponse_InvalidStateForThisOperation;
                case StateChangeResponse.IncorrectDataType:
                    return Strings.StateChangeResponse_IncorrectDataType;
                case StateChangeResponse.SystemIsNotAvailable:
                    return Strings.StateChangeResponse_SystemIsNotAvailable;
                case StateChangeResponse.OutOfMemory:
                    return Strings.StateChangeResponse_OutOfMemory;
                //case RequestStateChangeResponse.DmtfReserved:
                default:
                    return Strings.StateChangeResponse_UnknownError;
            }
        }
        private string GetVmStateString(VmState state)
        {
            // We use a switch to convert the VmState enum value into a displayable string.
            switch (state)
            {
                case VmState.Enabled:
                case VmState.RunningCritical:
                    return Strings.State_Running;
                case VmState.Disabled:
                case VmState.OffCritical:
                    return Strings.State_Stopped;
                case VmState.Paused:
                case VmState.Quiesce:
                case VmState.PausedCritical:
                    return Strings.State_Paused;
                case VmState.Suspended:
                case VmState.Offline:
                case VmState.SavedCritical:
                case VmState.FastSaved:
                case VmState.FastSavedCritical:
                    return Strings.State_Saved;
                case VmState.Starting:
                case VmState.StartingCritical:
                case VmState.RebootOrStarting:
                    return Strings.State_Starting;
                case VmState.Saving:
                case VmState.SavingCritical:
                case VmState.FastSaving:
                case VmState.FastSavingCritical:
                    return Strings.State_Saving;
                case VmState.Stopping:
                case VmState.StoppingCritical:
                case VmState.ShutDown:
                    return Strings.State_Stopping;
                case VmState.Pausing:
                case VmState.PausingCritical:
                    return Strings.State_Pausing;
                case VmState.Resuming:
                case VmState.ResumingCritical:
                    return Strings.State_Resuming;
                case VmState.Reset:
                case VmState.ResetCritical:
                    return Strings.State_Resetting;
                //case VmState.Unknown:
                //case VmState.Other:
                //case VmState.NotApplicable:
                //case VmState.Test:
                //case VmState.Defer:
                default:
                    return Strings.State_Unknown;
            }
        }
        private void InvokeMultipleChangeState(Func<string, StateChangeResponse> function, IEnumerable<VmOverview> vms)
        {
            var failedCount = vms.Count(vm =>
            {
                var response = function(vm.Name);
                return response != StateChangeResponse.CompletedwithNoError && response != StateChangeResponse.MethodParametersCheckedTransitionStarted;
            });

            if (failedCount > 0)
                Program.ShowDialog(string.Format(Strings.Error_ChangeState_TextMultiple, failedCount), Strings.Error_ChangeState_InstructionTextMultiple, TaskDialogStandardButtons.Close, TaskDialogStandardIcon.Error);
        }
        private void InvokeSingleChangeState(Func<string, StateChangeResponse> function, string name)
        {
            var response = function(name);
            if (response != StateChangeResponse.CompletedwithNoError && response != StateChangeResponse.MethodParametersCheckedTransitionStarted)
                Program.ShowDialog($"{GetRequestStateChangeResponseString(response)}.", string.Format(Strings.Error_ChangeState_InstructionText, name), TaskDialogStandardButtons.Close, TaskDialogStandardIcon.Error);
        }
        [SuppressMessage("ReSharper", "MemberCanBeMadeStatic.Local")]
        private bool IsCritical(VmState state)
        {
            // We use a switch to determine if the VmState enum value is for a
            // critical state.
            switch (state)
            {
                case VmState.RunningCritical:
                case VmState.OffCritical:
                case VmState.PausedCritical:
                case VmState.SavedCritical:
                case VmState.FastSavedCritical:
                case VmState.StartingCritical:
                case VmState.SavingCritical:
                case VmState.FastSavingCritical:
                case VmState.StoppingCritical:
                case VmState.PausingCritical:
                case VmState.ResumingCritical:
                case VmState.ResetCritical:
                    return true;
                default:
                    return false;
            }
        }
        private bool IsOff(VmState state)
        {
            // We use a switch to determine if the VmState enum value is for an
            // off state.
            switch (state)
            {
                case VmState.Disabled:
                case VmState.OffCritical:
                    return true;
                default:
                    return false;
            }
        }
        private bool IsPaused(VmState state)
        {
            // We use a switch to determine if the VmState enum value is for a
            // paused state.
            switch (state)
            {
                case VmState.Paused:
                case VmState.Quiesce:
                case VmState.PausedCritical:
                    return true;
                default:
                    return false;
            }
        }
        private bool IsRunning(VmState state)
        {
            // We use a switch to determine if the VmState enum value is for a
            // running state.
            switch (state)
            {
                case VmState.Enabled:
                case VmState.RunningCritical:
                    return true;
                default:
                    return false;
            }
        }
        private bool IsSaved(VmState state)
        {
            // We use a switch to determine if the VmState enum value is for a
            // saved state.
            switch (state)
            {
                case VmState.Suspended:
                case VmState.Offline:
                case VmState.SavedCritical:
                case VmState.FastSaved:
                case VmState.FastSavedCritical:
                    return true;
                default:
                    return false;
            }
        }
        private bool IsUnknownState(VmState state)
        {
            // We use a switch to determine if the VmState enum value is for a
            // saved state.
            switch (state)
            {
                case VmState.Unknown:
                case VmState.Other:
                case VmState.NotApplicable:
                case VmState.Test:
                case VmState.Defer:
                    return true;
                default:
                    return false;
            }
        }
        private void OpenHyperVManager() => Process.Start(_mmcPath, _hyperVMscPath);
        private StateChangeResponse PauseVm(string name) => _hyperVManager.RequestVmStateChange(name, Program.UseV2Api ? VmState.Quiesce : VmState.Paused);
        private void PopulateMenu()
        {
            // Get a list of the virtual machines running on this computer.
            var vms = _hyperVManager.GetVm();

            // Set a flag to indicate whether VMConnect is on disk.
            var enableConnect = File.Exists(Path.Combine(Environment.SystemDirectory, _vmConnectPath));

            // Set a flag to indicate whether Hyper-V Manager is on disk.
            var enableManager = File.Exists(_hyperVMscPath);

            // Set a flag to indicate whether this computer has any virtual machines.
            var hasVms = vms.Any();

            // Clear the menu so we can populate it from scratch.
            _trayIconContextMenu.MenuItems.Clear();

            // Iterate through the VMs configured on the localhost, and add the
            // appropriate menu items based on each VM's state.
            foreach (var vm in vms)
            {
                var name = vm.Name;
                var state = vm.State;

                // Create a root menu item for the VM.
                var entry = $"{name} ({GetVmStateString(state)})";
                if (IsCritical(state))
                    entry += $" [{Strings.State_Critical.ToUpperInvariant()}]";
                var vmItem = new MenuItem(entry) { Name = name };

                // If VMConnect is on disk add a sub-menu option to connect to the VM.
                if (enableConnect)
                {
                    vmItem.MenuItems.Add(new MenuItem(Strings.Menu_Connect, MenuItemConnect_Click));

                    if (IsOff(state) || IsSaved(state))
                        vmItem.MenuItems.Add(new MenuItem(Strings.Menu_ConnectAndStart, MenuItemConnectAndStart_Click));
                    else if (IsPaused(state))
                        vmItem.MenuItems.Add(new MenuItem(Strings.Menu_ConnectAndResume, MenuItemConnectAndStart_Click));
                }

                var subItems = new List<MenuItem>();

                if (IsOff(state) || IsSaved(state))
                    subItems.Add(new MenuItem(Strings.Menu_Start, MenuItemStart_Click));

                if (IsRunning(state) || IsPaused(state))
                    subItems.Add(new MenuItem(Strings.Menu_TurnOff, MenuItemStop_Click));

                if (IsRunning(state))
                    subItems.Add(new MenuItem(Strings.Menu_ShutDown, MenuItemShutdown_Click));

                if (IsRunning(state) || IsPaused(state))
                    subItems.Add(new MenuItem(Strings.Menu_Save, MenuItemSave_Click));

                if (subItems.Any() && IsRunning(state) || IsPaused(state))
                    subItems.Add(new MenuItem("-"));

                if (IsRunning(state))
                    subItems.Add(new MenuItem(Strings.Menu_Pause, MenuItemPause_Click));

                if (IsPaused(state))
                    subItems.Add(new MenuItem(Strings.Menu_Resume, MenuItemStart_Click));

                if (IsRunning(state) || IsPaused(state))
                    subItems.Add(new MenuItem(Strings.Menu_Reset, MenuItemReset_Click));

                // If we have added the 'Connect' command, and we know we can now
                // add other commands to control the state of the VM then we need
                // to add a seperator between the items (for consistency with
                // Hyper-V Manager context menus).
                if (enableConnect && subItems.Any())
                    vmItem.MenuItems.Add(new MenuItem("-"));

                vmItem.MenuItems.AddRange(subItems.ToArray());

                // If we have not added any sub-items to the VM menu
                if (vmItem.MenuItems.Count == 0)
                    vmItem.Enabled = false;

                // Add the VM to the context menu.
                _trayIconContextMenu.MenuItems.Add(vmItem);
            }

            if (hasVms)
            {
                _trayIconContextMenu.MenuItems.Add(new MenuItem("-"));

                // Create a root menu item for the VM.
                var vmItem = new MenuItem(Strings.Menu_AllVirtualMachines);

                var subItems = new List<MenuItem>();

                var isOff = vms.Any(vm => IsOff(vm.State));
                var isPaused = vms.Any(vm => IsPaused(vm.State));
                var isRunning = vms.Any(vm => IsRunning(vm.State));
                var isSaved = vms.Any(vm => IsSaved(vm.State));

                if (isOff || isSaved)
                    subItems.Add(new MenuItem(Strings.Menu_Start, MenuItemStartAll_Click));

                if (isRunning || isPaused)
                    subItems.Add(new MenuItem(Strings.Menu_TurnOff, MenuItemStopAll_Click));

                if (isRunning)
                    subItems.Add(new MenuItem(Strings.Menu_ShutDown, MenuItemShutdownAll_Click));

                if (isRunning || isPaused)
                    subItems.Add(new MenuItem(Strings.Menu_Save, MenuItemSaveAll_Click));

                if (subItems.Any() && isRunning || isPaused)
                    subItems.Add(new MenuItem("-"));

                if (isRunning)
                    subItems.Add(new MenuItem(Strings.Menu_Pause, MenuItemPauseAll_Click));

                if (isPaused)
                    subItems.Add(new MenuItem(Strings.Menu_Resume, MenuItemResumeAll_Click));

                if (isRunning || isPaused)
                    subItems.Add(new MenuItem(Strings.Menu_Reset, MenuItemResetAll_Click));

                vmItem.MenuItems.AddRange(subItems.ToArray());

                // Add the VM to the context menu.
                _trayIconContextMenu.MenuItems.Add(vmItem);
            }

            // Check if the Hyper-V MMC plugin is on disk to determine whether we
            // can show the 'Hyper-V Manager' menu item.
            if (enableManager)
            {
                if (hasVms)
                    _trayIconContextMenu.MenuItems.Add(new MenuItem("-"));
                _trayIconContextMenu.MenuItems.Add(new MenuItem(Strings.Menu_HyperVManager, MenuItemManager_Click));
            }

            // Finally we add an 'Exit' menu item to allow people to close the application.
            if (hasVms || enableManager)
                _trayIconContextMenu.MenuItems.Add(new MenuItem("-"));
            _trayIconContextMenu.MenuItems.Add(new MenuItem(Strings.Menu_Exit, MenuItemExit_Click));

            UpdateTrayIcon();
        }
        private StateChangeResponse ResetVm(string name) => _hyperVManager.RequestVmStateChange(name, VmState.Reset);
        private StateChangeResponse SaveVm(string name) => _hyperVManager.RequestVmStateChange(name, Program.UseV2Api ? VmState.Offline : VmState.Suspended);
        private StateChangeResponse StartVm(string name) => _hyperVManager.RequestVmStateChange(name, VmState.Enabled);
        public void StateChanged(VmOverview overview)
        {
            UpdateTrayIcon();

            if (IsUnknownState(overview.State))
                return;

            var stateString = GetVmStateString(overview.State);

            var message = string.Format(Strings.Balloon_Message, overview.Name, stateString.ToLowerInvariant());
            var toolTipIcon = ToolTipIcon.Info;
            if (IsCritical(overview.State))
            {
                message += Strings.Balloon_MessageExtensionCritical;
                toolTipIcon = ToolTipIcon.Error;
            }

            message += ".";

            _trayIcon.ShowBalloonTip(0, string.Format(Strings.Balloon_Title, stateString), message, toolTipIcon);
        }
        private StateChangeResponse TurnOffVm(string name) => _hyperVManager.RequestVmStateChange(name, VmState.Disabled);
        private void UpdateTrayIcon()
        {
            var baseIcon = new Icon(_baseIcon, SystemInformation.SmallIconSize);
            Icon iconOverlay = null;

            IList<VmOverview> vms = null;
            try
            {
                // Get a list of the virtual machines running on this computer.
                vms = _hyperVManager.GetVm();
            }catch (Exception ex)
            { }

            if (vms.Any(vm => IsCritical(vm.State)))
                iconOverlay = new Icon(Resources.TrayIconOverlay_Critical, SystemInformation.SmallIconSize);
            else if (vms.Any(vm => IsRunning(vm.State)))
                iconOverlay = new Icon(Resources.TrayIconOverlay_Running, SystemInformation.SmallIconSize);
            else if (vms.Any(vm => IsPaused(vm.State)))
                iconOverlay = new Icon(Resources.TrayIconOverlay_Paused, SystemInformation.SmallIconSize);

            var bitmap = new Bitmap(SystemInformation.SmallIconSize.Width, SystemInformation.SmallIconSize.Height);
            var canvas = Graphics.FromImage(bitmap);
            canvas.DrawImage(baseIcon.ToBitmap(), new Point(0, 0));
            if (iconOverlay != null)
                canvas.DrawImage(iconOverlay.ToBitmap(), new Point(0, 0));
            canvas.Save();

            // Get icon handle from bitmap
            var iconHandle = bitmap.GetHicon();

            // Create a new icon from the handle
            var icon = Icon.FromHandle(iconHandle);

            try
            {
                // Use a clone of the icon as the icon we create from the handle
                // doesn't take ownership of said handle and so won't destroy it when
                // the object is disposed, meaning we have a memory leak as the handle
                // is never cleaned up. The new icon created using clone will take
                // ownership of it's own icon handle and so clean it up properly.
                _trayIcon.Icon = icon.Clone() as Icon;
            }
            finally
            {
                // We now manually clean up the handle we created from the bitmap
                // to ensure we don't have a memory leak.
                NativeMethods.DestroyIcon(iconHandle);

                // Dispose of the icon oject
                icon.Dispose();
            }
        }

        #endregion

        #region Protected

        protected override void Dispose(bool disposing)
        {
            // Call the 'Dispose' method on the base class.
            base.Dispose(disposing);

            if (disposing)
            {
                
            }
        }

        #endregion

        #endregion
    }
}