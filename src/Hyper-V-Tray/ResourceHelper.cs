using System;
using System.Collections.Specialized;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Resources;

namespace HyperVTray
{
    internal static class ResourceHelper
    {
        #region Constants

        private static readonly StringDictionary StringsDictionary;

        #endregion

        #region Fields

        private static ResourceManager? _vmBrowserResourceManager;

        #endregion

        #region Construction

        static ResourceHelper()
        {
            StringsDictionary = new();
            Icon_HyperV = GetIconResource(Resources.Icon_HyperV);
        }

        #endregion

        #region Properties

        internal static string Command_Connect => GetStringResource("VMOpen_Name", Resources.Command_Connect);
        internal static string Command_Pause => GetStringResource("VMPause_Name", Resources.Command_Pause);
        internal static string Command_Reset => GetStringResource("VMReset_Name", Resources.Command_Reset);
        internal static string Command_Resume => GetStringResource("VMResume_Name", Resources.Command_Resume);
        internal static string Command_Save => GetStringResource("VMSaveState_Name", Resources.Command_Save);
        internal static string Command_ShutDown => GetStringResource("VMShutDown_Name", Resources.Command_ShutDown);
        internal static string Command_Start => GetStringResource("VMStart_Name", Resources.Command_Start);
        internal static string Command_TurnOff => GetStringResource("VMTurnOff_Name", Resources.Command_TurnOff);
        internal static Icon Icon_HyperV { get; }
        internal static string Message_PauseVMFailed => GetStringResource("Message_PauseVMFailed", Resources.Message_PauseVMFailed);
        internal static string Message_PowerOffVMFailed => GetStringResource("Message_PowerOffVMFailed", Resources.Message_PowerOffVMFailed);
        internal static string Message_ResetVMFailed => GetStringResource("Message_ResetVMFailed", Resources.Message_ResetVMFailed);
        internal static string Message_ResumeVMFailed => GetStringResource("Message_ResumeVMFailed", Resources.Message_ResumeVMFailed);
        internal static string Message_SaveStateVMFailed => GetStringResource("Message_SaveStateVMFailed", Resources.Message_SaveStateVMFailed);
        internal static string Message_ShutDownVMFailed => GetStringResource("Message_ShutDownVMFailed", Resources.Message_ShutDownVMFailed);
        internal static string Message_StartVMFailed => GetStringResource("Message_StartVMFailed", Resources.Message_StartVMFailed);
        internal static string State_Critical => GetStringResource("cccccccc", Resources.State_Critical);
        internal static string State_Off => GetStringResource("fffffff", Resources.State_Off);
        internal static string State_Paused => GetStringResource("DDDDDD", Resources.State_Paused);
        internal static string State_Running => GetStringResource("rrrrrrr", Resources.State_Running);
        internal static string State_Saved => GetStringResource("dddddddd", Resources.State_Saved);
        internal static string Toast_CriticalState => Resources.Toast_CriticalState;
        
        #endregion

        #region Methods

        #region Private Static

        private static Icon GetIconResource(byte[] bytes)
        {
            using (var ms = new MemoryStream(bytes))
            {
                return new Icon(ms);
            }
        }
        private static string GetStringResource(string resourceName, string fallbackValue)
        {
            if (!StringsDictionary.ContainsKey(resourceName))
            {
                StringsDictionary[resourceName] = _vmBrowserResourceManager?.GetString(resourceName) ?? fallbackValue;
            }

            return StringsDictionary[resourceName]!;
        }

        #endregion

        #region Internal Static

        internal static void Initialize(string hyperVInstallFolder)
        {
            // Handle the AssemblyResolve event to manually load missing assemblies
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                var assemblyName = new AssemblyName(args.Name);
                var assemblyPath = Path.Combine(hyperVInstallFolder, $"{assemblyName.Name}.dll");

                if (File.Exists(assemblyPath))
                {
                    return Assembly.LoadFile(assemblyPath);
                }

                return null; // Return null if the assembly cannot be resolved
            };

            var vmBrowserAssembly = Assembly.LoadFile(Path.Combine(hyperVInstallFolder, "Microsoft.Virtualization.Client.VMBrowser.dll"));
            _vmBrowserResourceManager = new ResourceManager(@"Microsoft.Virtualization.Client.VMBrowser.Resources", vmBrowserAssembly);
        }

        #endregion

        #endregion
    }
}