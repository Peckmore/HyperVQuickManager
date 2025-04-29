namespace HyperVTray
{
    /// <summary>
    /// Represents the possible states a Hyper-V Virtual Machine can be in or requested to go to. This list is taken from both the v1 and v2 WMI providers, and consists of all states that can be set or returned.
    /// </summary>
    internal enum VmState : ushort
    {
        Unknown = 0, // The state of the element could not be determined.
        Other = 1,
        Enabled = 2, // The element is running.
        Disabled = 3, //  The element is turned off.
        ShutDown = 4, // The element is in the process of going to a Disabled state.
        NotApplicable = 5, // The element does not support being enabled or disabled.
        Offline = 6, // The element might be completing commands, and it will drop any new requests.
        Test = 7, //  The element is in a test state.
        Defer = 8, //  The element might be completing commands, but it will queue any new requests.
        Quiesce = 9, //  The element is enabled but in a restricted mode.The behavior of the element is similar to the Enabled state(2), but it processes only a restricted set of commands.All other requests are queued.
        RebootOrStarting = 10, // The element is in the process of going to an Enabled state(2). New requests are queued.
        Reset = 11, // Reset the virtual machine. Corresponds to CIM_EnabledLogicalElement.EnabledState = Reset.
        Paused = 32768, // VM is paused
        Suspended = 32769, // VM is in a saved state
        Starting = 32770, // VM is starting
        Saving = 32773, // In version 1 (V1) of Hyper-V, corresponds to EnabledStateSaving.
        Stopping = 32774, // VM is turning off
        Pausing = 32776, // In version 1 (V1) of Hyper-V, corresponds to EnabledStatePausing.
        Resuming = 32777, // In version 1 (V1) of Hyper-V, corresponds to EnabledStateResuming. State transition from Paused to Running.
        FastSaved = 32779, // Corresponds to EnabledStateFastSuspend.
        FastSaving = 32780, // Corresponds to EnabledStateFastSuspending. State transition from Running to FastSaved.

        // The following values represent critical states:
        RunningCritical = 32781,
        OffCritical = 32782,
        StoppingCritical = 32783,
        SavedCritical = 32784,
        PausedCritical = 32785,
        StartingCritical = 32786,
        ResetCritical = 32787,
        SavingCritical = 32788,
        PausingCritical = 32789,
        ResumingCritical = 32790,
        FastSavedCritical = 32791,
        FastSavingCritical = 32792,
    }
}