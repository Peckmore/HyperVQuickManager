using System.Diagnostics.CodeAnalysis;

namespace HyperVQuickManager
{
    /// <summary>
    /// Represents the possible responses from requesting a Virtual Machine to change state.
    /// </summary>
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    public enum StateChangeResponse : uint
    {
        CompletedwithNoError = 0,
        MethodParametersCheckedTransitionStarted = 4096,
        Failed = 32768,
        AccessDenied = 32769,
        NotSupported = 32770,
        StatusIsUnknown = 32771,
        Timeout = 32772,
        InvalidParameter = 32773,
        SystemIsInUse = 32774,
        InvalidStateForThisOperation = 32775,
        IncorrectDataType = 32776,
        SystemIsNotAvailable = 32777,
        OutOfMemory = 32778,
        DmtfReserved = 74095,
    }
}