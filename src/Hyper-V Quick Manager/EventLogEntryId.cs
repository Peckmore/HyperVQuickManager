namespace HyperVQuickManager
{
    /// <summary>
    /// Entry IDs for the possible event log entries the Windows Service could make.
    /// </summary>
    internal enum EventLogEntryId
    {
        /// <summary>
        /// The OS the user is attempting to start the service on is unsupported.
        /// </summary>
        UnsupportedOs = 1,
        /// <summary>
        /// The service must be run under an account that has administrative rights.
        /// </summary>
        RequiresAdmin = 2,
        /// <summary>
        /// There was an error and the service failed to start.
        /// </summary>
        FailedToStart = 3
    }
}