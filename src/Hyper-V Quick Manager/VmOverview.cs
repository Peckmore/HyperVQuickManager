using System;
using System.Runtime.Serialization;

namespace HyperVQuickManager
{
    /// <summary>
    /// A lightweight class used to describe the name and state of a Virtual Machine. The main purpose of the class is to send this data via WCF. The class also inherits EventArgs so it can be used in events.
    /// </summary>
    [DataContract]
    public class VmOverview : EventArgs
    {
        /// <summary>
        /// The name of the Virtual Machine.
        /// </summary>
        [DataMember]
        public string Name { get; set; }
        /// <summary>
        /// The state of the Virtual Machine.
        /// </summary>
        [DataMember]
        public VmState State { get; set; }
    }
}