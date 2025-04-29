using System;
using System.Runtime.InteropServices;

namespace HyperVTray
{
    internal class PInvoke
    {
        #region Methods

        #region Private Static

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetricsForDpi(int smIndex, uint dpi);

        #endregion

        #region Internal Static

        internal static int GetTrayIconWidth(IntPtr handle)
        {
            uint dpi = GetDpiForWindow(handle);
            return GetSystemMetricsForDpi(49, dpi); // SM_CXSMICON
        }

        #endregion

        #endregion
    }
}