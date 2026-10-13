using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OutlookAI.QuickTools
{
    /// <summary>An Outlook Explorer or Inspector as the owner of our windows.</summary>
    internal sealed class OwnerWindow : IWin32Window
    {
        [ComImport]
        [Guid("00000114-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleWindow
        {
            void GetWindow(out IntPtr phwnd);

            void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
        }

        private OwnerWindow(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; }

        public static IWin32Window For(object outlookWindow)
        {
            var ole = outlookWindow as IOleWindow;
            if (ole == null) return null;
            try
            {
                IntPtr handle;
                ole.GetWindow(out handle);
                return handle == IntPtr.Zero ? null : new OwnerWindow(handle);
            }
            catch (COMException)
            {
                return null;
            }
        }
    }
}
