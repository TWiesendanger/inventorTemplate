using System;
using System.Windows.Forms;

namespace InventorTemplate.Helper
{

    ///<summary> This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    /// This is primarily used for parenting a dialog to the Inventor window.
    /// This provides the expected behavior when the Inventor window is collapsed
    /// myForm.Show(New WindowWrapper(invApp.MainFrameHWND))
    /// </summary>
    public class WindowWrapper : IWin32Window
    {
        private IntPtr _hwnd;

        /// <summary>
        /// Creates a new instance of the window wrapper class
        /// </summary>
        /// <param name="windowHandle">Represents the window handle necessary to obtain the IWin32Window</param>
        public WindowWrapper(IntPtr windowHandle)
        {
            //Assign arguments
            _hwnd = windowHandle;
        }

        /// <summary>
        /// Creates a new instance of the window wrapper class
        /// </summary>
        /// <param name="windowHandle">Represents the window handle number necessary to obtain the IWin32Window</param>
        public WindowWrapper(int windowHandle)
        {
            //Assign arguments
            _hwnd = new IntPtr(windowHandle);
        }

        public IntPtr Handle
        {
            get { return _hwnd; }
        }
    }
}