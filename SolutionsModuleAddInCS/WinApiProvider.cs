using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolutionsModuleAddInCS
{
    /// <summary>
    /// This class encapsulates all P/Invoke unmanaged functions.
    /// </summary>
    [SuppressUnmanagedCodeSecurity]
    class WinApiProvider
    {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        public const int SW_HIDE = 0;

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, ref Rect1 rectangle);

        [DllImport("user32.dll")]
        public static extern bool OffsetRect(ref Rect lpRect, int dx, int dy);

        public struct Rect1
        {
            public int Left { get; set; }
            public int Top { get; set; }
            public int Right { get; set; }
            public int Bottom { get; set; }
        }

        [DllImport("user32", CharSet = CharSet.Auto)]
        public static extern IntPtr GetActiveWindow();
        /// <summary>
        /// The <b>FindWindow</b> method finds a window by it's classname and caption. 
        /// </summary>
        /// <param name="lpClassName">The classname of the window (use Spy++)</param>
        /// <param name="lpWindowName">The Caption of the window.</param>
        /// <returns>Returns a valid window handle or 0.</returns>
        [DllImport("user32", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        /// <summary>
        /// Retrieves the WindowTest of the window given by the handle.
        /// </summary>
        /// <param name="hWnd">The windows handle</param>
        /// <param name="lpString">A stringbuilder object wich receives the window text</param>
        /// <param name="nMaxCount">The max length of the text to retrieve, usually 260</param>
        /// <returns>Returns the length of chars received.</returns>
        [DllImport("user32", CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        /// <summary>
        /// Retrieves the ClassName of the window given by the handle.
        /// </summary>
        /// <param name="hWnd">The windows handle</param>
        /// <param name="lpString">A stringbuilder object wich receives the window text</param>
        /// <param name="nMaxCount">The max length of the text to retrieve, usually 260</param>
        /// <returns>Returns the length of chars received.</returns>
        [DllImport("user32", CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        public static string GetWindowText(IntPtr hWnd)
        {
            StringBuilder windowName = new StringBuilder(260);
            int textLen = GetWindowText(hWnd, windowName, 260);
            return windowName.ToString();
        }

        public static string GetClassName(IntPtr hWnd)
        {
            StringBuilder className = new StringBuilder(260);
            int textLen = GetClassName(hWnd, className, 260);
            return className.ToString();
        }

        /// <summary>
        /// Returns a list of windowtext of the given list of window handles..
        /// </summary>
        /// <param name="windowHandles">A list of window handles.</param>
        /// <returns>Returns a list with the corresponding window text for each window.</returns>
        public static List<string> GetWindowNames(List<IntPtr> windowHandles)
        {
            List<string> windowNameList = new List<string>();

            // A Stringbuilder will receive our windownames...
            StringBuilder windowName = new StringBuilder(260);
            foreach (IntPtr hWnd in windowHandles)
            {
                int textLen = GetWindowText(hWnd, windowName, 260);

                // get the windowtext
                windowNameList.Add(windowName.ToString());
            }
            return windowNameList;
        }

        /// <summary>
        /// Returns a list of windowtext of the given list of window handles..
        /// </summary>
        /// <param name="windowHandles">A list of window handles.</param>
        /// <returns>Returns a list with the corresponding window text for each window.</returns>
        public static List<string> GetClassNames(List<IntPtr> windowHandles)
        {
            List<string> classNameList = new List<string>();

            // A Stringbuilder will receive our windownames...
            StringBuilder className = new StringBuilder(260);
            foreach (IntPtr hWnd in windowHandles)
            {
                int textLen = GetClassName(hWnd, className, 260);

                // get the windowtext
                classNameList.Add(className.ToString());
            }
            return classNameList;
        }

        public static int FindChildByClassName(List<IntPtr> windowHandles, string className)
        {
            int index = -1;
            StringBuilder windowName = new StringBuilder(260);
            foreach (IntPtr hWnd in windowHandles)
            {
                int textLen = GetClassName(hWnd, windowName, 260);

                if (windowName.ToString() == className)
                {
                    index = windowHandles.IndexOf(hWnd);
                    break;
                }
            }
            return index;
        }

        public static IntPtr FindParentByClassName(IntPtr hWndChild, string className)
        {
            IntPtr windowHandle = hWndChild;
            string windowClassName = null;
            do
            {
                windowHandle = GetParent(windowHandle);
                if (windowHandle != IntPtr.Zero)
                {
                    StringBuilder sb = new StringBuilder(260);
                    GetClassName(windowHandle, sb, 260);
                    windowClassName = sb.ToString();
                }
            }
            while (windowHandle != IntPtr.Zero && windowClassName != className);
            return windowHandle;
        }

        /// <summary>
        /// Returns a list of all child window handles for the given window handle.
        /// </summary>
        /// <param name="hParentWnd">Handle of the parent window.</param>
        /// <returns>A list of all child window handles recursively.</returns>
        public static List<IntPtr> EnumChildWindows(IntPtr hParentWnd)
        {
            // The list will hold all child handles. 
            List<IntPtr> childWindowHandles = new List<IntPtr>();

            // We will allocate an unmanaged handle and pass a pointer to the EnumWindow method.
            GCHandle hChilds = GCHandle.Alloc(childWindowHandles);
            try
            {
                // Define the callback method.
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                // Call the unmanaged function to enum all child windows
                EnumChildWindows(hParentWnd, childProc, GCHandle.ToIntPtr(hChilds));
            }
            finally
            {
                // Free unmanaged resources.
                if (hChilds.IsAllocated)
                    hChilds.Free();
            }

            return childWindowHandles;
        }

        public static List<IntPtr> EnumChildWindows2(IntPtr hParentWnd)
        {
            List<IntPtr> childWindows = EnumChildWindows(hParentWnd);
            for (int i = 0; i < childWindows.Count;)
            {
                if (GetParent(childWindows[i]) != hParentWnd)
                    childWindows.RemoveAt(i);
                else
                    i++;
            }
            return childWindows;
        }

        /// <summary>
        /// A method to enummerate all childwindows of the given windows handle.
        /// </summary>
        /// <param name="hWnd">The parent window handle.</param>
        /// <param name="callback">The callback method wich is called for each childwindow.</param>
        /// <param name="userObject">A pointer to a userdefined object, e.g a list.</param>
        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr hWnd, EnumWindowProc callback, IntPtr userObject);

        /// <summary>
        /// Callback method to be used when enumerating windows.
        /// </summary>
        /// <param name="hChildWindow">Handle of the next window</param>
        /// <param name="pointer">Pointer to a GCHandle that holds a reference to the dictionary for our windowHandles.</param>
        /// <returns>True to continue the enumeration, false to bail</returns>
        private static bool EnumWindow(IntPtr hChildWindow, IntPtr pointer)
        {
            GCHandle hChilds = GCHandle.FromIntPtr(pointer);
            ((List<IntPtr>)hChilds.Target).Add(hChildWindow);

            return true;
        }

        /// <summary>
        /// Delegate for the EnumChildWindows method
        /// </summary>
        /// <param name="hWnd">Window handle</param>
        /// <param name="parameter">Caller-defined variable</param>
        /// <returns>True to continue enumerating, false to exit the search.</returns>
        public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        /// <summary>
        /// Sends a message command to the give window address.
        /// </summary>
        /// <param name="hWnd">handle to destination window</param>
        /// <param name="Msg">message</param>
        /// <param name="wParam">first message parameter</param>
        /// <param name="lParam">second message parameter</param>
        /// <returns></returns>
        [DllImport("user32")]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        [DllImport("user32")]
        public static extern int PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        [DllImport("user32")]
        public static extern int GetDlgCtrlID(IntPtr hWnd);

        /// <summary>
        /// Constant defines a System command message
        /// </summary>
        public const uint WM_SYSCOMMAND = 0x0112;

        public const uint WM_NULL = 0x0000;
        public const uint WM_DESTROY = 0x0002;
        public const uint WM_SIZE = 0x0005;
        public const uint WM_ACTIVATE = 0x0006;
        public const uint WM_SETFOCUS = 0x0007;
        public const uint WM_KILLFOCUS = 0x0008;
        public const uint WM_CLOSE = 0x0010;
        public const uint WM_CONTEXTMENU = 0x007B;
        public const uint WM_KEYDOWN = 0x0100;
        public const uint WM_KEYUP = 0x0101;

        /// <summary>
        /// Defines the Windows Close command
        /// </summary>
        public const int SC_CLOSE = 0xF060;

        public const int GWL_WNDPROC = -4;
        public const int GWL_STYLE = -16;

        /// <summary>
        /// Set a new parent for the given window handle
        /// </summary>
        /// <param name="hWndChild">The handle of the target window</param>
        /// <param name="hWndNewParent">The window handle of the parent window</param>
        [DllImport("user32", SetLastError = true)]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32")]
        public static extern IntPtr GetParent(IntPtr hWndChild);

        [DllImport("user32", EntryPoint = "GetProp")]
        public static extern int GetPropA(int hwnd, string lpString);

        [DllImport("user32")]
        public static extern int RemoveProp(int hwnd, string lpString);

        [DllImport("user32")]
        public static extern uint GetDpiForWindow(IntPtr hWnd);

        /// <summary>
        /// Create a new window.
        /// Description see http://msdn2.microsoft.com/en-us/library/ms632680.aspx
        /// </summary>
        /// <param name="dwExStyle">Specifies the extended window style of the window being created</param>
        /// <param name="lpClassName">A class name - see http://msdn2.microsoft.com/en-us/library/ms633574.aspx</param>
        /// <param name="lpWindowName">Pointer to a null-terminated string that specifies the window name</param>
        /// <param name="dwStyle">Specifies the style of the window being created</param>
        /// <param name="x">The window startposition X</param>
        /// <param name="y">The window startposition Y</param>
        /// <param name="nWidth">Width</param>
        /// <param name="nHeight">Height</param>
        /// <param name="hWndParent">Parent window handle</param>
        /// <param name="hMenu">Handle to a menu</param>
        /// <param name="hInstance">Handle to the instance of the module to be associated with the window</param>
        /// <param name="lpParam">Pointer to a value to be passed to the window through the CREATESTRUCT structure </param>
        /// <returns>If the function succeeds, the return value is a handle to the new window</returns>
        [DllImport("user32.dll")]
        public static extern IntPtr CreateWindowEx(
           uint dwExStyle,
           string lpClassName,
           string lpWindowName,
           uint dwStyle,
           int x,
           int y,
           int nWidth,
           int nHeight,
           IntPtr hWndParent,
           IntPtr hMenu,
           IntPtr hInstance,
           IntPtr lpParam);

        [DllImport("user32.dll")]
        public static extern uint GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        public static extern uint SetWindowLong(IntPtr hWnd, int nIndex, WndProcProc callback);

        [DllImport("user32.dll")]
        public static extern uint SetWindowLong(IntPtr hWnd, int nIndex, uint dwNewLong);

        public delegate int WndProcProc(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        public static extern int CallWindowProc(IntPtr lpPrevWndFunc, IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, ref Rect lpRect);

        [DllImport("user32.dll")]
        public static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool TrackPopupMenu(IntPtr hMenu, uint uFlags, int x, int y, int nReserved, IntPtr hWnd, IntPtr prcRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DestroyMenu(IntPtr hMenu);

        public const uint TPM_LEFTALIGN = 0;
        public const uint TPM_CENTERALIGN = 4;
        public const uint TPM_RIGHTALIGN = 8;
        public const uint TPM_TOPALIGN = 0;
        public const uint TPM_VCENTERALIGN = 0x10;
        public const uint TPM_BOTTOMALIGN = 0x20;

        [Flags]
        private enum KeyStates
        {
            None = 0,
            Down = 1,
            Toggled = 2
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern short GetKeyState(int keyCode);

        private static KeyStates GetKeyState(Keys key)
        {
            KeyStates state = KeyStates.None;

            short retVal = GetKeyState((int)key);

            //If the high-order bit is 1, the key is down
            //otherwise, it is up.
            if ((retVal & 0x8000) == 0x8000)
                state |= KeyStates.Down;

            //If the low-order bit is 1, the key is toggled.
            if ((retVal & 1) == 1)
                state |= KeyStates.Toggled;

            return state;
        }

        public static bool IsKeyDown(Keys key)
        {
            return KeyStates.Down ==
              (GetKeyState(key) & KeyStates.Down);
        }

        public static bool IsKeyToggled(Keys key)
        {
            return KeyStates.Toggled ==
              (GetKeyState(key) & KeyStates.Toggled);
        }

        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        public static extern int GetDeviceCaps(IntPtr hDc, int index);

        public const int LOGPIXELSX = 88;    /* Logical pixels/inch in X */
        public const int LOGPIXELSY = 90;    /* Logical pixels/inch in Y */

        public static Point GetScreenDpi()
        {
            IntPtr hDc = GetDC(IntPtr.Zero);
            return new Point(GetDeviceCaps(hDc, LOGPIXELSX), GetDeviceCaps(hDc, LOGPIXELSY));
        }

        [DllImport("user32.dll")]
        public static extern IntPtr SetWindowsHookEx(int idHook, HookCallbackProc lpfn, IntPtr hMod, int dwThreadId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        public const int WH_CALLWNDPROC = 4;
        public const int WH_KEYBOARD_LL = 13;

        public const UInt32 WS_POPUP = 0x80000000;
        public const UInt32 WS_CHILD = 0x40000000;
        public const UInt32 WS_CLIPSIBLINGS = 0x04000000;
        public const UInt32 WS_CLIPCHILDREN = 0x02000000;
        public const UInt32 WS_VISIBLE = 0x10000000;
        public const UInt32 WS_TABSTOP = 0x00010000;
        public const UInt32 WS_OVERLAPPEDWINDOW = 0x00CF0000;

        public const UInt32 WS_EX_CONTROLPARENT = 0x00010000;
        public const UInt32 WS_EX_APPWINDOW = 0x00040000;

        [DllImport("user32.dll")]
        public static extern int GetThreadDpiAwarenessContext();

        [DllImport("user32.dll")]
        public static extern int GetWindowDpiAwarenessContext(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern int SetThreadDpiAwarenessContext(int dpi);

        [DllImport("user32.dll")]
        public static extern bool IsValidDpiAwarenessContext(int first);

        public delegate int HookCallbackProc(int nCode, int wParam, int lParam);

        [DllImport("user32.dll")]
        public static extern int CallNextHookEx(IntPtr hhk, int nCode, int wParam, int lParam);

        public const int HC_ACTION = 0;

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        public const int WM_USER = 0x400;
        public const int WM_SMS_SENDED = WM_USER + 1;

        public static IntPtr GetExplorerWindowHandle(object explorer)
        {
            IntPtr explorerHWND = IntPtr.Zero;
            (explorer as IOLEWindow)?.GetWindow(out explorerHWND);
            return explorerHWND;
        }

        /// <summary>
        /// Find direct child window in parent
        /// </summary>
        /// <param name="hwndParent">Parent HWND</param>
        /// <param name="hwndChildAfter">Direct child HWND of window, search after this HWND</param>
        /// <param name="className">Class name</param>
        /// <param name="windowName">Window name</param>
        /// <returns>Finded HWND</returns>
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string className, string windowName);
    }

    [ComImport]
    [ComVisible(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("00000114-0000-0000-C000-000000000046")]
    internal interface IOLEWindow
    {
        void GetWindow(out IntPtr wnd);

        void ContextSensitiveHelp(bool fEnterMode);
    }

    public struct Rect
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }

    public struct CWPStruct
    {
        public int lParam;
        public int wParam;
        public uint Msg;
        public IntPtr hWnd;
    }

    public struct KbdLLStruct
    {
        public int vkCode;
        public int scanCode;
        public int flags;
        public int time;
        public IntPtr dwExtraInfo;
    }
}
