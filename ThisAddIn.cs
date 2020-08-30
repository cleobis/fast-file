using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Input;
using System.Security.AccessControl;

namespace QuickFile
{
    public partial class ThisAddIn
    {
        public TaskPaneControlWrapper taskPaneControl;
        public Microsoft.Office.Tools.CustomTaskPane customTaskPane;
        private InterceptKeys interceptKeys;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            taskPaneControl = new TaskPaneControlWrapper();
            customTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "My Task Pane");
            customTaskPane.Visible = true;
            interceptKeys = new InterceptKeys();
            interceptKeys.Attach();
        }   

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }

    // http://web.archive.org/web/20190828074433/https://blogs.msdn.microsoft.com/toub/2006/05/03/low-level-keyboard-hook-in-c/
    class InterceptKeys

    {

        private const int WH_KEYBOARD_LL = 13;
        private const int WH_KEYBOARD = 2;
        private LowLevelKeyboardProc _proc ;
        private IntPtr _hookID = IntPtr.Zero;

        private bool LeftCtrl = false;
        private bool RightCtrl = false;
        private bool LeftAlt = false;
        private bool RightAlt = false;
        private bool LeftShift = false;
        private bool RightShift = false;

        public InterceptKeys()
        {
            _proc = HookCallback;
        }
        ~InterceptKeys()
        {
            Detach();
        }

        public bool Attach()
        {
            if (_hookID != IntPtr.Zero)
            {
                Debug.WriteLine("Already attached.");
                return false;
            }
            _hookID = SetHook(_proc);
            return _hookID != IntPtr.Zero;
            ;
        }

        public void Detach()
        {
            if (_hookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
            }
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            using (ProcessThread thread = curProcess.Threads[0])
            {
                return SetWindowsHookEx(WH_KEYBOARD, proc, IntPtr.Zero, (uint)curProcess.Threads[0].Id);
            }
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            // Calls once n = 3 (call to peek), then n = 0 (calls to get the key)

            var key = KeyInterop.KeyFromVirtualKey((int)wParam);
            var flags = new KeystrokeFlags(lParam);
            switch (key)
            {
                case Key.LeftCtrl:
                    LeftCtrl = flags.IsDown;
                    break;
                case Key.RightCtrl:
                    RightCtrl = flags.IsDown;
                    break;
                case Key.LeftShift:
                    LeftShift = flags.IsDown;
                    break;
                case Key.RightShift:
                    RightShift = flags.IsDown;
                    break;
                case Key.LeftAlt:
                    LeftAlt = flags.IsDown;
                    break;
                case Key.RightAlt:
                    RightAlt = flags.IsDown;
                    break;
                case Key.V:
                    if ((LeftCtrl || RightCtrl) && (LeftShift || RightShift) && !(LeftAlt || RightAlt) && flags.IsDown)
                    {
                        Debug.WriteLine("Ctrl+Shift+V Down");
                        return IntPtr.Zero + 1;
                    }
                    break;
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        internal struct KeystrokeFlags
        {
            /* 0-15
                The repeat count. The value is the number of times the keystroke is repeated as a result of the user's holding down the key.
                16-23
                The scan code. The value depends on the OEM.
                24
                Indicates whether the key is an extended key, such as a function key or a key on the numeric keypad. The value is 1 if the key is an extended key; otherwise, it is 0.
                25-28
                Reserved.
                29
                The context code. The value is 1 if the ALT key is down; otherwise, it is 0.
                30
                The previous key state. The value is 1 if the key is down before the message is sent; it is 0 if the key is up.
                31
                The transition state. The value is 0 if the key is being pressed and 1 if it is being released.
            */
            private long raw;
            public KeystrokeFlags(IntPtr _in)
            {
                raw = _in.ToInt64();
            }
            public int Repeat
            {
                get { return (int)(raw & 0x0000FFFF); }
                //set { raw = (uint)(raw & ~mask0 | (value << loc0) & mask0); }
            }
            public int ScanCode
            {
                get { return (int)(raw & 0x00FF0000) >> 16; }
            }
            public bool Alt
            {
                get { return Convert.ToBoolean(raw & 0x20000000); }
            }
            public bool WasDown
            {
                get { return Convert.ToBoolean(raw & 0x40000000); }
            }
            public bool WasUp { 
                get { return !WasDown ;}
            }
            public bool IsUp
            {
                get { return Convert.ToBoolean(raw & 0x80000000); }
            }
            public bool IsDown
            {
                get { return !IsUp; }
            }
            public override String ToString()
            {
                return String.Format("KeystrokeFlags: Repeat {0}, ScanCode {1}, Alt {2}, WasDown {3}, IsDown {4}.", Repeat, ScanCode, Alt, WasDown, IsDown);
            }

        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

    }
}
