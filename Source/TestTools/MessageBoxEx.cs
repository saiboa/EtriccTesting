using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

//[assembly: SecurityPermission(SecurityAction.RequestMinimum, UnmanagedCode = true)]

namespace TestTools
{
    public class MessageBoxEx
    {
        #region Delegates

        public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        public delegate void TimerProc(IntPtr hWnd, uint uMsg, UIntPtr nIDEvent, uint dwTime);

        #endregion

        public const int WH_CALLWNDPROCRET = 12;
        public const int WM_DESTROY = 0x0002;
        public const int WM_INITDIALOG = 0x0110;
        public const int WM_TIMER = 0x0113;
        public const int WM_USER = 0x400;
        public const int DM_GETDEFID = WM_USER + 0;
        private const int TimerID = 42;
        private static readonly HookProc hookProc;
        private static readonly TimerProc hookTimer;
        private static uint hookTimeout;
        private static string hookCaption;
        private static IntPtr hHook;

        static MessageBoxEx()
        {
            hookProc = MessageBoxHookProc;
            hookTimer = MessageBoxTimerProc;
            hookTimeout = 0;
            hookCaption = null;
            hHook = IntPtr.Zero;
        }

        public static DialogResult Show(string text, uint uTimeout)
        {
            Setup("", uTimeout);
            return MessageBox.Show(text);
        }

        public static DialogResult Show(string text, string caption, uint uTimeout)
        {
            Setup(caption, uTimeout);
            return MessageBox.Show(text, caption);
        }

        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, uint uTimeout)
        {
            Setup(caption, uTimeout);
            return MessageBox.Show(text, caption, buttons);
        }

        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
                                        uint uTimeout)
        {
            Setup(caption, uTimeout);
            return MessageBox.Show(text, caption, buttons, icon);
        }

        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
                                        MessageBoxDefaultButton defButton, uint uTimeout)
        {
            Setup(caption, uTimeout);
            return MessageBox.Show(text, caption, buttons, icon, defButton);
        }

        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
                                        MessageBoxDefaultButton defButton, MessageBoxOptions options, uint uTimeout)
        {
            Setup(caption, uTimeout);
            return MessageBox.Show(text, caption, buttons, icon, defButton, options);
        }

        public static DialogResult Show(IWin32Window owner, string text, uint uTimeout)
        {
            Setup("", uTimeout);
            return MessageBox.Show(owner, text);
        }

        public static DialogResult Show(IWin32Window owner, string text, string caption, uint uTimeout)
        {
            Setup(caption, uTimeout);
            return MessageBox.Show(owner, text, caption);
        }

        public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons,
                                        uint uTimeout)
        {
            Setup(caption, uTimeout);
            return MessageBox.Show(owner, text, caption, buttons);
        }

        public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons,
                                        MessageBoxIcon icon, uint uTimeout)
        {
            Setup(caption, uTimeout);
            return MessageBox.Show(owner, text, caption, buttons, icon);
        }

        public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons,
                                        MessageBoxIcon icon, MessageBoxDefaultButton defButton, uint uTimeout)
        {
            Setup(caption, uTimeout);
            return MessageBox.Show(owner, text, caption, buttons, icon, defButton);
        }

        public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons,
                                        MessageBoxIcon icon, MessageBoxDefaultButton defButton,
                                        MessageBoxOptions options, uint uTimeout)
        {
            Setup(caption, uTimeout);
            return MessageBox.Show(owner, text, caption, buttons, icon, defButton, options);
        }

        [DllImport("User32.dll")]
        public static extern UIntPtr SetTimer(IntPtr hWnd, UIntPtr nIDEvent, uint uElapse, TimerProc lpTimerFunc);

        [DllImport("User32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);

        [DllImport("user32.dll")]
        public static extern int UnhookWindowsHookEx(IntPtr idHook);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr idHook, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxLength);

        [DllImport("user32.dll")]
        public static extern int EndDialog(IntPtr hDlg, IntPtr nResult);

        private static void Setup(string caption, uint uTimeout)
        {
            if (hHook != IntPtr.Zero)
                throw new NotSupportedException("multiple calls are not supported");

            hookTimeout = uTimeout;
            hookCaption = caption != null ? caption : "";
            hHook = SetWindowsHookEx(WH_CALLWNDPROCRET, hookProc, IntPtr.Zero, AppDomain.GetCurrentThreadId());
        }

        private static IntPtr MessageBoxHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
                return CallNextHookEx(hHook, nCode, wParam, lParam);

            var msg = (CWPRETSTRUCT) Marshal.PtrToStructure(lParam, typeof (CWPRETSTRUCT));
            IntPtr hook = hHook;

            if (hookCaption != null && msg.message == WM_INITDIALOG)
            {
                int nLength = GetWindowTextLength(msg.hwnd);
                var text = new StringBuilder(nLength + 1);

                GetWindowText(msg.hwnd, text, text.Capacity);

                if (hookCaption == text.ToString())
                {
                    hookCaption = null;
                    SetTimer(msg.hwnd, (UIntPtr) TimerID, hookTimeout, hookTimer);
                    UnhookWindowsHookEx(hHook);
                    hHook = IntPtr.Zero;
                }
            }

            return CallNextHookEx(hook, nCode, wParam, lParam);
        }

        private static void MessageBoxTimerProc(IntPtr hWnd, uint uMsg, UIntPtr nIDEvent, uint dwTime)
        {
            if (nIDEvent == (UIntPtr) TimerID)
            {
                var dw = (short) SendMessage(hWnd, DM_GETDEFID, IntPtr.Zero, IntPtr.Zero);

                EndDialog(hWnd, (IntPtr) dw);
            }
        }

        #region Nested type: CWPRETSTRUCT

        [StructLayout(LayoutKind.Sequential)]
        public struct CWPRETSTRUCT
        {
            public IntPtr lResult;
            public IntPtr lParam;
            public IntPtr wParam;
            public uint message;
            public IntPtr hwnd;
        };

        #endregion
    }
}