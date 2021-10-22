using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace TestTools
{
    public class UITestUtility
    {
        #region Methods of UITestUtility (22)

        public static void ClickOn(IntPtr hControl)
        {
            uint WM_LBUTTONDOWN = 0x0201;
            uint WM_LBUTTONUP = 0x0202;
            PostMessage1(hControl, WM_LBUTTONDOWN, 0, 0);
            PostMessage1(hControl, WM_LBUTTONUP, 0, 0);
        }

        public static IntPtr FindMainWindowHandle(string caption, int delay, int maxTries)
        {
            return FindTopLevelWindow(caption, delay, maxTries);
        }

        public static IntPtr FindMessageBox(string caption)
        {
            int delay = 100;
            int maxTries = 25;
            return FindTopLevelWindow(caption, delay, maxTries);
        }

        public static IntPtr FindTopLevelWindow(string caption, int delay, int maxTries)
        {
            IntPtr mwh = IntPtr.Zero;
            bool formFound = false;
            int attempts = 0;

            do
            {
                mwh = FindWindow(null, caption);
                if (mwh == IntPtr.Zero)
                {
                    Console.WriteLine("Form not yet found");
                    Thread.Sleep(delay);
                    ++attempts;
                }
                else
                {
                    Console.WriteLine("Form has been found");
                    formFound = true;
                }
            } while (!formFound && attempts < maxTries);

            //if (mwh != IntPtr.Zero)
            return mwh;
            //else
            //    return IntPtr.Zero;
            //else
            //    throw new Exception("Could not find Main Window");
        }

        // P/Invoke Aliases
        [DllImport("user32.dll", EntryPoint = "FindWindow", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public static IntPtr FindWindowByIndex(IntPtr hwndParent, int index)
        {
            if (index == 0)
                return hwndParent;
            else
            {
                int ct = 0;
                IntPtr result = IntPtr.Zero;
                do
                {
                    result = FindWindowEx(hwndParent, result, null, null);
                    if (result != IntPtr.Zero)
                        ++ct;
                } while (ct < index && result != IntPtr.Zero);

                return result;
            }
        }

        [DllImport("user32.dll", EntryPoint = "FindWindowEx", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass,
                                                 string lpszWindow);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="installScriptsDir"></param>
        /// <param name="buildBaseDir"></param>
        /// <param name="buildNr"></param>
        /// <param name="testApp"></param>
        /// <param name="buildDef"></param>
        /// <param name="buildConfig"></param>
        public static void GetAllParameters(string installScriptsDir, ref string buildBaseDir,
                                            ref string buildNr, ref string testApp, ref string buildDef,
                                            ref string buildConfig)
        {
            buildBaseDir = getBuildBasePath(installScriptsDir);
            buildNr = getBuildnr(installScriptsDir);

            if (buildNr.StartsWith("Etricc"))
                testApp = "Etricc";
            else if (buildNr.StartsWith("Epia"))
                testApp = "Epia";

            int ib = buildNr.IndexOf("-");
            int ie = buildNr.IndexOf("_");
            buildDef = buildNr.Substring(ib + 1, ie - (ib + 1)).Trim();

            if (installScriptsDir.IndexOf("Debug") > 0)
                buildConfig = "Debug";
            else
                buildConfig = "Release";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path">the path that include buildnumber
        /// example: X:\Nightly\Epia 3\Epia\Epia - Nightly_20080528.1\Mixed Platforms\Debug\InstallScripts
        /// it return X:\Nightly\Epia 3\Epia\Epia - Nightly_20080528.1
        /// </param>
        /// <returns></returns>
        public static string getBuildBasePath(string path)
        {
            int ib = path.LastIndexOf("\\");
            while (ib > 0)
            {
                string y = path.Substring(ib + 1);
                //MessageBox.Show(path.Substring(ib + 1));
                if (y.StartsWith("E") && y.IndexOf("-") > 0)
                {
                    return path;
                }
                path = path.Substring(0, ib);
                ib = path.LastIndexOf("\\");
            }

            return string.Empty;
        }

        /// <summary>
        /// return a buildnr
        /// example: X:\Nightly\Epia 3\Epia\Epia - Nightly_20080528.1\Mixed Platforms\Debug\InstallScripts
        /// return : Epia - Nightly_20080528.1
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string getBuildnr(string path)
        {
            string nr = string.Empty;
            int ib = path.LastIndexOf("\\");
            while (ib > 0)
            {
                string y = path.Substring(ib + 1);
                if (y.StartsWith("E") && y.IndexOf("-") > 0)
                    nr = y;
                // one level up
                path = path.Substring(0, ib);
                ib = path.LastIndexOf("\\");
            }

            return nr;
        }

        // Menu routines
        [DllImport("user32.dll")] // 
        public static extern IntPtr GetMenu(IntPtr hWnd);

        [DllImport("user32.dll")] // 
        public static extern int GetMenuItemID(IntPtr hMenu, int nPos);

        [DllImport("user32.dll")] // 
        public static extern IntPtr GetSubMenu(IntPtr hMenu, int nPos);

        // PostMessage
        /// <summary>
        ///     for WM_LBUTTONDOWN and WM_LBUTTONUP messages, I think WM_COMMAND message can be used too
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="Msg"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns>
        ///     returns without waiting for the message to be processed
        /// </returns>
        [DllImport("user32.dll", EntryPoint = "PostMessage", CharSet = CharSet.Auto)]
        public static extern bool PostMessage1(IntPtr hWnd, uint Msg, int wParam, int lParam);

        public static void SendChar(IntPtr hControl, char c)
        {
            uint WM_CHAR = 0x0102;
            SendMessage1(hControl, WM_CHAR, c, 0);
        }

        public static void SendChars(IntPtr hControl, string s)
        {
            foreach (char c in s)
            {
                SendChar(hControl, c);
            }
        }

        // SendMessage
        /// <summary>
        ///     calls the specified procedure and does not return until after 
        ///     the procedure has processed the Windows message
        ///     for WM_CHAR message
        /// </summary>
        /// <param name="hWnd"></param>
        /// <param name="Msg"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        [DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = CharSet.Auto)]
        public static extern void SendMessage1(IntPtr hWnd, uint Msg, int wParam, int lParam);

        // for WM_COMMAND message
        //
        [DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = CharSet.Auto)]
        public static extern void SendMessage2(IntPtr hWnd, uint Msg, int wParam, IntPtr lParam);

        // for WM_GETTEXT message
        // check the contents of a control on the AUT.
        [DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = CharSet.Auto)]
        public static extern int SendMessage3(IntPtr hWndControl, uint Msg, int wParam, byte[] lParam);

        // for LB_FINDSTRING message
        // used to determine if a particular string is in a listbox control.
        [DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = CharSet.Auto)]
        public static extern int SendMessage4(IntPtr hWnd, uint Msg, int wParam, string lParam);

        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
    }
}