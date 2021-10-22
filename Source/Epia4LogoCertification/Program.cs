using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.Windows.Automation;
using System.ServiceProcess;
using System.Configuration;
using System.Windows;
using System.Diagnostics;

using TestTools;

namespace Epia4LogoCertification
{
    class Program
    {
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Fields of Program (38)

        // Test Param. =======================================
        static string[] sTestCaseName = new string[100];
        static DateTime sTestStartUpTime = DateTime.Now;
        static DateTime sStartTime = DateTime.Now;
        static string sTestApp = string.Empty;
        static AutomationElement aeForm;
        static TimeSpan sTime;
        static int Counter;
        static int sTotalCounter;
        static int sTotalException;
        static int sTotalFailed;
        static int sTotalPassed;
        static int sTotalUntested;
        static int TestCheck = ConstCommon.TEST_UNDEFINED;
        static bool sOnlyUITest = false;
        static bool sStressTest = false;
        static string sEpiaClientRoot = ConstCommon.EPIA_CLIENT_ROOT;
        // PCinfo
        static public string PCName;
        static public string OSName;
        static public string OSVersion;
        static public string UICulture;
        static public string TimeOnPC;
        // Build info
        static string sBuildBaseDir = string.Empty;
        static string sBuildConfig = string.Empty;
        static string sBuildDef = string.Empty;
        static string sBuildNr = string.Empty;
        static string sTestToolsVersion = string.Empty;
        static string m_SystemDrive = string.Empty;
        static string UserPassword = "Egemin01";
        // Testcase not used =================================
        public static string sConfigurationName = string.Empty;
        static string sErrorMessage;
        static bool sEventEnd;
        static string sExcelVisible = string.Empty;
        static bool sAutoTest = true;
        static string sInstallScriptsDir = string.Empty;
        public static string sLayoutName = string.Empty;
        static string sServerRunAs = string.Empty;


        // LOG=================================================================
        public static string slogFilePath = @"C:\";
        static string sOutFilename = "TestResult";
        static string sOutFilePath = string.Empty;
        static StreamWriter Writer;
        // Build param ========================================================
        //static BuildStore   buildStore      = null;
        // default layout
        public const ToggleState DEFAULT_FULLSCREEN = ToggleState.Off;
        public const ToggleState DEFAULT_MAXIMIZED = ToggleState.Off;
        public const ToggleState DEFAULT_ALLOWRESIZE = ToggleState.On;
        public const string DEFAULT_XPOS = "0";
        public const string DEFAULT_YPOS = "0";
        public const string DEFAULT_WIDTH = "792";
        public const string DEFAULT_HEIGHT = "606";
        public const string DEFAULT_TITLE = "Egemin Shell";
        public const ToggleState DEFAULT_SHOW_RIBBON = ToggleState.Off;
        public const ToggleState DEFAULT_SHOW_MAINENU = ToggleState.On;
        public const ToggleState DEFAULT_SHOW_TOOLBARS = ToggleState.Off;
        public const ToggleState DEFAULT_SHOW_NAVIGATOR = ToggleState.On;
        public const ToggleState DEFAULT_STACK_SCREENS = ToggleState.Off;

        public static bool sLicWindowOpen = false;
        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Methods of Program (1)
        /// <summary>
        /// Retrieves the top-level window that contains the specified UI Automation element.
        /// </summary>
        /// <param name="element">The contained element.</param>
        /// <returns>The containing top-level window element.</returns>
        private AutomationElement GetTopLevelWindow(AutomationElement element)
        {
            TreeWalker walker = TreeWalker.ControlViewWalker;
            AutomationElement elementParent;
            AutomationElement node = element;
            //if (node == elementRoot) return node;
            do
            {
                elementParent = walker.GetParent(node);
                if (elementParent == AutomationElement.RootElement) break;
                node = elementParent;
            }
            while (true);
            return node;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="args">Application inputs
        ///     1   InstallScripts Directory
        ///     2   Visible or Invisible (Excel )
        ///     3   true or false (auto test) 
        ///     </param>
        [STAThread]
        static void Main(string[] args)
        {
            AutomationEventHandler UIALayoutRibbonOnEventHandler = new AutomationEventHandler(OnLicenseConnectionUIAEvent);
            // Add LicenseReconnection window Event Handler
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutRibbonOnEventHandler);

            string pa = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE", EnvironmentVariableTarget.Machine);
            //System.Windows.Forms.MessageBox.Show("Intpyr size: " + IntPtr.Size, "System");

            int sysBit = ((String.IsNullOrEmpty(pa) || String.Compare(pa, 0, "x86", 0, 3, true) == 0) ? 32 : 64);
            //System.Windows.Forms.MessageBox.Show( "system (bit): "+ sysBit , "System");

            if (sysBit == 64)
                //sEpiaClientRoot = "Team Systems\\Epia 4\\Epia\\Development\\Dev07\\bin\\Any CPU\\Debug\\Epia Shell";
                sEpiaClientRoot = "Program Files (x86)\\Egemin\\Epia Shell";

            //System.Windows.Forms.MessageBox.Show("system (bit): " + sysBit, sEpiaClientRoot);
            sOnlyUITest = true;
            string x = System.Configuration.ConfigurationManager.AppSettings.Get("StressTest");
            if (x.ToLower().StartsWith("true"))
                sStressTest = true;

            UserPassword = System.Configuration.ConfigurationManager.AppSettings.Get("CurrentUserPassword");

            Epia3Common.CreateTestLog(ref slogFilePath, sOutFilePath, sOutFilename, ref Writer);

            Console.WriteLine("Test started:");
            Epia3Common.WriteTestLogMsg(slogFilePath, "Test started: ", sOnlyUITest);

            try  // Get test PC info======================================
            {
                TestTools.HelpUtilities.SavePCInfo("y");
                TestTools.HelpUtilities.GetPCInfo(out PCName, out OSName, out OSVersion, out UICulture, out TimeOnPC);
                Console.WriteLine("PCName : " + PCName);
                Console.WriteLine("OSName : " + OSName);
                Console.WriteLine("OSVersion : " + OSVersion);
                Console.WriteLine("UICulture : " + UICulture);
                Console.WriteLine("TimeOnPC : " + TimeOnPC);
                string windir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
                m_SystemDrive = Path.GetPathRoot(windir);

                Epia3Common.WriteTestLogMsg(slogFilePath, "Test PC : " + PCName, !sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "OSName : " + OSName, !sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "OSVersion : " + OSVersion, !sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "UICulture : " + UICulture, !sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "TimeOnPC : " + TimeOnPC, !sOnlyUITest);

                #region  Shell
                Utilities.CloseProcess("Egemin.Epia.Shell");
                AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);
                sEventEnd = false;
                TestCheck = ConstCommon.TEST_PASS;

                Thread.Sleep(5000);

                // Start Shell
                //TestTools.Utilities.StartProcessNoWait(sInstallScriptsDir, Constants.SHELL_BAT, string.Empty);
                TestTools.Utilities.StartProcessNoWait(m_SystemDrive + sEpiaClientRoot,
                    ConstCommon.EGEMIN_EPIA_SHELL_EXE, string.Empty);
                //--------------------------
                sStartTime = DateTime.Now;
                sTime = DateTime.Now - sStartTime;
                int wt = 0;
                Console.WriteLine(" time is :" + sTime.TotalMilliseconds);
                while (sEventEnd == false && wt < 60)
                {
                    Thread.Sleep(2000);
                    //sTime = DateTime.Now - sStartTime;
                    wt = wt + 2;
                    Console.WriteLine("wait shell start up time is (sec) : " + wt);
                }

                Console.WriteLine("Shell started after (sec) : " + 2 * wt);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIAShellEventHandler);

                Thread.Sleep(4000);
                Console.WriteLine("TestCkeck : " + TestCheck.ToString());
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    throw new Exception("shell start up failed:" + sErrorMessage);
                }
                Thread.Sleep(4000);
                #endregion

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Check PC-info");
            }

            sTestCaseName[0] = LAYOUT_FIND_LAYOUT_PANEL;
            sTestCaseName[1] = LAYOUT_INITIAL_X_POSITION;
            sTestCaseName[2] = LAYOUT_INITIAL_Y_POSITION;
            sTestCaseName[3] = LAYOUT_INITIAL_WIDTH;
            sTestCaseName[4] = LAYOUT_INITIAL_HEIGHT;
            sTestCaseName[5] = LAYOUT_ALLOW_RESIZE;
            sTestCaseName[6] = LAYOUT_FULL_SCREEN;
            sTestCaseName[7] = LAYOUT_MAXIMIZED;
            sTestCaseName[8] = LAYOUT_RIBBON_ON;
            sTestCaseName[9] = LAYOUT_CANCEL_BUTTON;
            sTestCaseName[10] = SETTING_LANGUAGE;
            sTestCaseName[11] = SHELL_CONFIGURATION_SECURITY;
            sTestCaseName[12] = LOGON_CURRENT_USER;
            sTestCaseName[13] = SHELL_SHUTDOWN;
            sTestCaseName[14] = SHELL_LOGOFF;
            sTestCaseName[15] = EPIA3_CLOSE;
            //=============================================//
            sTestCaseName[34] = LAYOUT_OPEN;
            sTestCaseName[35] = CONFIGURATION_OPEN;
            sTestCaseName[40] = CONFIGURATION_SAVE;
            sTestCaseName[36] = CONFIGURATION_SECURITY_UNSECURED;

            //sTestCaseName[0] = CONFIGURATION_FIND_GRIDVIEW;
            //sTestCaseName[19] = LAYOUT_MAINMENU_ON;

            try
            {
                Console.WriteLine("TestCkeck : " + TestCheck.ToString());
                Console.WriteLine("Application is started : ");
                aeForm = null;
                string formID = "MainForm";
                DateTime mAppTime = DateTime.Now;
                AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeForm, formID, mAppTime, 300);

                if (aeForm == null)
                {
                    AutomationElement aeError = AUIUtilities.FindElementByID("ErrorScreen", AutomationElement.RootElement);
                    if (aeError != null)
                        AUICommon.ErrorWindowHandling(aeError, ref sErrorMessage);
                    else
                        sErrorMessage = "Application Startup failed.";

                    throw new Exception(sErrorMessage);
                }
                else
                    Console.WriteLine("Application maeForm name : " + aeForm.Current.Name);

                // Excel file not for EpiaTestPC3 and EPIATESTSERVER3
                if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
                    || PCName.ToUpper().Equals("EPIATESTSRV3V1"))
                    Epia3Common.WriteTestLogMsg(slogFilePath, "No Excel due to: " + PCName, sOnlyUITest);
                else
                {
                    Console.WriteLine("Write test result :   ToDo1 ");
                    //System.Windows.Forms.MessageBox.Show("slogFilePath:" + slogFilePath);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Start test on : " + PCName, !sOnlyUITest);
                }

                // start test----------------------------------------------------------
                int sResult = ConstCommon.TEST_UNDEFINED;
                int aantal = 16;
                if (sOnlyUITest)
                {
                    string testType = System.Configuration.ConfigurationManager.AppSettings.Get("TestCase");
                    if (testType.ToLower().StartsWith("all"))
                    {
                        aantal = 11;
                    }
                    else
                    {
                        int thisTest = Convert.ToInt16(testType);
                        aantal = 1;
                        sTestCaseName[0] = sTestCaseName[thisTest - 1];
                    }
                }
                else
                {
                    // check Shell proc. if exist , get Proc id
                    System.Diagnostics.Process ShellProcess = null;
                    int pID = TestTools.Utilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out ShellProcess);
                    Console.WriteLine("Proc ID:" + pID);
                    aeForm = AutomationElement.FromHandle(ShellProcess.MainWindowHandle);

                    if (aeForm == null)
                    {
                        Console.WriteLine("aeForm  not found : ");
                        return;
                    }
                    else
                        Console.WriteLine("aeForm found name : " + aeForm.Current.Name);
                }


                TimeSpan mTime;
                DateTime sTestStartTime = DateTime.Now;
                while (sTestStartTime.Minute % 30 < 2 || sTestStartTime.Minute % 30 > 25)
                {
                    Console.WriteLine("Not right minute to testing  : " + sTestStartTime.Minute);
                    System.Threading.Thread.Sleep(15000);
                    ////TestTools.AUIUtilities.ClickElement(AutomationElement.RootElement.);
                    sTestStartTime = DateTime.Now;
                }

                while (Counter < aantal)
                {

                    if (Counter == 0)
                    {
                        Console.WriteLine("Counter is  : " + Counter + "    --------- first get aeForm");
                        aeForm = null;
                        mAppTime = DateTime.Now;
                        AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeForm, formID, mAppTime, 300);

                        if (aeForm == null)
                        {
                            AutomationElement aeError = AUIUtilities.FindElementByID("ErrorScreen", AutomationElement.RootElement);
                            if (aeError != null)
                                AUICommon.ErrorWindowHandling(aeError, ref sErrorMessage);
                            else
                                sErrorMessage = "Application Startup failed.";

                            throw new Exception(sErrorMessage);
                        }
                        else
                            Console.WriteLine("Application maeForm name : " + aeForm.Current.Name);

                    }


                    string testHour = System.Configuration.ConfigurationManager.AppSettings.Get("hour");
                    int hourInt = Convert.ToInt16(testHour);

                    //mTime = DateTime.Now - sTestStartTime;
                    mTime = DateTime.Now.Subtract(sTestStartTime);
                    Console.WriteLine("mTime.TotalHours : " + mTime.TotalHours);
                    if (mTime.Hours >= hourInt)
                        break;

                    sResult = ConstCommon.TEST_UNDEFINED;
                    switch (sTestCaseName[Counter])
                    {
                        case LAYOUT_FIND_LAYOUT_PANEL:
                            LayoutFindLayoutPanel(LAYOUT_FIND_LAYOUT_PANEL, aeForm, out sResult);
                            break;
                        case LAYOUT_INITIAL_X_POSITION:
                            LayoutInitialXPosition(LAYOUT_INITIAL_X_POSITION, aeForm, out sResult);
                            break;
                        case LAYOUT_INITIAL_Y_POSITION:
                            LayoutInitialYPosition(LAYOUT_INITIAL_Y_POSITION, aeForm, out sResult);
                            break;
                        case LAYOUT_INITIAL_WIDTH:
                            LayoutInitialWidth(LAYOUT_INITIAL_WIDTH, aeForm, out sResult);
                            break;
                        case LAYOUT_INITIAL_HEIGHT:
                            LayoutInitialHeight(LAYOUT_INITIAL_HEIGHT, aeForm, out sResult);
                            break;
                        case LAYOUT_TITLE:
                            LayoutTitle(LAYOUT_TITLE, aeForm, out sResult);
                            break;
                        case LAYOUT_ALLOW_RESIZE:
                            LayoutAllowResize(LAYOUT_ALLOW_RESIZE, aeForm, out sResult);
                            break;
                        case LAYOUT_FULL_SCREEN:
                            LayoutFullScreen(LAYOUT_FULL_SCREEN, aeForm, out sResult);
                            break;
                        case LAYOUT_MAXIMIZED:
                            LayoutMaximized(LAYOUT_MAXIMIZED, aeForm, out sResult);
                            break;
                        case LAYOUT_RIBBON_ON:
                            LayoutRibbonOn(LAYOUT_RIBBON_ON, aeForm, out sResult);
                            break;
                        case LAYOUT_NAVIGATOR_OFF:
                            LayoutNavigatorOff(LAYOUT_NAVIGATOR_OFF, aeForm, out sResult);
                            break;
                        case LAYOUT_CANCEL_BUTTON:
                            LayoutCancelButton(LAYOUT_CANCEL_BUTTON, aeForm, out sResult);
                            break;
                        case SETTING_LANGUAGE:
                            LanguageSetting(SETTING_LANGUAGE, aeForm, out sResult);
                            break;
                        case SHELL_CONFIGURATION_SECURITY:
                            ShellConfigSecurityWindowsUser(SHELL_CONFIGURATION_SECURITY, aeForm, out sResult);
                            break;
                        case LOGON_CURRENT_USER:
                            LogonCurrentUser(LOGON_CURRENT_USER, aeForm, out sResult);
                            break;
                        case LOGON_EPIA_ADMINISTRATOR:
                            LogonEpiaAdministrator(LOGON_EPIA_ADMINISTRATOR, aeForm, out sResult);
                            break;
                        case SHELL_SHUTDOWN:
                            ShellShutdown(SHELL_SHUTDOWN, aeForm, out sResult);
                            break;
                        case SHELL_LOGOFF:
                            ShellLogoff(SHELL_LOGOFF, aeForm, out sResult);
                            break;
                        case EPIA3_CLOSE:
                            Epia3Close(EPIA3_CLOSE, aeForm, out sResult);
                            break;
                        //======================================================================//
                        //case LAYOUT_MAINMENU_ON:
                        //    LayoutMainMenuOn(LAYOUT_MAINMENU_ON, aeForm, out sResult);
                        //    break
                        case LAYOUT_OPEN:
                            LayoutOpen(LAYOUT_OPEN, aeForm, out sResult);
                            break;
                        case CONFIGURATION_SAVE:
                            ConfigSave(CONFIGURATION_SAVE, aeForm, out sResult);
                            break;
                        case CONFIGURATION_OPEN:
                            ConfigOpen(CONFIGURATION_OPEN, aeForm, out sResult);
                            break;
                        case CONFIGURATION_SECURITY_UNSECURED:
                            ConfigSecurityUnsecured(CONFIGURATION_SECURITY_UNSECURED, aeForm, out sResult);
                            break;
                        case CONFIGURATION_FIND_GRIDVIEW:
                            ConfigFindGridView(CONFIGURATION_FIND_GRIDVIEW, aeForm, out sResult);
                            break;
                        default:
                            break;
                    }

                    if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
                        || PCName.ToUpper().Equals("EPIATESTSRV3V1"))
                        Console.WriteLine("No Excel due to: " + PCName);
                    else
                    {
                        Console.WriteLine("Write test result :  wait 20 sec  ");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "TestCase : " + sTestCaseName[Counter] + "  tested", !sOnlyUITest);
                        //WriteResult(sResult, Counter, sTestCaseName[Counter], xSheet, sErrorMessage);
                    }

                    sErrorMessage = string.Empty;
                    ++sTotalCounter;
                    if (sResult == ConstCommon.TEST_PASS)
                        ++sTotalPassed;
                    if (sResult == ConstCommon.TEST_FAIL)
                        ++sTotalFailed;
                    if (sResult == ConstCommon.TEST_EXCEPTION)
                        ++sTotalException;
                    if (sResult == ConstCommon.TEST_UNDEFINED)
                        ++sTotalUntested;


                    ++Counter;

                    System.Threading.Thread.Sleep(15000);
                    System.Windows.Point point = new Point(1, 100);
                    Input.MoveToAndRightClick(point);
                    Thread.Sleep(2000);
                    point.Y = point.Y + 300;
                    Input.MoveToAndClick(point);


                    if (sStressTest)
                        Counter = Counter % 10;

                    sTestStartTime = DateTime.Now;
                    while (sTestStartTime.Minute % 30 < 2 || sTestStartTime.Minute % 30 > 25)
                    {
                        Console.WriteLine("=========>   time interval for Crash Recoveries test  : " + sTestStartTime.Minute);
                        System.Threading.Thread.Sleep(15000);
                        sTestStartTime = DateTime.Now;
                        System.Threading.Thread.Sleep(15000);
                        point = new Point(1, 100);
                        Input.MoveToAndRightClick(point);
                        Thread.Sleep(2000);
                        point.Y = point.Y + 300;
                        Input.MoveToAndClick(point);
                    }
                }

                Epia3Common.WriteTestLogMsg(slogFilePath, "Total Tests: " + Counter, sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "Total Passed: " + sTotalPassed, sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "Total Failed: " + sTotalFailed, sOnlyUITest);

                if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
                    || PCName.ToUpper().Equals("EPIATESTSRV3V1"))
                    Epia3Common.WriteTestLogMsg(slogFilePath, "No Excel due to: " + PCName, sOnlyUITest);
                else
                {
                    Console.WriteLine("Write test result :   ToDo3 ");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "TestCase : " + sTestCaseName[Counter] + "  tested", sOnlyUITest);
                }


                // Close LogFile
                Epia3Common.CloseTestLog(slogFilePath, sOnlyUITest);

                Console.WriteLine("\nClosing application in 10 seconds");
                if (sOnlyUITest)
                    Thread.Sleep(10000000);
                else
                    Thread.Sleep(10000);
                /*
                AutomationElement aeForm1 = AUIUtilities.FindElementByID("MainForm", AutomationElement.RootElement);
                if (aeForm1 != null)
                {
                    WindowPattern wpCloseForm =
                      (WindowPattern)aeForm1.GetCurrentPattern(WindowPattern.Pattern);
                    wpCloseForm.Close();
                }
                */
                // close CommandHost
                Thread.Sleep(10000);
                TestTools.Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                TestTools.Utilities.CloseProcess("cmd");
                Console.WriteLine("\nEnd test run\n");
                //Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Thread.Sleep(200000000);
                if (sAutoTest)
                {

                    FileManipulation.UpdateStatusInTestInfoFile(sBuildBaseDir, "GUI Tests Exception -->" + sOutFilename + ".log", ConstCommon.EPIA);

                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Exception -->" + sOutFilename + ".log:" + ConstCommon.EPIA, sOnlyUITest);

                    Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace, sOnlyUITest);

                    Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                    Utilities.CloseProcess("cmd");
                    FileManipulation.UpdateTestWorkingFile(sBuildBaseDir, "false");
                }
            }
        }
        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region TestCase Name
        private const string LAYOUT_FIND_LAYOUT_PANEL = "LayoutFindLayoutPanel";
        private const string LAYOUT_INITIAL_X_POSITION = "LayoutInitialXPosition";
        private const string LAYOUT_INITIAL_Y_POSITION = "LayoutInitialYPosition";
        private const string LAYOUT_INITIAL_WIDTH = "LayoutInitialWidth";
        private const string LAYOUT_INITIAL_HEIGHT = "LayoutInitialHeight";
        private const string LAYOUT_TITLE = "LayoutTitle";
        private const string LAYOUT_ALLOW_RESIZE = "LayoutAllowResize";
        private const string LAYOUT_FULL_SCREEN = "LayoutFullScreen";
        private const string LAYOUT_MAXIMIZED = "LayoutMaximized";
        private const string LAYOUT_RIBBON_ON = "LayoutRibbonOn";
        private const string LAYOUT_CANCEL_BUTTON = "LayoutCancelButton";
        private const string SETTING_LANGUAGE = "LanguageSetting";
        private const string SHELL_CONFIGURATION_SECURITY = "ShellConfigSecurityWindowsUser";
        private const string LOGON_CURRENT_USER = "LogonCurrentUser";
        private const string LOGON_EPIA_ADMINISTRATOR = "LogonEpiaAdmin";
        private const string SHELL_SHUTDOWN = "ShellShutdown";
        private const string SHELL_LOGOFF = "ShellLogOff";
        private const string EPIA3_CLOSE = "Epia3Close";
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        private const string LAYOUT_NAVIGATOR_OFF = "LayoutNavigatorOff";

        private const string LAYOUT_OPEN = "LayoutOpen";
        private const string CONFIGURATION_OPEN = "ConfigOpen";
        private const string CONFIGURATION_SAVE = "ConfigSave";
        private const string CONFIGURATION_SECURITY_UNSECURED = "ConfigSecurityUnsecured";
        //private const string CONFIGURATION_SECURITY_EPIA = "ConfigSecurityEpia";
        private const string CONFIGURATION_FIND_GRIDVIEW = "ConfigFindGridView";
        #endregion TestCase Name
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Test Cases -------------------------------------------------------------------------------------------
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutFindLayoutPanel
        public static void LayoutFindLayoutPanel(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIAFindLayoutPanelEventHandler = new AutomationEventHandler(OnFindLayoutPanelUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                //AutomationEventHandler UIAFindLayoutPanelEventHandler = new AutomationEventHandler(OnFindLayoutPanelUIAEvent);
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIAFindLayoutPanelEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIAFindLayoutPanelEventHandler);

                if (mTime.Seconds > 600)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "After 10 min, Test is still running";
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    return;
                }

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    Console.WriteLine(testname);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                }

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIAFindLayoutPanelEventHandler);


            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutInitialXPosition
        public static void LayoutInitialXPosition(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutXPosEventHandler = new AutomationEventHandler(OnLayoutXPosUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutXPosEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Console.WriteLine("click my layout :");
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Console.WriteLine("click my layout :");
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                    Console.WriteLine("wait time is :" + mTime.TotalMilliseconds);

                }
                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                Thread.Sleep(5000);
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // GetClickablePoint point.
                    //Console.WriteLine("GetClickablePoint point");
                    //System.Windows.Point clickablePoint = root.GetClickablePoint();
                    //Console.WriteLine("clickablePointX=" + clickablePoint.X);
                    //Console.WriteLine("clickablePointY=" + clickablePoint.Y);
                    //Epia3Common.WriteTestLogMsg(slogFilePath,"clickablePointX=" + clickablePoint.X);
                    //System.Windows.Forms.Cursor.Position = new System.Drawing.Point((int)clickablePoint.X, (int)clickablePoint.Y, sOnlyUITest);

                    double leftValue = root.Current.BoundingRectangle.Left;
                    Console.WriteLine("Current UI left value " + leftValue);

                    if (Math.Abs(leftValue - 200) < 3)
                    {
                        Console.WriteLine("left value near 200");
                        Console.WriteLine("\nTest scenario Check Initial X position: Pass");
                        result = ConstCommon.TEST_PASS;
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    }
                    else
                    {
                        Console.WriteLine("Failed Xpos, should near to 200 , left value =" + leftValue);
                        Console.WriteLine("\nTest scenario Check Initial X position: *FAIL*");
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ": Failed Xpos, should near 200 , left value =" + leftValue, sOnlyUITest);
                    }
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutXPosEventHandler);

            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutInitialYPosition
        public static void LayoutInitialYPosition(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutYPosEventHandler = new AutomationEventHandler(OnLayoutYPosUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutYPosEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement,
                     UIALayoutYPosEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // getTop value.
                    Console.WriteLine("Get Form Top Value");
                    double topValue = root.Current.BoundingRectangle.Top;
                    Console.WriteLine("Current Y top value " + topValue);

                    if (topValue == 100)
                    {
                        Console.WriteLine("top value = 100");
                        Console.WriteLine("\nTest scenario Check Initial Y position: Pass");
                        result = ConstCommon.TEST_PASS;
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    }
                    else
                    {
                        Console.WriteLine("top value = " + topValue);
                        Console.WriteLine("\nTest scenario Check Initial Y position: *FAIL*");
                        sErrorMessage = " top value is " + topValue;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                    }
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutYPosEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutInitialWidth
        public static void LayoutInitialWidth(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutWidthEventHandler = new AutomationEventHandler(OnLayoutWidthUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutWidthEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutWidthEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // GetClickablePoint point.
                    Console.WriteLine("Get Form Width");
                    double width = root.Current.BoundingRectangle.Width - 8;
                    Console.WriteLine("Form Width=" + width);

                    if (Math.Abs(width - 600) < 10)
                    {
                        Console.WriteLine("Width == 600");
                        Console.WriteLine("\nTest scenario Check LayoutInitialWidth: Pass");
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        sErrorMessage = "width is" + width;
                        Console.WriteLine("Width ==" + width);
                        Console.WriteLine("\nTest scenario Check LayoutInitialWidth: *FAIL*");
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                    }
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutWidthEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutInitialHeight
        public static void LayoutInitialHeight(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutHeightEventHandler = new AutomationEventHandler(OnLayoutHeightUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                     ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutHeightEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutHeightEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Get Form Height");
                    double height = root.Current.BoundingRectangle.Height - 34;
                    Console.WriteLine("Form Height=" + height);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Get Current Height:" + height, sOnlyUITest);

                    if (System.Math.Abs(height - 500) < 10)
                    {
                        Console.WriteLine("Height == 500");
                        Console.WriteLine("\nTest scenario Check LayoutInitialHeight: Pass");
                        result = ConstCommon.TEST_PASS;
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    }
                    else
                    {
                        Console.WriteLine("Height should be near 300 , but it is ==" + height);
                        Console.WriteLine("\nTest scenario Check LayoutInitialHeight: *FAIL*");
                        sErrorMessage = "Height is " + height;
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    }
                }

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutHeightEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutTitle
        public static void LayoutTitle(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutTitleEventHandler = new AutomationEventHandler(OnLayoutTitleUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutTitleEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutTitleEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    string title = root.Current.Name;
                    if (title.Equals("abcdefg"))
                    {
                        Console.WriteLine("window title = " + "abcdefg");
                        Console.WriteLine("\nTest scenario Check Title: Pass");
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        sErrorMessage = "Title is:" + title;
                        Console.WriteLine("window title abcdefg not found ,but it is " + title);
                        Console.WriteLine("\nTest scenario Check Title: *FAIL*");
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                    }
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutTitleEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutAllowResize
        public static void LayoutAllowResize(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutResizeEventHandler = new AutomationEventHandler(OnLayoutResizeUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutResizeEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }
                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    double Width = 400;
                    double Height = 800;
                    TransformPattern tranform =
                    root.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                    if (tranform != null)
                        tranform.Resize(Width, Height);

                    Thread.Sleep(3000);

                    if (root.Current.BoundingRectangle.Width == Width &&
                        root.Current.BoundingRectangle.Height == Height)
                    {
                        Console.WriteLine("\nTest scenario Resize: Pass1");
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        sErrorMessage = "current width=" + root.Current.BoundingRectangle.Width
                            + " --- "
                            + "current height=" + root.Current.BoundingRectangle.Height;
                        Console.WriteLine("current width=" + root.Current.BoundingRectangle.Width);
                        Console.WriteLine("current height=" + root.Current.BoundingRectangle.Height);
                        Console.WriteLine("\nTest scenario Resize: *FAIL*");
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ": " + sErrorMessage, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                    }
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutResizeEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutFullScreen
        public static void LayoutFullScreen(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutFullScreenEventHandler = new AutomationEventHandler(OnLayoutFullScreenUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutFullScreenEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(2000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(2000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement,
                     UIALayoutFullScreenEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    double left = root.Current.BoundingRectangle.Left;
                    double right = root.Current.BoundingRectangle.Right;
                    double top = root.Current.BoundingRectangle.Top;
                    double bottom = root.Current.BoundingRectangle.Bottom;

                    Console.WriteLine("Left=" + left);
                    //Console.WriteLine("right=" + right);
                    Console.WriteLine("top=" + top);
                    //Console.WriteLine("bottom=" + bottom);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Left=" + left, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "right=" + right, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "top=" + top, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "bottom=" + bottom, sOnlyUITest);

                    if (left == 0 && top == 0)
                    {
                        Console.WriteLine("\nTest scenario Check Full Screen: Pass");
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        Console.WriteLine("\nTest scenario Check Full Screen: *FAIL*");
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    }
                }

                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutFullScreenEventHandler);

                if (result == ConstCommon.TEST_PASS)
                {
                    ReturnDefaultLayout(root, out result);

                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(2000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutFullScreenEventHandler);
            }

        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutMaximized
        public static void LayoutMaximized(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutMaximizedScreenEventHandler = new AutomationEventHandler(OnLayoutMaximizedScreenUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MaximizedScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutMaximizedScreenEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutMaximizedScreenEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    double left = root.Current.BoundingRectangle.Left;
                    double right = root.Current.BoundingRectangle.Right;
                    double top = root.Current.BoundingRectangle.Top;
                    double bottom = root.Current.BoundingRectangle.Bottom;

                    Console.WriteLine("Left=" + left);
                    Console.WriteLine("top=" + top);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Left=" + left, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "right=" + right, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "top=" + top, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "bottom=" + bottom, sOnlyUITest);

                    if (Math.Abs(left - 0) < 10 && Math.Abs(top - 0) < 10)
                    {
                        Console.WriteLine("\nTest scenario Check Full Screen: Pass");
                        //Epia3Common.WriteTestLogPass(slogFilePath,testname, sOnlyUITest, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        Console.WriteLine("\nTest scenario Check Full Screen: *FAIL*");
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    }
                }

                if (result == ConstCommon.TEST_PASS)
                {
                    string BtnCloseID = "Close";
                    AutomationElement aeClose = AUIUtilities.FindElementByID(BtnCloseID, root);

                    if (aeClose != null)
                    {
                        Console.WriteLine("\nTest scenario Check Miximized Screen: Pass");
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        Console.WriteLine("\nTest scenario Check Miximized Screen: *FAIL*");
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    }
                }

                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutMaximizedScreenEventHandler);

                if (result == ConstCommon.TEST_PASS)
                {
                    ReturnDefaultLayout(root, out result);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutMaximizedScreenEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutRibbonOn
        public static void LayoutRibbonOn(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;

            AutomationEventHandler UIALayoutRibbonOnEventHandler = new AutomationEventHandler(OnLayoutRibbonOnUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutRibbonOnEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement,
                     UIALayoutRibbonOnEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    AutomationElement aeRibbon = AUIUtilities.FindElementByID("_MainForm_Toolbars_Dock_Area_Top", root);
                    if (aeRibbon == null)
                    {
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        sErrorMessage = "ribbon is not found";
                    }
                    else
                    {
                        if (aeRibbon.Current.BoundingRectangle.Height > 50)
                        {
                            Console.WriteLine("\nTest scenario Check Ribbon ON: Pass");
                            Epia3Common.WriteTestLogMsg(slogFilePath, "ribbon height is: " + aeRibbon.Current.BoundingRectangle.Height, sOnlyUITest);
                            Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                            result = ConstCommon.TEST_PASS;
                        }
                        else
                        {
                            Console.WriteLine("\nTest scenario Check Ribbon: *FAIL*");
                            sErrorMessage = "ribbon height is: " + aeRibbon.Current.BoundingRectangle.Height;
                            result = ConstCommon.TEST_FAIL;
                            Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        }
                    }
                }

                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutRibbonOnEventHandler);

                if (result == ConstCommon.TEST_PASS)
                {
                    ReturnDefaultLayout(root, out result);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutRibbonOnEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutCancelButton
        public static void LayoutCancelButton(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;

            AutomationEventHandler UIALayoutCancelButtonEventHandler = new AutomationEventHandler(OnLayoutCancelButtonUIAEvent);

            try
            {
                ReturnDefaultLayout(root, out result);

                AutomationElement aeYourLayouts = null;
                if (result == ConstCommon.TEST_PASS)
                {
                    aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                        ConstCommon.MY_LAYOUT, ref sErrorMessage);
                    if (aeYourLayouts == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                }
                else
                {
                    sErrorMessage = "ReturnDefaultLayout failed";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutCancelButtonEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutCancelButtonEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutCancelButtonEventHandler);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    ValidateDefaultLayout(root, out result);
                    if (result == ConstCommon.TEST_PASS)
                    {
                        Console.WriteLine("\nTest scenario Check Cancel Button: Pass");
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        Console.WriteLine("\nTest scenario Check Cancel Button: *FAIL*");
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    }
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutCancelButtonEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutNavigatorOff
        public static void LayoutNavigatorOff(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;

            AutomationEventHandler UIALayoutNavigatorOffEventHandler = new AutomationEventHandler(OnLayoutNavigatorOffUIAEvent);
            try
            {
                AutomationElement aeYourLayouts = null;
                aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutNavigatorOffEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutNavigatorOffEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutNavigatorOffEventHandler);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    DateTime time = DateTime.Now;
                    AutomationElement aeNav = null;
                    aeNav = AUIUtilities.FindElementByID("m_TreeView", root);
                    //AUIUtilities.WaitUntilElementByIDFound(root, ref aeNav, "m_TreeView", time, 60);
                    if (aeNav == null)
                        result = ConstCommon.TEST_PASS;
                    else
                    {
                        sErrorMessage = "Navigator still exist";
                        result = ConstCommon.TEST_FAIL;
                    }

                    Thread.Sleep(3000);

                    if (result == ConstCommon.TEST_PASS)
                    {
                        Console.WriteLine("\nTest scenario Navigator Off: Pass");
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        Console.WriteLine("\nTest scenario Navigator Off: *FAIL*");
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    }
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALayoutNavigatorOffEventHandler);

                Input.SendKeyboardInput(System.Windows.Input.Key.Space, true);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LanguageSetting
        public static void LanguageSetting(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALanguageSettingEventHandler = new AutomationEventHandler(OnLanguageSettingUIAEvent);

            #region test
            try
            {
                AutomationElement aeMySettings = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    "My settings", ref sErrorMessage);
                if (aeMySettings == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLanguageSetting window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALanguageSettingEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeMySettings);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement,
                     UIALanguageSettingEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                AutomationElement aeTreeView = null;
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }
                else
                {
                    string id = "m_TreeView";
                    DateTime sTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(root, ref aeTreeView, id, sTime, 60);

                    if (aeTreeView == null)
                    {
                        sErrorMessage = "aeTreeView not found name : ";
                        Console.WriteLine("aeTreeView not found name : ");
                        result = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        result = ConstCommon.TEST_PASS;
                        Console.WriteLine("aeTreeView found name : " + aeTreeView.Current.Name);
                    }
                }

                AutomationElement aeNLnode = null;
                if (result == ConstCommon.TEST_PASS)
                {
                    aeNLnode = TestTools.AUIUtilities.FindTreeViewNodeByName(testname, aeTreeView, "Mijn instellingen", ref sErrorMessage);
                    if (aeNLnode != null)
                    {
                        result = ConstCommon.TEST_PASS;
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    }
                    else
                    {
                        sErrorMessage = "Mijn instellingen not found  :";
                        Console.WriteLine(sErrorMessage);
                        Console.WriteLine("\nTest LanguageSetting: *FAIL*");
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    }
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALanguageSettingEventHandler);
            }
            #endregion

            #region setting language to English

            AutomationEventHandler UIALanguageSettingEventHandler2 = new AutomationEventHandler(OnLanguageSettingUIAEvent2);
            try
            {
                AutomationElement aeMySettings = AUICommon.FindTreeViewNodeLevel1(testname, root, "Mijn plek",
                    "Mijn instellingen", ref sErrorMessage);
                if (aeMySettings == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLanguageSetting window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALanguageSettingEventHandler2);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeMySettings);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement,
                     UIALanguageSettingEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                AutomationElement aeTreeView = null;
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }
                else
                {
                    string id = "m_TreeView";
                    DateTime sTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(root, ref aeTreeView, id, sTime, 60);

                    if (aeTreeView == null)
                    {
                        sErrorMessage = "aeTreeView not found name : ";
                        Console.WriteLine("aeTreeView not found name : ");
                        result = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        result = ConstCommon.TEST_PASS;
                        Console.WriteLine("aeTreeView found name : " + aeTreeView.Current.Name);
                    }
                }

                AutomationElement aeNLnode = null;
                if (result == ConstCommon.TEST_PASS)
                {
                    aeNLnode = TestTools.AUIUtilities.FindTreeViewNodeByName(testname, aeTreeView, "My settings", ref sErrorMessage);
                    if (aeNLnode != null)
                    {
                        result = ConstCommon.TEST_PASS;
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    }
                    else
                    {
                        sErrorMessage = "My settings not found  :";
                        Console.WriteLine(sErrorMessage);
                        Console.WriteLine("\nTest LanguageSetting: *FAIL*");
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    }
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIALanguageSettingEventHandler2);
            }
            #endregion
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ShellConfigSecurityWindowsUser
        public static void ShellConfigSecurityWindowsUser(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;


            #region // Config Security Windows user
            AutomationEventHandler UIAConfigSecurityEventHandler = new AutomationEventHandler(OnConfigSecurityWindowsUserUIAEvent);
            Thread.Sleep(5000);
            try
            {
                string shellID = "_MainForm_Toolbars_Dock_Area_Top";
                AutomationElement aeShell = AUIUtilities.FindElementByID(shellID, root);
                if (aeShell == null)
                {
                    sErrorMessage = shellID + "not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                double x = aeShell.Current.BoundingRectangle.Left + 5;
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom) / 2;
                Point shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Thread.Sleep(3000);

                // Add Open MyLanguageSetting window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIAConfigSecurityEventHandler);
                sEventEnd = false;

                double y2 = y + 15;
                Point securityPoint = new Point(x, y2);
                Input.MoveTo(securityPoint);
                Thread.Sleep(2000);

                Input.MoveToAndClick(securityPoint);
                Thread.Sleep(3000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIAConfigSecurityEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                Thread.Sleep(5000);
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }
                else
                    result = ConstCommon.TEST_PASS;

                // logon with current user

                if (result == ConstCommon.TEST_PASS)
                {
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIAConfigSecurityEventHandler);
            }
            #endregion
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LogonCurrentUser
        public static void LogonCurrentUser(string testname, AutomationElement root, out int result)
        {
            //TestTools.Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);


            Process[] ps = Process.GetProcessesByName(ConstCommon.EGEMIN_EPIA_SHELL);
            try
            {

                string shellID = "_MainForm_Toolbars_Dock_Area_Top";
                AutomationElement aeShell = AUIUtilities.FindElementByID(shellID, root);
                if (aeShell == null)
                {
                    sErrorMessage = shellID + "not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                double x = aeShell.Current.BoundingRectangle.Left + 5;
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom) / 2;
                Point shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);

                Console.WriteLine("aaaaaaaaaaaaaaaaaaaaaaaa ");
                Thread.Sleep(3000);

                double y2 = y + 40;
                Point securityPoint = new Point(x, y2);
                Input.MoveTo(securityPoint);

                Console.WriteLine("bbbbbbbbbbbbbbbbbbb ");
                Thread.Sleep(3000);

                Input.MoveToAndClick(securityPoint);
                Thread.Sleep(3000);

                Console.WriteLine("After log off Shell, wait 2 second : ");
                Thread.Sleep(2000);
                Console.WriteLine("=== Test " + testname + " ===");
                Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
                result = ConstCommon.TEST_UNDEFINED;
                AutomationElement aeSecurityForm = null;
                Console.WriteLine("Starting : ");
                Thread.Sleep(3000);

                AutomationEventHandler UIACurrentUserEventHandler = new AutomationEventHandler(OnUIACurrentUserEvent);
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIACurrentUserEventHandler);
                sEventEnd = false;
                TestCheck = ConstCommon.TEST_PASS;

                //string path = Path.Combine(sInstallScriptsDir, Constants.SHELL_BAT);
                string path = Path.Combine(m_SystemDrive + sEpiaClientRoot,
                    ConstCommon.EGEMIN_EPIA_SHELL_EXE);
                System.Diagnostics.Process proc = System.Diagnostics.Process.Start(path);
                Console.WriteLine("*****" + proc.Id);
                Thread.Sleep(9000);

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIACurrentUserEventHandler);

                Console.WriteLine("Application is started : ");
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;

                while (aeSecurityForm == null && mTime.TotalMilliseconds < 120000)
                {
                    aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", AutomationElement.RootElement);
                    Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                    mTime = DateTime.Now - mStartTime;
                }

                if (aeSecurityForm != null)
                {
                    Console.WriteLine("Find Application aeSecurityForm : " + System.DateTime.Now);
                }
                else
                {
                    sErrorMessage = "LogonForm not found";
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Console.WriteLine("Application aeSecurityForm name : " + aeSecurityForm.Current.Name);
                string UserNameID = "m_TextBoxUsername"; //"ControlType.Edit" Name : "with Windows credentials
                string PasswordID = "m_TextBoxPassword";
                string BtnOKID = "m_BtnOK";

                string origUser = string.Empty;
                string tester = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

                if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester, ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + UserNameID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser, UserPassword, ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + PasswordID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Logon into Application
                Thread.Sleep(3000);

                // Find Logon OK Button and click 
                if (AUIUtilities.FindElementAndClick(BtnOKID, aeSecurityForm))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindElementAndClick failed:" + BtnOKID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                AutomationElement.AutomationIdProperty, "MainForm", PropertyConditionFlags.IgnoreCase);

                Console.WriteLine(" find total mainForm :");

                // Find the element.
                AutomationElementCollection aes = AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                Thread.Sleep(10000);

                DateTime mAppTime = DateTime.Now;
                TimeSpan Time = DateTime.Now - mAppTime;
                while (aes.Count != 1 && Time.Minutes < 2)
                {
                    Console.WriteLine("Find Application aeForm : " + System.DateTime.Now);
                    aes = AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                    mTime = DateTime.Now - mAppTime;
                    Console.WriteLine(" find time is :" + Time.TotalMilliseconds / 1000);
                    Thread.Sleep(2000);
                }

                if (aes.Count == 1)
                {
                    result = ConstCommon.TEST_PASS;
                }
                else
                    result = ConstCommon.TEST_FAIL;

                Console.WriteLine(" total mainForm is :" + aes.Count);

                if (result == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\nTest Return Standard Screen: Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    sErrorMessage = "Number of mainForm should be 2, but now it is:" + aes.Count;
                    Console.WriteLine("\nTest scenario Return Standard Screen: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
        }
        #endregion LogonCurrentUser
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LogonEpiaAdministrator
        public static void LogonEpiaAdministrator(string testname, AutomationElement root, out int result)
        {
            TestTools.Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIAConfigSecurityEventHandler = new AutomationEventHandler(OnLogonEpiaAdminUIAEvent);

            Thread.Sleep(7000);
            try
            {
                string shellID = "_MainForm_Toolbars_Dock_Area_Top";
                AutomationElement aeShell = AUIUtilities.FindElementByID(shellID, root);
                if (aeShell == null)
                {
                    sErrorMessage = shellID + "not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                double x = aeShell.Current.BoundingRectangle.Left + 5;
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom) / 2;
                Point shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Thread.Sleep(3000);

                // Add Open MyLanguageSetting window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIAConfigSecurityEventHandler);
                sEventEnd = false;

                double y2 = y + 15;
                Point securityPoint = new Point(x, y2);
                Input.MoveTo(securityPoint);
                Thread.Sleep(2000);

                Input.MoveToAndClick(securityPoint);
                Thread.Sleep(3000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIAConfigSecurityEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                Thread.Sleep(5000);
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // first logoff than logon Administrator
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Thread.Sleep(3000);

                Point logoffPoint = new Point(shellPoint.X + 5, shellPoint.Y + 45);
                Input.MoveTo(logoffPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(logoffPoint);
                Thread.Sleep(3000);

                // logon with ikke
                Console.WriteLine("Application is started : ");

                DateTime mStartTime2 = DateTime.Now;
                TimeSpan mTime2 = DateTime.Now - mStartTime2;
                AutomationElement aeSecurityForm = null;
                while (aeSecurityForm == null && mTime2.TotalMilliseconds < 120000)
                {
                    aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", AutomationElement.RootElement);
                    Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime2.TotalMilliseconds);
                    mTime2 = DateTime.Now - mStartTime2;
                }

                if (aeSecurityForm != null)
                {
                    Console.WriteLine("Find Application aeSecurityForm : " + System.DateTime.Now);
                }
                else
                {
                    sErrorMessage = "LogonForm not found";
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Find Logon button and Select with Generic Credentials
                AutomationElement aeLogonButton = AUIUtilities.FindElementByID("m_MenuStrip", aeSecurityForm);
                if (aeLogonButton == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Failed find " + "Logon BUTTon " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(sErrorMessage);
                    return;
                }
                else
                {
                    Console.WriteLine("Logon button " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    Point LogonPoint = AUIUtilities.GetElementCenterPoint(aeLogonButton);
                    Thread.Sleep(1000);
                    Input.MoveToAndClick(LogonPoint);
                    Thread.Sleep(2000);
                    Point withGenericPoint = new Point(LogonPoint.X, LogonPoint.Y + 35);
                    Input.MoveTo(withGenericPoint);
                    Console.WriteLine("Generic");
                    Thread.Sleep(3000);
                    Input.MoveToAndClick(withGenericPoint);
                    Thread.Sleep(2000);
                }

                Console.WriteLine("Application aeSecurityForm name : " + aeSecurityForm.Current.Name);
                string UserNameID = "m_TextBoxUsername"; //"ControlType.Edit" Name : "with Windows credentials
                string PasswordID = "m_TextBoxPassword";
                string BtnOKID = "m_BtnOK";

                string origUser = string.Empty;
                string tester = "ikke";

                if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester, ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + UserNameID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser, "ikke", ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + PasswordID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Logon into Application
                Thread.Sleep(3000);

                // Find Logon OK Button and click 
                if (AUIUtilities.FindElementAndClick(BtnOKID, aeSecurityForm))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindElementAndClick failed:" + BtnOKID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                AutomationElement.AutomationIdProperty, "MainForm", PropertyConditionFlags.IgnoreCase);

                Console.WriteLine(" find total mainForm :");

                // Find the element.
                AutomationElementCollection aes = AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                Thread.Sleep(10000);

                DateTime mAppTime = DateTime.Now;
                TimeSpan Time = DateTime.Now - mAppTime;
                while (aes.Count != 1 && Time.Minutes < 2)
                {
                    Console.WriteLine("Find Application aeForm : " + System.DateTime.Now);
                    aes = AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                    mTime = DateTime.Now - mAppTime;
                    Console.WriteLine(" find time Time.TotalMinutes is :" + Time.TotalMinutes);
                    Console.WriteLine(" aes.Count :" + aes.Count);
                    Thread.Sleep(2000);
                }

                if (aes.Count == 1)
                {
                    result = ConstCommon.TEST_PASS;
                }
                else
                    result = ConstCommon.TEST_FAIL;

                Console.WriteLine(" total mainForm is :" + aes.Count);

                if (result == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\nTest Return Standard Screen: Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    sErrorMessage = "Number of mainForm should be 2, but now it is:" + aes.Count;
                    Console.WriteLine("\nTest scenario Return Standard Screen: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIAConfigSecurityEventHandler);
            }
        }
        #endregion LogonEpiaAdministrator
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ShellShutdown
        public static void ShellShutdown(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;

            try
            {
                System.Diagnostics.Process ShellProcess = null;
                int pID = TestTools.Utilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out ShellProcess);
                Console.WriteLine("Proc ID:" + pID);
                root = AutomationElement.FromHandle(ShellProcess.MainWindowHandle);

                if (root == null)
                {
                    Console.WriteLine("aeForm  not found : ");
                    return;
                }
                else
                    Console.WriteLine("aeForm found name : " + root.Current.Name);

                string shellID = "_MainForm_Toolbars_Dock_Area_Top";
                AutomationElement aeShell = AUIUtilities.FindElementByID(shellID, root);
                if (aeShell == null)
                {
                    sErrorMessage = shellID + "not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                double x = aeShell.Current.BoundingRectangle.Left + 5;
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom) / 2;
                Point shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Thread.Sleep(3000);

                //Point securityPoint = new Point(x, y + 90);
                Point securityPoint = new Point(x, y + 75);
                Input.MoveTo(securityPoint);

                Thread.Sleep(2000);
                Input.MoveToAndClick(securityPoint);

                Thread.Sleep(3000);

                Epia3Common.WriteTestLogMsg(slogFilePath, "Epia shutdown:", sOnlyUITest);
                // Check total number of main form
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                AutomationElement.AutomationIdProperty, "MainForm", PropertyConditionFlags.IgnoreCase);

                Console.WriteLine(" find total mainForm :");

                // Find the element.
                AutomationElementCollection aes = AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                Thread.Sleep(10000);

                DateTime mAppTime = DateTime.Now;
                TimeSpan Time = DateTime.Now - mAppTime;
                while (aes.Count != 0 && Time.Minutes < 2)
                {
                    Console.WriteLine("Find Application aeForm aes.Count: " + aes.Count);
                    aes = AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                    Time = DateTime.Now - mAppTime;
                    Console.WriteLine(" find time is :" + Time.TotalMilliseconds / 1000);
                    Thread.Sleep(2000);
                }

                if (aes.Count == 0)
                {
                    result = ConstCommon.TEST_PASS;
                }
                else
                    result = ConstCommon.TEST_FAIL;

                Console.WriteLine(" total mainForm is :" + aes.Count);

                if (result == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\nTest Shell Shutdown: Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    sErrorMessage = "Number of mainForm should be 1, but now it is:" + aes.Count;
                    Console.WriteLine("\nTest Test Shell Shutdown: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        #region ShellLogoff
        public static void ShellLogoff(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;

            Thread.Sleep(5000);

            try
            {
                AutomationEventHandler UIACurrentUserEventHandler = new AutomationEventHandler(OnUIACurrentUserEvent);
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIACurrentUserEventHandler);
                sEventEnd = false;
                TestCheck = ConstCommon.TEST_PASS;

                //string path = Path.Combine(sInstallScriptsDir, Constants.SHELL_BAT);
                string path = Path.Combine(m_SystemDrive + sEpiaClientRoot,
                    ConstCommon.EGEMIN_EPIA_SHELL_EXE);


                //System.Windows.Forms.MessageBox.Show("path : "+path);
                //System.Diagnostics.Process proc = System.Diagnostics.Process.Start(path);

                //path = Path.Combine(processDir, procFilename);
                Process proc = new Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = ConstCommon.EGEMIN_EPIA_SHELL_EXE;
                proc.StartInfo.Arguments = string.Empty;
                proc.StartInfo.WorkingDirectory = m_SystemDrive + sEpiaClientRoot;
                proc.Start();
                Thread.Sleep(5000);

                Console.WriteLine("*****" + proc.Id);
                Thread.Sleep(9000);

                // Start Shell
                //TestTools.Utilities.StartProcessNoWait(m_SystemDrive + sEpiaClientRoot,
                //    ConstCommon.EGEMIN_EPIA_SHELL_EXE, string.Empty);

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIACurrentUserEventHandler);

                Console.WriteLine("Application is started : ");
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;

                AutomationElement aeSecurityForm = null;
                while (aeSecurityForm == null && mTime.TotalMilliseconds < 120000)
                {
                    aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", AutomationElement.RootElement);
                    Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                    mTime = DateTime.Now - mStartTime;
                }

                if (aeSecurityForm != null)
                {
                    Console.WriteLine("Find Application aeSecurityForm : " + System.DateTime.Now);
                }
                else
                {
                    sErrorMessage = "LogonForm not found";
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Console.WriteLine("Application aeSecurityForm name : " + aeSecurityForm.Current.Name);
                string UserNameID = "m_TextBoxUsername"; //"ControlType.Edit" Name : "with Windows credentials
                string PasswordID = "m_TextBoxPassword";
                string BtnOKID = "m_BtnOK";

                string origUser = string.Empty;
                string tester = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

                if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester, ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + UserNameID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // input password
                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser, UserPassword, ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + PasswordID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Logon into Application
                Thread.Sleep(3000);

                // Find Logon OK Button and click 
                if (AUIUtilities.FindElementAndClick(BtnOKID, aeSecurityForm))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindElementAndClick failed:" + BtnOKID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Thread.Sleep(5000);
                AutomationElement aeMainForm = null;
                System.Diagnostics.Process ShellProcess = null;

                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                int pID = 0;
                while (pID == 0 && mTime.TotalMilliseconds < 120000)
                {
                    pID = TestTools.Utilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out ShellProcess);
                    Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                    mTime = DateTime.Now - mStartTime;
                }

                Console.WriteLine("Proc ID:" + pID);
                Thread.Sleep(3000);

                //aeMainForm = AutomationElement.FromHandle(ShellProcess.MainWindowHandle);

                string formID = "MainForm";
                DateTime mAppTime = DateTime.Now;
                AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeMainForm, formID, mAppTime, 300);
                if (aeMainForm == null)
                {
                    Console.WriteLine("aeForm  not found : ");
                    return;
                }
                else
                    Console.WriteLine("aeForm found name : " + aeMainForm.Current.Name);


                Thread.Sleep(13000);

                string shellID = "_MainForm_Toolbars_Dock_Area_Top";
                AutomationElement aeShell = AUIUtilities.FindElementByID(shellID, aeMainForm);
                if (aeShell == null)
                {
                    sErrorMessage = shellID + " not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                double x = aeShell.Current.BoundingRectangle.Left + 5;
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom) / 2;
                Point shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Thread.Sleep(3000);

                double y2 = y + 50;
                Point securityPoint = new Point(x, y2);
                Input.MoveTo(securityPoint);
                Thread.Sleep(5000);

                Input.MoveToAndClick(securityPoint);
                Thread.Sleep(3000);

                aeSecurityForm = null;
                DateTime mStartTime1 = DateTime.Now;
                mTime = DateTime.Now - mStartTime1;

                //====================
                while (aeSecurityForm == null && mTime.TotalMilliseconds < 120000)
                {
                    aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", AutomationElement.RootElement);
                    Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                    mTime = DateTime.Now - mStartTime1;
                }

                if (aeSecurityForm != null)
                {
                    Console.WriteLine("\nTest Shell LogOff: Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;

                }
                else
                {
                    sErrorMessage = "LogonForm not found";
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Console.WriteLine("Application aeSecurityForm name : " + aeSecurityForm.Current.Name);
                UserNameID = "m_TextBoxUsername"; //"ControlType.Edit" Name : "with Windows credentials
                PasswordID = "m_TextBoxPassword";
                BtnOKID = "m_BtnOK";

                origUser = string.Empty;
                tester = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

                if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester, ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + UserNameID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser, UserPassword, ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + PasswordID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Logon into Application
                Thread.Sleep(3000);

                // Find Logon OK Button and click 
                if (AUIUtilities.FindElementAndClick(BtnOKID, aeSecurityForm))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindElementAndClick failed:" + BtnOKID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
            }




            #region // Config Security Current Windows user
            AutomationEventHandler UIAConfigSecurityEventHandler = new AutomationEventHandler(OnConfigSecurityCurrentWindowsUserUIAEvent);
            Thread.Sleep(5000);


            try
            {

                aeForm = null;
                string formID = "MainForm";
                DateTime mAppTime = DateTime.Now;
                AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeForm, formID, mAppTime, 300);

                if (aeForm == null)
                {
                    AutomationElement aeError = AUIUtilities.FindElementByID("ErrorScreen", AutomationElement.RootElement);
                    if (aeError != null)
                        AUICommon.ErrorWindowHandling(aeError, ref sErrorMessage);
                    else
                        sErrorMessage = "Application Startup failed.";

                    throw new Exception(sErrorMessage);
                }
                else
                    Console.WriteLine("Application maeForm name : " + aeForm.Current.Name);

                // return to Current Windows User
                string shellID = "_MainForm_Toolbars_Dock_Area_Top";
                AutomationElement aeShell = AUIUtilities.FindElementByID(shellID, aeForm);
                if (aeShell == null)
                {
                    sErrorMessage = shellID + "not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                double x = aeShell.Current.BoundingRectangle.Left + 5;
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom) / 2;
                Point shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Thread.Sleep(3000);

                // Add Open MyLanguageSetting window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIAConfigSecurityEventHandler);
                sEventEnd = false;

                double y2 = y + 15;
                Point securityPoint = new Point(x, y2);
                Input.MoveTo(securityPoint);
                Thread.Sleep(2000);

                Input.MoveToAndClick(securityPoint);
                Thread.Sleep(3000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIAConfigSecurityEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                Thread.Sleep(5000);
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }
                else
                    result = ConstCommon.TEST_PASS;

                // logon with current user

                if (result == ConstCommon.TEST_PASS)
                {
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIAConfigSecurityEventHandler);
            }
            #endregion
        }
        #endregion


        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Epia3Close
        public static void Epia3Close(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            string BtnCloseID = "Close";
            try
            {
                Thread.Sleep(5000);
                // Close the other mainForms
                System.Diagnostics.Process[] pShell = System.Diagnostics.Process.GetProcessesByName(ConstCommon.EGEMIN_EPIA_SHELL);
                for (int i = 0; i < pShell.Length; i++)
                {
                    AutomationElement aeMainForm = AutomationElement.FromHandle(pShell[i].MainWindowHandle);
                    AutomationElement aeCloses = AUIUtilities.FindElementByID(BtnCloseID, aeMainForm);
                    if (aeCloses == null)
                    {
                        sErrorMessage = "failed to find aeCloses of mainForm";
                        Console.WriteLine(testname + " failed to find aeCloses at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                    else
                    {
                        Console.WriteLine(testname + " aeClose found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        InvokePattern ipc =
                            aeCloses.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        ipc.Invoke();
                    }

                    Thread.Sleep(5000);

                }

                // Close the other LogonForms
                System.Diagnostics.Process[] pLogon = System.Diagnostics.Process.GetProcessesByName(ConstCommon.EGEMIN_EPIA_SHELL);
                for (int i = 0; i < pLogon.Length; i++)
                {
                    AutomationElement aeLogonForm = AutomationElement.FromHandle(pLogon[i].MainWindowHandle);
                    AutomationElement aeCancel = AUIUtilities.FindElementByID("m_BtnCancel", aeLogonForm);
                    if (aeCancel == null)
                    {
                        sErrorMessage = "failed to find aeCancel of LogonForm";
                        Console.WriteLine(testname + " failed to find aeCloses at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                    else
                    {
                        Console.WriteLine(testname + " aeCancel found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        InvokePattern ipc =
                            aeCancel.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        ipc.Invoke();
                    }

                    Thread.Sleep(5000);

                }

                System.Diagnostics.Process proc = null;
                int pID = TestTools.Utilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out proc);
                Console.WriteLine("Proc ID:" + pID);
                if (pID == 0)
                {
                    Console.WriteLine("Epia3 Closed");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Epia3 Closed", sOnlyUITest);
                    Console.WriteLine("\nTest scenario: Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    Console.WriteLine("process id :" + pID);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "process id :" + pID, sOnlyUITest);
                    Console.WriteLine("\nTest scenario: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                }
                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
        }
        #endregion Epia3Close
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ReturnDefaultLayout
        public static void ReturnDefaultLayout(AutomationElement root, out int result)
        {
            string testname = "ReturnDefaultLayout";
            Console.WriteLine("\n=== Test ReturnDefaultLayout ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutStandardScreenEventHandler
                    = new AutomationEventHandler(OnLayoutStandardScreenUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutStandardScreenEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(2000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(2000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement,
                      UIALayoutStandardScreenEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    double left = root.Current.BoundingRectangle.Left;
                    double right = root.Current.BoundingRectangle.Right;
                    double top = root.Current.BoundingRectangle.Top;
                    double bottom = root.Current.BoundingRectangle.Bottom;

                    Console.WriteLine("Left=" + left);
                    //Console.WriteLine("right=" + right);
                    Console.WriteLine("top=" + top);
                    //Console.WriteLine("bottom=" + bottom);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Left=" + left, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "right=" + right, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "top=" + top, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "bottom=" + bottom, sOnlyUITest);

                    Epia3Common.WriteTestLogMsg(slogFilePath, "top=" + top + "--- " + "Left=" + left, sOnlyUITest);

                    /*if (left != 0 && top != 0)
                    {
                        Console.WriteLine("\nTest Return Standard Screen: Pass");
                        Epia3Common.WriteTestLogPass(slogFilePath,testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        Console.WriteLine("\nTest scenario Return Standard Screen: *FAIL*");
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath,testname, sOnlyUITest);
                    }
                    */
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                       UIALayoutStandardScreenEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ValidateDefaultLayout
        public static void ValidateDefaultLayout(AutomationElement root, out int result)
        {
            string testname = "ValidateDefaultLayout";
            Console.WriteLine("\n=== Test ValidateDefaultLayout ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutValidateDefaultEventHandler
                    = new AutomationEventHandler(OnLayoutValidateDefaultUIAEvent);

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                    ConstCommon.MY_LAYOUT, ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutValidateDefaultEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement,
                      UIALayoutValidateDefaultEventHandler);

                Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    double left = root.Current.BoundingRectangle.Left;
                    double right = root.Current.BoundingRectangle.Right;
                    double top = root.Current.BoundingRectangle.Top;
                    double bottom = root.Current.BoundingRectangle.Bottom;

                    Console.WriteLine("Left=" + left);
                    //Console.WriteLine("right=" + right);
                    Console.WriteLine("top=" + top);
                    //Console.WriteLine("bottom=" + bottom);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Left=" + left, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "right=" + right, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "top=" + top, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "bottom=" + bottom, sOnlyUITest);

                    Epia3Common.WriteTestLogMsg(slogFilePath, "top=" + top + "--- " + "Left=" + left, sOnlyUITest);

                    /*if (left != 0 && top != 0)
                    {
                        Console.WriteLine("\nTest Return Standard Screen: Pass");
                        Epia3Common.WriteTestLogPass(slogFilePath,testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        Console.WriteLine("\nTest scenario ValidateDefault Screen: *FAIL*");
                        result = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogFail(slogFilePath,testname, sOnlyUITest);
                    }
                    */
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                       UIALayoutValidateDefaultEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutMainMenuOn
        public static void LayoutMainMenuOn(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            string BtnApplyID = "m_BtnApply";
            string ChkRibbonID = "m_ChkShowRibbon";
            string RibbonID = "_MainForm_Toolbars_Dock_Area_Top";

            // mainmenu ID "_MainForm_Toolbars_Dock_Area_Top";

            try
            {
                // Get Ribbob Height Panel.
                AutomationElement aeRibbon
                    = AUIUtilities.FindElementByID(RibbonID, root);

                Console.WriteLine("before aeRibbon height:" + aeRibbon.Current.BoundingRectangle.Height);

                bool check = AUIUtilities.FindElementAndToggle(ChkRibbonID, root, ToggleState.Off);
                if (check)
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndToggle failed:" + ChkRibbonID);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Find Apply Button
                if (AUIUtilities.FindElementAndClick(BtnApplyID, root))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnApplyID);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                double ribbonHeight = aeRibbon.Current.BoundingRectangle.Height;
                Console.WriteLine("after aeRibbon height:" + ribbonHeight);

                if (ribbonHeight < 22)
                {
                    Console.WriteLine("aeRibbon is off");
                    Console.WriteLine("\nTest scenario Ribbon Off: Pass");
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    Console.WriteLine("aeRibbon not Off and height is:" + ribbonHeight);
                    Console.WriteLine("\nTest scenario Ribbon Off: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                }
                Thread.Sleep(3000);

                // return original window size
                Console.WriteLine("\nCheck Navigator CheckBox On ");
                AUIUtilities.FindElementAndToggle(ChkRibbonID, root, ToggleState.On);
                // Find and Click Apply Button
                if (AUIUtilities.FindElementAndClick(BtnApplyID, root))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnApplyID);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LayoutOpen
        public static void LayoutOpen(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            string BtnOpenID = "m_BtnOpen";
            TestCheck = ConstCommon.TEST_PASS;

            try
            {
                // Find Open Button Element
                System.Windows.Automation.Condition c = new PropertyCondition(
                AutomationElement.AutomationIdProperty, "m_BtnOpen", PropertyConditionFlags.IgnoreCase);

                // Find the element.
                AutomationElement aeOpenButton = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOpenButton == null)
                {
                    Console.WriteLine("FindElement Root OpenButton failed:");
                    Epia3Common.WriteTestLogFail(slogFilePath, "FindElement Root OpenButton failed:", sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open  Layout window Event Handler
                AutomationEventHandler UIALayoutOpenEventHandler = new AutomationEventHandler(OnLayoutOpenUIAEvent);
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutOpenEventHandler);

                // Find Open Button and click 
                if (AUIUtilities.FindElementAndClick(BtnOpenID, root))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                    Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndClick failed:" + BtnOpenID, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(testname + ":" + sErrorMessage);
                    Console.WriteLine(testname + " error msg is :" + sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + " error msg is :" + sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement,
                    UIALayoutOpenEventHandler);
                    return;
                }

                // Check Result
                string LbLLayoutID = "m_LbLLayoutId";
                AutomationElement aeLbLLayout = AUIUtilities.FindElementByID(LbLLayoutID, root);
                if (aeLbLLayout == null)
                {
                    Console.WriteLine("aeLbLLayout name empty:");
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                // validate Layout name 
                if (aeLbLLayout.Current.Name.Equals(sLayoutName))
                {
                    Console.WriteLine("\nTest scenario Layout Open: Pass");
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    Console.WriteLine("\nName should be " + sLayoutName + " , But it is: " + aeLbLLayout.Current.Name);
                    Console.WriteLine("\nTest scenario Layout Open: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement,
                    UIALayoutOpenEventHandler);

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ConfigOpen
        public static void ConfigOpen(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            string BtnOpenID = "m_BtnOpen";
            try
            {
                // Add Open  window Event Handler
                AutomationEventHandler UIAConfigOpenEventHandler = new AutomationEventHandler(OnConfigOpenUIAEvent);
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    root, TreeScope.Descendants, UIAConfigOpenEventHandler);

                // Find Open Button and click 
                if (AUIUtilities.FindElementAndClick(BtnOpenID, root))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent, root,
                   UIAConfigOpenEventHandler);

                // Check Result
                string LbLConfigurationID = "m_LbLConfigurationId";
                AutomationElement aeLbLConfiguration = AUIUtilities.FindElementByID(LbLConfigurationID, root);
                if (aeLbLConfiguration == null)
                {
                    Console.WriteLine("aeLbLConfiguration name empty:");
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                // validate Config name 
                if (aeLbLConfiguration.Current.Name.Equals(sConfigurationName))
                {
                    Console.WriteLine("\nTest scenario Config Open: Pass");
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    Console.WriteLine("\nName should be " + sConfigurationName + " , But it is: " + aeLbLConfiguration.Current.Name);
                    Console.WriteLine("\nTest scenario Open: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ConfigSave
        public static void ConfigSave(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            string BtnOpenID = "m_BtnOpen";
            try
            {
                AutomationElement aeGrid = AUIUtilities.FindElementByID("69632", root);
                if (aeGrid == null)
                {
                    Console.WriteLine("aeGrid not found:");
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                AutomationPattern[] ap = aeGrid.GetSupportedPatterns();
                for (int i = 0; i < ap.Length; i++)
                {
                    Console.WriteLine("ap[i].ProgrammaticName:" + ap[i].ProgrammaticName);
                    Console.WriteLine("ap[i].Id:" + ap[i].Id);
                    Console.WriteLine("ap[i].ToString():" + ap[i].ToString());
                }
                Thread.Sleep(3000000);

                // Add Open  window Event Handler
                AutomationEventHandler UIAConfigOpenEventHandler = new AutomationEventHandler(OnConfigOpenUIAEvent);
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    root, TreeScope.Descendants, UIAConfigOpenEventHandler);

                // Find Open Button and click 
                if (AUIUtilities.FindElementAndClick(BtnOpenID, root))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent, root,
                   UIAConfigOpenEventHandler);


                // Check Result
                string LbLConfigurationID = "m_LbLConfigurationId";
                AutomationElement aeLbLConfiguration = AUIUtilities.FindElementByID(LbLConfigurationID, root);
                if (aeLbLConfiguration == null)
                {
                    Console.WriteLine("aeLbLConfiguration name empty:");
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                // validate Config name 
                if (aeLbLConfiguration.Current.Name.Equals(sConfigurationName))
                {
                    Console.WriteLine("\nTest scenario Config Open: Pass");
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    Console.WriteLine("\nName should be " + sConfigurationName + " , ButôÎit is: " + aeLbLConfiguration.Current.Name);
                    Console.WriteLine("\nTest scenario Open: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ConfigSecurityUnsecured
        public static void ConfigSecurityUnsecured(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            string BtnOpenID = "m_BtnOpen";
            try
            {
                // Add Open  window Event Handler
                AutomationEventHandler UIAConfigSecurityUnsecuredEventHandler = new AutomationEventHandler(OnConfigSecurityUnsecuredUIAEvent);
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    root, TreeScope.Descendants, UIAConfigSecurityUnsecuredEventHandler);

                // Find Open Button and click 
                if (AUIUtilities.FindElementAndClick(BtnOpenID, root))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                //---------------------------------------------------------------------------------------------
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent, root,
                   UIAConfigSecurityUnsecuredEventHandler);

                WindowPattern wpCloseForm =
                 (WindowPattern)root.GetCurrentPattern(WindowPattern.Pattern);
                wpCloseForm.Close();
                Console.WriteLine("Application  Closing...: ");

                Thread.Sleep(3000);
                /*string  dir = @"C:\Epia 3\Epia\Main\Source\Presentation.CompositeUI.Shell\bin\Debug";
                string procFilename = "Egemin.Epia.Presentation.CompositeUI.Shell.exe";
                string procName = "Egemin.Epia.Presentation.CompositeUI.Shell";

                bool startApp = Utilities.StartProcessAndWaitUntilUIWindowFound(dir, procFilename,
                        null, procName, 2, ref aeForm);
			   
                if (!startApp)
                {
                    Console.WriteLine("Application not started or start failed : ");
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                */
                string LogonID = "LogonScreen";
                //string UserNameID   = "m_TxtUserName";
                //string PasswordID   = "m_TxtPassword";

                //Console.WriteLine("Application naeForm.Current.Name : " + AutomationElement.RootElement.Current.Name);
                // Check Result
                AutomationElement aeLogonScreen = AUIUtilities.FindElementByID(LogonID, AutomationElement.RootElement);
                if (aeLogonScreen == null)
                {
                    Console.WriteLine("aeLogonScreen not found, scenario Config UnsecuredSecurity: Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    Console.WriteLine("\nTest scenario Config UnsecuredSecurity: Fail");
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                // start the Application again
                Thread.Sleep(3000);

                string dir = m_SystemDrive + @"Program Files\Egemin\Epia Shell";
                string procShellExe = "Egemin.Epia.Shell.exe";
                string procName = "Egemin.Epia.Shell";
                bool startApp = TestTools.Utilities.StartProcessAndWaitUntilUIWindowFound(dir, procShellExe,
                        null, procName, 2, ref aeForm);

                if (!startApp)
                {
                    Console.WriteLine("Application not started or start failed : ");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Application not started or start failed : ", sOnlyUITest);
                    return;
                }

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ConfigFindGridView
        public static void ConfigFindGridView(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            string GridViewName = "m_PropertyGrid";
            //string GridViewName = "4852986";
            //string GridViewName = "m_BtnOpen";
            string BtnOpenID = "xx";
            try
            {
                AutomationElement agv = AUIUtilities.FindElementByID(GridViewName, root);
                if (agv == null)
                    Console.WriteLine("GRIDVIEW not found");
                else
                {
                    Console.WriteLine("GRIDVIEW is found");

                    AutomationElement table = AUIUtilities.FindElementByType(ControlType.Table, agv);
                    //AutomationElement table = AUIUtilities.FindElementByID(id, root);

                    if (table == null)
                        Console.WriteLine("Table not found");
                    else
                    {
                        Console.WriteLine("Table is found");
                        Console.WriteLine("Table name is " + table.Current.Name);
                        Console.WriteLine("Table ItemType is " + table.Current.ItemType);


                        AutomationPattern[] ps = table.GetSupportedPatterns();
                        Console.WriteLine("Table pattern length is " + ps.Length);

                        //TablePattern p = (TablePattern)table.GetCurrentPattern(TablePattern.Pattern);
                        //Console.WriteLine("Table is found Pattern is " + p.GetItem(1,1).Current.Name);
                    }

                    //InvokePattern patterns = (InvokePattern)agv.GetCurrentPattern(InvokePattern.Pattern);
                    //TablePattern patterns = (TablePattern)agv.GetCurrentPattern(TablePattern.Pattern);
                    GridPattern patterns = (GridPattern)agv.GetCurrentPattern(GridPattern.Pattern);

                    Console.WriteLine("current pattern is " + patterns.ToString());

                    //patterns.Invoke();
                    Thread.Sleep(300000);

                    TablePattern gP = (TablePattern)agv.GetCurrentPattern(TablePattern.Pattern);
                    gP.GetItem(1, 1).ToString();

                    Console.WriteLine("GRIDVIEW (1,1)is " + gP.GetItem(1, 1).ToString());

                }
                Thread.Sleep(300000);

                // Add Open  window Event Handler
                AutomationEventHandler UIAConfigSecurityEpiaEventHandler = new AutomationEventHandler(OnConfigSecurityEpiaUIAEvent);
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    root, TreeScope.Descendants, UIAConfigSecurityEpiaEventHandler);

                // Find Open Button and click 
                if (AUIUtilities.FindElementAndClick(BtnOpenID, root))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                //---------------------------------------------------------------------------------------------
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent, root,
                   UIAConfigSecurityEpiaEventHandler);

                WindowPattern wpCloseForm =
                 (WindowPattern)root.GetCurrentPattern(WindowPattern.Pattern);
                wpCloseForm.Close();
                Console.WriteLine("Application  Closing...: ");

                Thread.Sleep(3000);
                /*string  dir = @"C:\Epia 3\Epia\Main\Source\Presentation.CompositeUI.Shell\bin\Debug";
                string procFilename = "Egemin.Epia.Presentation.CompositeUI.Shell.exe";
                string procName = "Egemin.Epia.Presentation.CompositeUI.Shell";

                bool startApp = Utilities.StartProcessAndWaitUntilUIWindowFound(dir, procFilename,
                        null, procName, 2, ref aeForm);
			   
                if (!startApp)
                {
                    Console.WriteLine("Application not started or start failed : ");
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                */
                string LogonID = "LogonScreen";
                //string UserNameID     = "m_TxtUserName";
                //string PasswordID     = "m_TxtPassword";
                string BtnLogOnID = "m_BtnLogOn";
                string LogonScreenTitle = "Enter User Name and Password";

                //Console.WriteLine("Application naeForm.Current.Name : " + AutomationElement.RootElement.Current.Name);
                // Check Result
                AutomationElement aeLogonScreen = AUIUtilities.FindElementByID(LogonID, AutomationElement.RootElement);
                if (aeLogonScreen == null)
                {
                    Console.WriteLine("aeLogonScreen not found:");
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                // validate Logon Screen
                if (aeLogonScreen.Current.Name.Equals(LogonScreenTitle))
                {
                    Console.WriteLine("\nTest scenario Config EpiaSecurity: Pass");
                    result = ConstCommon.TEST_PASS;
                }
                else
                {
                    Console.WriteLine("\nName should be " + LogonScreenTitle + " , But it is: " + aeLogonScreen.Current.Name);
                    Console.WriteLine("\nTest scenario Open: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                }
                // Logon into Application
                Thread.Sleep(3000);

                // Find Logon Button and click 
                if (AUIUtilities.FindElementAndClick(BtnLogOnID, aeLogonScreen))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnLogOnID);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(testname + "ôÎ=== " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
        }
        #endregion ConfigFindGridView
        #endregion Test Cases ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Event ------------------------------------------------------------------------------------------------
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnUIAServerEvent
        public static void OnUIAServerEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnUIAServerEvent");
            TestCheck = ConstCommon.TEST_PASS;
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnUIAServerEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);
            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, "shell start exception: " + sErrorMessage, sOnlyUITest);
                Thread.Sleep(6000);
            }
            else if (name.Equals("<NoName>"))
            {
                Console.WriteLine("Do nothing Name is ------------:" + name);
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Text);

                // Find the element.
                AutomationElement err = element.FindFirst(TreeScope.Element | TreeScope.Descendants, c);

                if (err != null)
                {
                    sErrorMessage = err.Current.Name;
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("shell start exception: " + sErrorMessage);
                    //Epia3Common.WriteTestLogFail(slogFilePath, "shell start exception: " + sErrorMessage, sOnlyUITest);
                    Thread.Sleep(6000);
                }
                else
                {
                    Console.WriteLine("Not a Error window ------------:" + name);
                    return;
                }
            }
            else if (name.Equals(m_SystemDrive + "WINDOWS\\system32\\cmd.exe"))
            {
                Console.WriteLine("SERVER OOOOOOOOOOOO Name is ------------:" + name);
                Console.WriteLine("SERVER OOOOOOOOOOOO Name is ------------:" + name);
                Console.WriteLine("SERVER OOOOOOOOOOOO Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "SERVER open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("ThemeManagerNotification"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("Epia security"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("Egemin e'pia User Interface Shell"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                Thread.Sleep(3000);
            }
            else if (name.Equals("Egemin Shell"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                Thread.Sleep(3000);
            }
            else if (name.Equals("Open File - Security Warning"))
            {
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                System.Windows.Automation.Condition c = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Run"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                );

                // Find the element.
                AutomationElement aeRun = element.FindFirst(TreeScope.Element | TreeScope.Descendants, c);

                if (aeRun != null)
                {
                    Point pt = AUIUtilities.GetElementCenterPoint(aeRun);
                    Input.MoveTo(pt);
                    Thread.Sleep(1000);
                    Input.ClickAtPoint(pt);
                }
                else
                {
                    Console.WriteLine("Run Button not Found ------------:" + name);
                    return;
                }
            }
            else
            {
                Console.WriteLine("SERVER Do Other Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "SERVER open other window name: " + name, sOnlyUITest);
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnUIAShellEvent
        public static void OnUIAShellEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnUIAShellEvent");
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnUIAShellEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(2000);
            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, "shell start exception: " + sErrorMessage, sOnlyUITest);
                Thread.Sleep(6000);
            }
            else if (name.Equals("<NoName>"))
            {
                Console.WriteLine("Do nothing Name is ------------:" + name);
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Text);

                // Find the element.
                AutomationElement err = element.FindFirst(TreeScope.Element | TreeScope.Descendants, c);

                if (err != null)
                {
                    sErrorMessage = err.Current.Name;
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("shell start exception: " + sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, "shell start exception: " + sErrorMessage, sOnlyUITest);
                }
                else
                {
                    Console.WriteLine("Not a Error window ------------:" + name);
                    return;
                }

            }
            else if (name.Equals(m_SystemDrive + "WINDOWS\\system32\\cmd.exe"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("ThemeManagerNotification"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("Epia security"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("Egemin e'pia User Interface Shell"))
            {
                Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                Thread.Sleep(3000);
            }
            else if (name.Equals("Egemin Shell"))
            {
                if (element.Current.AutomationId.Equals("ErrorScreen"))
                {
                    AutomationElement aeBtn = AUIUtilities.FindElementByID("m_BtnDetails", element);
                    if (aeBtn != null)
                    {
                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeBtn));
                        Thread.Sleep(3000);
                        AutomationElement aeTxt = AUIUtilities.FindElementByID("m_TxtDetails", element);
                        if (aeTxt != null)
                        {
                            TextPattern tp = (TextPattern)aeTxt.GetCurrentPattern(TextPattern.Pattern);
                            Thread.Sleep(1000);
                            sErrorMessage = tp.DocumentRange.GetText(-1).Trim();
                            Console.WriteLine("Error Message Catched ------------:");
                            //Console.WriteLine("Error Message is ------------:" + sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, "start shell failed: " + sErrorMessage, sOnlyUITest);
                        }
                        else
                        {
                            Console.WriteLine("Error Message not found ------------:");
                            Epia3Common.WriteTestLogFail(slogFilePath, "Error Message pane not found: ", sOnlyUITest);
                        }
                    }

                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {
                    Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                    Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                    Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                    Thread.Sleep(3000);
                }
            }
            else if (name.Equals("Open File - Security Warning"))
            {
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                System.Windows.Automation.Condition c = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Run"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                );

                // Find the element.
                AutomationElement aeRun = element.FindFirst(TreeScope.Element | TreeScope.Descendants, c);

                if (aeRun != null)
                {
                    if (aeRun != null)
                    {
                        Point pt = AUIUtilities.GetElementCenterPoint(aeRun);
                        Input.MoveTo(pt);
                        Thread.Sleep(1000);
                        Input.ClickAtPoint(pt);
                    }
                }
                else
                {
                    Console.WriteLine("Run Button not Found ------------:" + name);
                    return;
                }
            }
            else
            {
                Console.WriteLine("Do ELSE OTHER is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLicenseConnectionUIAEvent
        public static void OnLicenseConnectionUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLicenseConnectionUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen"))
            {
                Console.WriteLine("OnLicenseConnectionUIAEvent: AutomationId=" + element.Current.AutomationId);
                sLicWindowOpen = true;
            }
            else if (element.Current.AutomationId.StartsWith("ErrorScreen"))
            {
                Console.WriteLine("OnLicenseConnectionUIAEvent: AutomationId=" + element.Current.AutomationId);
            }
            else
                return;

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnLicenseConnectionUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);
            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
            }
            else if (name.Equals("<NoName>"))
            {
                Console.WriteLine("Do nothing Name is ------------:" + name);
            }
            else
            {
                string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                if (element == null)
                {
                    Console.WriteLine("element is null");
                }
                else
                {
                    Console.WriteLine("element is not null name is " + element.Current.AutomationId);
                }
                Console.WriteLine("Retry registration ------------:" + BtnBtnRetryRegistrationId);
                Thread.Sleep(9000);


                AutomationElement aeBtnRetry = AUIUtilities.FindElementByID(BtnBtnRetryRegistrationId, element);

                if (aeBtnRetry != null)
                    AUIUtilities.ClickElement(aeBtnRetry);

                Thread.Sleep(1000);
                if (aeBtnRetry != null)
                    AUIUtilities.ClickElement(aeBtnRetry);

                Thread.Sleep(1000);
                //AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, element);

                Thread.Sleep(1000);
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnFindLayoutPanelUIAEvent
        public static void OnFindLayoutPanelUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnFindLayoutPanelUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }


            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            Console.WriteLine("OnFindLayoutPanelUIAEvent1");

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnFindLayoutPanelUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);
            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
            }
            else if (name.Equals("<NoName>"))
            {
                Console.WriteLine("Do nothing Name is ------------:" + name);
            }
            else
            {
                string autoID = "tableLayoutPanel4";
                string BtnCancelID = "m_BtnCancel";
                try
                {
                    // find shell layout details pane
                    AutomationElement aePanel4 = AUIUtilities.FindElementByID(autoID, element);
                    if (aePanel4 == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "My Layout pane not found";
                        Console.WriteLine("My layout pane not found:");
                    }
                    else
                    {
                        Console.WriteLine("Text Name is :" + aePanel4.Current.ToString());
                        TestCheck = ConstCommon.TEST_PASS;
                    }

                    Thread.Sleep(1000);
                    string text = string.Empty;
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        // find Shell Layout Details Text
                        AutomationElement aeLabel1 = AUIUtilities.FindElementByID("m_lblDescription", aePanel4);
                        if (aeLabel1 == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            Console.WriteLine("My layout Text not found");
                            sErrorMessage = "My layout Text not found";
                        }
                        else
                        {
                            Console.WriteLine("aeLabel1 Name is :" + aeLabel1.Current.Name);
                            text = aeLabel1.Current.Name;
                            TestCheck = ConstCommon.TEST_PASS;
                        }
                    }

                    Thread.Sleep(1000);
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        if (text.Equals("Id:"))
                        {
                            Console.WriteLine("My Layout Text found and is correct");
                            TestCheck = ConstCommon.TEST_PASS;
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            Console.WriteLine("Shell Layout Details Text not correct");
                            sErrorMessage = "Shell Layout Details Text not correct:" + text;
                        }
                    }

                    Thread.Sleep(1000);
                    AUIUtilities.FindElementAndClick(BtnCancelID, element);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("OnFindLayoutPanelUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutXPosUIAEvent
        public static void OnLayoutXPosUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutXPosUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }


            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LayoutXPos:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string InitalXPositionTextBoxID = "initialXPositionTextBox";
                string BtnSaveID = "m_btnSave";
                try
                {
                    string origXPos = "";
                    // Change XPos TxtBox
                    string getValue = string.Empty;
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalXPositionTextBoxID, element, out origXPos, "200", ref sErrorMessage))
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalXPositionTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalXPositionTextBoxID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    Thread.Sleep(3000);
                    #region //save
                    // Find and Click Save Button
                    Console.WriteLine("Save button ------------:" + BtnSaveID);
                    AutomationElement aeBtnSave = AUIUtilities.FindElementByID(BtnSaveID, element);
                    if (aeBtnSave != null)
                    {
                        AUIUtilities.ClickElement(aeBtnSave);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("OnLayoutXPosUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutYPosUIAEvent
        public static void OnLayoutYPosUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutYPosUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LayoutYPos:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string InitalYPositionTextBoxID = "initialYPositionTextBox";
                string BtnSaveID = "m_btnSave";
                try
                {
                    string origYPos = "";
                    // Change YPos TxtBox
                    string getValue = string.Empty;
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalYPositionTextBoxID, element, out origYPos, "100", ref sErrorMessage))
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalYPositionTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalYPositionTextBoxID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    Thread.Sleep(3000);
                    #region //save
                    // Find and Click Save Button
                    Console.WriteLine("Save button ------------:" + BtnSaveID);
                    AutomationElement aeBtnSave = AUIUtilities.FindElementByID(BtnSaveID, element);
                    if (aeBtnSave != null)
                    {
                        AUIUtilities.ClickElement(aeBtnSave);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("OnLayoutYPosUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }


            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutWidthUIAEvent
        public static void OnLayoutWidthUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutWidthUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnLayoutWidthUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);
            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string InitalWidthTextBoxID = "initialWidthTextBox";
                string BtnSaveID = "m_btnSave";
                try
                {
                    string origWidth = "";
                    // Change YPos TxtBox
                    string getValue = string.Empty;
                    // Change Width TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalWidthTextBoxID, element, out origWidth, "600", ref sErrorMessage))
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalWidthTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalWidthTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    Thread.Sleep(3000);
                    #region //save
                    // Find and Click Save Button
                    Console.WriteLine("Save button ------------:" + BtnSaveID);
                    AutomationElement aeBtnSave = AUIUtilities.FindElementByID(BtnSaveID, element);
                    if (aeBtnSave != null)
                    {
                        AUIUtilities.ClickElement(aeBtnSave);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("OnLayoutWidthUIAEvent:" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutHeightUIAEvent
        public static void OnLayoutHeightUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutHeightUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnLayoutHeightUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string InitalHeightTextBoxID = "initialHeightTextBox";
                string BtnSaveID = "m_btnSave";
                try
                {
                    string origWidth = "";
                    // Change YPos TxtBox
                    string getValue = string.Empty;
                    // Change Width TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalHeightTextBoxID, element, out origWidth, "500", ref sErrorMessage))
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalHeightTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalHeightTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    Thread.Sleep(3000);
                    #region //save
                    // Find and Click Save Button
                    Console.WriteLine("Save button ------------:" + BtnSaveID);
                    AutomationElement aeBtnSave = AUIUtilities.FindElementByID(BtnSaveID, element);
                    if (aeBtnSave != null)
                    {
                        AUIUtilities.ClickElement(aeBtnSave);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("OnLayoutHeightUIAEvent:" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutTitleUIAEvent
        public static void OnLayoutTitleUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutTitleUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnLayoutTitleUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string titleTextBoxID = "titleTextBox";
                string BtnSaveID = "m_btnSave";
                try
                {
                    string origTitle = "";
                    // Change YPos TxtBox
                    string getValue = string.Empty;
                    // Change Title TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(titleTextBoxID, element, out origTitle, "abcdefg", ref sErrorMessage))
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + titleTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + titleTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    Thread.Sleep(3000);
                    // Find and Click Save Button
                    if (AUIUtilities.FindElementAndClickPoint(BtnSaveID, element))
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                }
                catch (Exception ex)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("OnLayoutHeightUIAEvent:" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutResizeUIAEvent
        public static void OnLayoutResizeUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutResizeUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LayoutResize:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string ChkResizeID = "allowResizeCheckBox";
                string BtnSaveID = "m_btnSave";
                try
                {
                    bool check = AUIUtilities.FindElementAndToggle(ChkResizeID, element, ToggleState.On);
                    if (check)
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkResizeID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkResizeID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkResizeID;
                        sEventEnd = true;
                        return;
                    }

                    Thread.Sleep(3000);
                    #region //save
                    // Find and Click Save Button
                    Console.WriteLine("Save button ------------:" + BtnSaveID);
                    AutomationElement aeBtnSave = AUIUtilities.FindElementByID(BtnSaveID, element);
                    if (aeBtnSave != null)
                    {
                        AUIUtilities.ClickElement(aeBtnSave);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("OnLayoutFullScreenUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutFullScreenUIAEvent
        public static void OnLayoutFullScreenUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutFullScreenUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LayoutFullScreen:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {   // Automation Element ID
                string ChkFullScreenID = "fullScreenCheckBox";
                string BtnSaveID = "m_btnSave";
                try
                {
                    bool check = AUIUtilities.FindElementAndToggle("fullScreenCheckBox", element, ToggleState.On);
                    if (check)
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkFullScreenID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkFullScreenID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkFullScreenID;
                        sEventEnd = true;
                        return;
                    }

                    Thread.Sleep(3000);
                    #region //save
                    // Find and Click Save Button
                    Console.WriteLine("Save button ------------:" + BtnSaveID);
                    AutomationElement aeBtnSave = AUIUtilities.FindElementByID(BtnSaveID, element);
                    if (aeBtnSave != null)
                    {
                        AUIUtilities.ClickElement(aeBtnSave);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("OnLayoutFullScreenUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutMaximizedScreenUIAEvent
        public static void OnLayoutMaximizedScreenUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutMaximizedScreenUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LayoutMaximizedScreen:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {   // Automation Element ID
                string ChkFullScreenID = "fullScreenCheckBox";
                string BtnSaveID = "m_btnSave";
                try
                {
                    bool check = AUIUtilities.FindElementAndToggle(ChkFullScreenID, element, ToggleState.Off);
                    if (check)
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkFullScreenID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkFullScreenID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkFullScreenID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkMaximizedScreenID = "maximizedCheckBox";
                    bool checkMs = AUIUtilities.FindElementAndToggle(ChkMaximizedScreenID, element, ToggleState.On);
                    if (checkMs)
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkMaximizedScreenID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkMaximizedScreenID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkMaximizedScreenID;
                        sEventEnd = true;
                        return;
                    }

                    Thread.Sleep(3000);
                    #region //save
                    // Find and Click Save Button
                    Console.WriteLine("Save button ------------:" + BtnSaveID);
                    AutomationElement aeBtnSave = AUIUtilities.FindElementByID(BtnSaveID, element);
                    if (aeBtnSave != null)
                    {
                        AUIUtilities.ClickElement(aeBtnSave);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("OnLayoutFullScreenUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutRibbonOnUIAEvent
        public static void OnLayoutRibbonOnUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutRibbonOnUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LayoutRibbonOn:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {   // Automation Element ID   
                string BtnSaveID = "m_btnSave";
                try
                {
                    string ChkShowRibbonID = "showRibbonCheckBox"; //-------------- Ribbon ON -----------------------
                    bool checkRibbon = AUIUtilities.FindElementAndToggle(ChkShowRibbonID, element, ToggleState.On);
                    if (checkRibbon)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowRibbonID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowRibbonID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowRibbonID;
                        sEventEnd = true;
                        return;
                    }

                    Thread.Sleep(3000);
                    #region //save
                    // Find and Click Save Button
                    Console.WriteLine("Save button ------------:" + BtnSaveID);
                    AutomationElement aeBtnSave = AUIUtilities.FindElementByID(BtnSaveID, element);
                    if (aeBtnSave != null)
                    {
                        AUIUtilities.ClickElement(aeBtnSave);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("OnLayoutRibbonOnUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutNavigatorOffUIAEvent
        public static void OnLayoutNavigatorOffUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutNavigatorOffUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LayoutNavigatorOff:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {   // Automation Element ID   
                string BtnSaveID = "m_btnSave";
                try
                {
                    string ChkShowNavigatorID = "showNavigatorCheckBox"; //-------------- Navigator Off-------------------------
                    bool checkNav = AUIUtilities.FindElementAndToggle(ChkShowNavigatorID, element, ToggleState.Off);
                    if (checkNav)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowNavigatorID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowNavigatorID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowNavigatorID;
                        sEventEnd = true;
                        return;
                    }


                    Thread.Sleep(3000);
                    // Find and Click Save Button
                    if (AUIUtilities.FindElementAndClick(BtnSaveID, element))
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                }
                catch (Exception ex)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("OnLayoutNavigatorOffUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutCancelButtonUIAEvent
        public static void OnLayoutCancelButtonUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutCancelButtonUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }
            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnLayoutCancelButtonUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string BtnCancelID = "m_btnCancel";
                try
                {
                    string ChkFullScreenID = "fullScreenCheckBox"; // -----------------FullScreen ON -----------------
                    bool check = AUIUtilities.FindElementAndToggle("fullScreenCheckBox", element, ToggleState.On);
                    if (check)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkFullScreenID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkFullScreenID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkFullScreenID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkMaximizedID = "maximizedCheckBox"; // ------------------Maximized On-----------------
                    bool checkMax = AUIUtilities.FindElementAndToggle(ChkMaximizedID, element, ToggleState.On);
                    if (checkMax)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkMaximizedID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkMaximizedID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkMaximizedID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkResizeID = "allowResizeCheckBox"; // ------------------Allow resize Off-----------------
                    bool checkAr = AUIUtilities.FindElementAndToggle(ChkResizeID, element, ToggleState.Off);
                    if (checkAr)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkResizeID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkResizeID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkResizeID;
                        sEventEnd = true;
                        return;
                    }

                    string InitalXPositionTextBoxID = "initialXPositionTextBox"; // ------------------XPos 200-----------------
                    string origXPos = "";
                    // Change XPos TxtBox
                    string getValue = string.Empty;
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalXPositionTextBoxID, element, out origXPos, "200", ref sErrorMessage))
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalXPositionTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalXPositionTextBoxID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalYPositionTextBoxID = "initialYPositionTextBox"; // ------------------YPos 100-----------------
                    string origYPos = "";
                    // Change YPos TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalYPositionTextBoxID, element, out origYPos, "100", ref sErrorMessage))
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalYPositionTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalYPositionTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalWidthTextBoxID = "initialWidthTextBox"; // ------------------Width 600-----------------
                    string origWidth = "";
                    // Change Width TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalWidthTextBoxID, element, out origWidth, "600", ref sErrorMessage))
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalWidthTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalWidthTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalHeightTextBoxID = "initialHeightTextBox"; // ------------------Height 700-----------------
                    string origHeight = "";
                    // Change Height TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalHeightTextBoxID, element, out origHeight, "700", ref sErrorMessage))
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalHeightTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalHeightTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    /*string TitleTextBoxID = "titleTextBox"; //-----------Title  "Egemin Shell"-----------------------------
                    string origTitle = "";
                    // Change Height TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(TitleTextBoxID, element, out origTitle, "ButtonCancel", ref sErrorMessage))
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + TitleTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + TitleTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }*/

                    string ChkShowRibbonID = "showRibbonCheckBox"; //-------------- Ribbon On ---------------------------
                    bool checkRibbon = AUIUtilities.FindElementAndToggle(ChkShowRibbonID, element, ToggleState.On);
                    if (checkRibbon)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowRibbonID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowRibbonID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowRibbonID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowMainMenuID = "showMainMenuCheckBox"; //-------------- Main Menu Off---------------------------
                    bool checkMm = AUIUtilities.FindElementAndToggle(ChkShowMainMenuID, element, ToggleState.Off);
                    if (checkMm)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowMainMenuID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowMainMenuID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowMainMenuID;
                        sEventEnd = true;
                        return;
                    }

                    /*string ChkShowToolBarsID = "showToolBarsCheckBox"; //-------------- Tool bars On---------------------------
                    bool checktb = AUIUtilities.FindElementAndToggle(ChkShowToolBarsID, element, ToggleState.On);
                    if (checktb)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowToolBarsID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowToolBarsID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowToolBarsID;
                        sEventEnd = true;
                        return;
                    }*/

                    string ChkShowNavigatorID = "showNavigatorCheckBox"; //-------------- Navigator Off-------------------------
                    bool checkNav = AUIUtilities.FindElementAndToggle(ChkShowNavigatorID, element, ToggleState.Off);
                    if (checkNav)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowNavigatorID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowNavigatorID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowNavigatorID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkStackScreensID = "stackScreensCheckBox"; //-------------- Stack Screens On-------------------------
                    bool checkSs = AUIUtilities.FindElementAndToggle(ChkStackScreensID, element, ToggleState.On);
                    if (checkSs)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkStackScreensID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkStackScreensID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkStackScreensID;
                        sEventEnd = true;
                        return;
                    }

                    // Find and Click Cancel Button
                    #region //Cancel
                    // Find and Click Cancel Button
                    Console.WriteLine("Cancel button ------------:" + BtnCancelID);
                    AutomationElement aeBtnSave = AUIUtilities.FindElementByID(BtnCancelID, element);
                    if (aeBtnSave != null)
                    {
                        AUIUtilities.ClickElement(aeBtnSave);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnCancelID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    sErrorMessage = "OnLayoutCancelButtonUIAEvent" + ex.Message + " --- " + ex.StackTrace;
                    Console.WriteLine("OnLayoutCancelButtonUIAEvent" + sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLanguageSettingUIAEvent
        public static void OnLanguageSettingUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLanguageSettingUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.UserSettingsDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LanguageSetting:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string BtnSaveID = "m_btnSave";
                string BtnCancelID = "m_btnCancel";
                Console.WriteLine("Finding Language combo");
                try
                {
                    AutomationElement aeCombo = AUIUtilities.FindElementByID("languageIdComboBox", element);
                    if (aeCombo == null)
                    {
                        Console.WriteLine("LanguageSettings failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        sErrorMessage = "LanguageSettings failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    bool select = false; //Utilities.SelectItemFromList("nl", aeCombo);
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        SelectionPattern selectPattern =
                           aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                        AutomationElement item
                            = AUIUtilities.FindElementByName("Nederlands", aeCombo);
                        if (item != null)
                        {
                            Console.WriteLine("LanguageSettings item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(2000);

                            SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            itemPattern.Select();
                            select = true;
                        }
                        else
                        {
                            Console.WriteLine("Finding Language item nl failed");
                            sErrorMessage = "Finding Language item nl failed";
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }

                        if (!select)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            Console.WriteLine("Finding Language combo nl failed");
                            sErrorMessage = "Finding Language combo nl failed";
                        }
                    }

                    Thread.Sleep(3000);
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        if (AUIUtilities.FindElementAndClickPoint(BtnSaveID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                            sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                    else
                    {
                        if (AUIUtilities.FindElementAndClickPoint(BtnCancelID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                            sErrorMessage = "FindElementAndClick failed:" + BtnCancelID;
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("OnLanguageSettingUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion

        #region OnLanguageSettingUIAEvent2
        public static void OnLanguageSettingUIAEvent2(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLanguageSettingUIAEvent2");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.UserSettingsDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }
            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LanguageSetting:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string BtnSaveID = "m_btnSave";
                string BtnCancelID = "m_btnCancel";
                Console.WriteLine("Finding Language combo");
                try
                {
                    AutomationElement aeCombo = AUIUtilities.FindElementByID("languageIdComboBox", element);
                    if (aeCombo == null)
                    {
                        Console.WriteLine("LanguageSettings failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        sErrorMessage = "LanguageSettings failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    bool select = false; //Utilities.SelectItemFromList("nl", aeCombo);
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        SelectionPattern selectPattern =
                           aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                        AutomationElement item
                            = AUIUtilities.FindElementByName("English", aeCombo);
                        if (item != null)
                        {
                            Console.WriteLine("LanguageSettings item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(2000);

                            SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            itemPattern.Select();
                            select = true;
                        }
                        else
                        {
                            Console.WriteLine("Finding Language item nl failed");
                            sErrorMessage = "Finding Language item nl failed";
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }

                        if (!select)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            Console.WriteLine("Finding Language combo nl failed");
                            sErrorMessage = "Finding Language combo nl failed";
                        }
                    }

                    Thread.Sleep(3000);
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        if (AUIUtilities.FindElementAndClickPoint(BtnSaveID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                            sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                    else
                    {
                        if (AUIUtilities.FindElementAndClickPoint(BtnCancelID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                            sErrorMessage = "FindElementAndClick failed:" + BtnCancelID;
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("OnLanguageSettingUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnConfigSecurityWindowsUserUIAEvent
        public static void OnConfigSecurityWindowsUserUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnConfigSecurityWindowsUserUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellConfigurationDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }


            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnConfigSecurityWindowsUserUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string ComboSecurityModesID = "m_ComboSecurityModes";
                string BtnSaveID = "m_btnSave";
                string BtnCancelID = "m_btnCancel";
                Console.WriteLine("Finding Logon mode combo");
                try
                {
                    AutomationElement aeCombo = AUIUtilities.FindElementByID(ComboSecurityModesID, element);
                    if (aeCombo == null)
                    {
                        Console.WriteLine("ConfigSecurity failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        sErrorMessage = "ConfigSecurity failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    else
                    {
                        Input.MoveTo(aeCombo);
                        Thread.Sleep(2000);
                    }

                    bool select = false;
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        SelectionPattern selectPattern =
                           aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                        AutomationElement item
                            = AUIUtilities.FindElementByName("Windows user", aeCombo);
                        if (item != null)
                        {
                            Console.WriteLine("ConfigSecurity item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(2000);

                            SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            itemPattern.Select();
                            select = true;
                        }
                        else
                        {
                            Console.WriteLine("Finding ConfigSecurity item failed");
                            sErrorMessage = "Finding ConfigSecurity item nl failed";
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }

                        if (!select)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            Console.WriteLine("Finding ConfigSecurity combo failed");
                            sErrorMessage = "Finding ConfigSecurity combo failed";
                        }
                    }

                    Thread.Sleep(3000);
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        if (AUIUtilities.FindElementAndClickPoint(BtnSaveID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                            sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                            TestCheck = ConstCommon.TEST_FAIL;
                            Thread.Sleep(5000);
                            sEventEnd = true;
                            return;
                        }
                    }
                    else
                    {
                        if (AUIUtilities.FindElementAndClickPoint(BtnCancelID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                            sErrorMessage = "FindElementAndClick failed:" + BtnCancelID;
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("OnConfigSecurityWindowsUserUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #endregion
        #region OnConfigSecurityCurrentWindowsUserUIAEvent
        public static void OnConfigSecurityCurrentWindowsUserUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnConfigSecurityCurrentWindowsUserUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellConfigurationDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnConfigSecurityCurrentWindowsUserUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string ComboSecurityModesID = "m_ComboSecurityModes";
                string BtnSaveID = "m_btnSave";
                string BtnCancelID = "m_btnCancel";
                Console.WriteLine("Finding logon mode combo");
                try
                {
                    AutomationElement aeCombo = AUIUtilities.FindElementByID(ComboSecurityModesID, element);
                    if (aeCombo == null)
                    {
                        Console.WriteLine("ConfigSecurity failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        sErrorMessage = "ConfigSecurity failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    else
                    {
                        Input.MoveTo(aeCombo);
                        Thread.Sleep(2000);
                    }

                    bool select = false;
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        SelectionPattern selectPattern =
                           aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                        AutomationElement item
                            = AUIUtilities.FindElementByName("Current windows user", aeCombo);
                        if (item != null)
                        {
                            Console.WriteLine("ConfigSecurity item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(2000);

                            SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            itemPattern.Select();
                            select = true;
                        }
                        else
                        {
                            Console.WriteLine("Finding ConfigSecurity item failed");
                            sErrorMessage = "Finding ConfigSecurity item nl failed";
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }

                        if (!select)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            Console.WriteLine("Finding ConfigSecurity combo failed");
                            sErrorMessage = "Finding ConfigSecurity combo failed";
                        }
                    }

                    Thread.Sleep(3000);
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        if (AUIUtilities.FindElementAndClickPoint(BtnSaveID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                            sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                            TestCheck = ConstCommon.TEST_FAIL;
                            Thread.Sleep(5000);
                            sEventEnd = true;
                            return;
                        }
                    }
                    else
                    {
                        if (AUIUtilities.FindElementAndClickPoint(BtnCancelID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                            sErrorMessage = "FindElementAndClick failed:" + BtnCancelID;
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("OnConfigSecurityCurrentWindowsUserUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutStandardScreenUIAEvent
        public static void OnLayoutStandardScreenUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutStandardScreenUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }
            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnLayoutStandardScreenUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string BtnSaveID = "m_btnSave";
                try
                {
                    string ChkFullScreenID = "fullScreenCheckBox"; // -----------------FullScreen OFF -----------------
                    bool check = AUIUtilities.FindElementAndToggle("fullScreenCheckBox", element, ToggleState.Off);
                    if (check)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkFullScreenID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkFullScreenID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkFullScreenID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkMaximizedID = "maximizedCheckBox"; // ------------------Maximized OFF-----------------
                    bool checkMax = AUIUtilities.FindElementAndToggle(ChkMaximizedID, element, ToggleState.Off);
                    if (checkMax)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkMaximizedID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkMaximizedID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkMaximizedID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkResizeID = "allowResizeCheckBox"; // ------------------Allow resize ON-----------------
                    bool checkAr = AUIUtilities.FindElementAndToggle(ChkResizeID, element, ToggleState.On);
                    if (checkAr)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkResizeID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkResizeID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkResizeID;
                        sEventEnd = true;
                        return;
                    }

                    string InitalXPositionTextBoxID = "initialXPositionTextBox"; // ------------------XPos 0-----------------
                    string origXPos = "";
                    // Change XPos TxtBox
                    string getValue = string.Empty;
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalXPositionTextBoxID,
                            element, out origXPos, DEFAULT_XPOS, ref sErrorMessage))
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalXPositionTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalXPositionTextBoxID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalYPositionTextBoxID = "initialYPositionTextBox"; // ------------------YPos 0-----------------
                    string origYPos = "";
                    // Change YPos TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalYPositionTextBoxID,
                            element, out origYPos, DEFAULT_YPOS, ref sErrorMessage))
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalYPositionTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalYPositionTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalWidthTextBoxID = "initialWidthTextBox"; // ------------------Width 792-----------------
                    string origWidth = "";
                    // Change Width TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalWidthTextBoxID,
                            element, out origWidth, DEFAULT_WIDTH, ref sErrorMessage))
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalWidthTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalWidthTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalHeightTextBoxID = "initialHeightTextBox"; // ------------------Height 606-----------------
                    string origHeight = "";
                    // Change Height TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalHeightTextBoxID,
                            element, out origHeight, DEFAULT_HEIGHT, ref sErrorMessage))
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalHeightTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalHeightTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    /*string TitleTextBoxID = "titleTextBox"; //-----------Title  "Egemin Shell"-----------------------------
                    string origTitle = "";
                    // Change Height TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(TitleTextBoxID,
                            element, out origTitle, DEFAULT_TITLE, ref sErrorMessage))
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + TitleTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + TitleTextBoxID;
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }*/

                    string ChkShowRibbonID = "showRibbonCheckBox"; //-------------- Ribbon OFF ---------------------------
                    bool checkRibbon = AUIUtilities.FindElementAndToggle(ChkShowRibbonID, element, ToggleState.Off);
                    if (checkRibbon)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowRibbonID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowRibbonID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowRibbonID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowMainMenuID = "showMainMenuCheckBox"; //-------------- Main Menu ON---------------------------
                    bool checkMm = AUIUtilities.FindElementAndToggle(ChkShowMainMenuID, element, ToggleState.On);
                    if (checkMm)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowMainMenuID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowMainMenuID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowMainMenuID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowToolBarsID = "showToolBarsCheckBox"; //-------------- Tool bars OFF---------------------------
                    bool checktb = AUIUtilities.FindElementAndToggle(ChkShowToolBarsID, element, ToggleState.Off);
                    if (checktb)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowToolBarsID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowToolBarsID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowToolBarsID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowNavigatorID = "showNavigatorCheckBox"; //-------------- Navigator ON-------------------------
                    bool checkNav = AUIUtilities.FindElementAndToggle(ChkShowNavigatorID, element, ToggleState.On);
                    if (checkNav)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowNavigatorID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowNavigatorID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowNavigatorID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkStackScreensID = "stackScreensCheckBox"; //-------------- Stack Screens OFF-------------------------
                    bool checkSs = AUIUtilities.FindElementAndToggle(ChkStackScreensID, element, ToggleState.Off);
                    if (checkSs)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkStackScreensID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkStackScreensID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkStackScreensID;
                        sEventEnd = true;
                        return;
                    }

                    // Find and Click Save Button
                    #region //save
                    // Find and Click Save Button
                    Console.WriteLine("Save button ------------:" + BtnSaveID);
                    AutomationElement aeBtnSave = AUIUtilities.FindElementByID(BtnSaveID, element);
                    if (aeBtnSave != null)
                    {
                        AUIUtilities.ClickElement(aeBtnSave);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    sErrorMessage = "OnLayoutStandardScreenUIAEvent" + ex.Message + " --- " + ex.StackTrace;
                    Console.WriteLine("OnLayoutStandardScreenUIAEvent" + ex.Message + " --- " + ex.StackTrace);
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutValidateDefaultUIAEvent
        public static void OnLayoutValidateDefaultUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutValidateDefaultUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
            else
            {
                return;
            }

            while (element.Current.IsEnabled == false)
            {
                //ppp
                string DialogID = "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen";
                AutomationElement aeLicDialog = AUIUtilities.FindElementByID(DialogID, AutomationElement.RootElement);
                if (aeLicDialog == null)
                {
                    Console.WriteLine(DialogID + " failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    break;
                }
                else
                {
                    Console.WriteLine("Retry registration ------------:" + name);
                    Thread.Sleep(5000);
                    string BtnBtnRetryRegistrationId = "m_BtnRetryRegistration";
                    AUIUtilities.FindElementAndClickPoint(BtnBtnRetryRegistrationId, aeLicDialog);
                    Thread.Sleep(5000);
                }
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnLayoutValidateDefaultUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string BtnCancelID = "m_btnCancel";
                try
                {
                    string ChkFullScreenID = "fullScreenCheckBox"; // -----------------FullScreen OFF -----------------
                    ToggleState fstg = AUIUtilities.FindCheckBoxAndToggleState(ChkFullScreenID, element, ref sErrorMessage);
                    if (fstg == DEFAULT_FULLSCREEN)
                        Console.WriteLine("OK for :" + ChkFullScreenID);
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string ChkMaximizedID = "maximizedCheckBox"; // ------------------Maximized OFF-----------------
                    ToggleState mstg = AUIUtilities.FindCheckBoxAndToggleState(ChkMaximizedID, element, ref sErrorMessage);
                    if (mstg == DEFAULT_MAXIMIZED)
                        Console.WriteLine("OK for :" + ChkMaximizedID);
                    else
                    {
                        sErrorMessage = ChkMaximizedID + ": is unfound or state is ON";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string ChkResizeID = "allowResizeCheckBox"; // ------------------Allow resize ON-----------------
                    ToggleState artg = AUIUtilities.FindCheckBoxAndToggleState(ChkResizeID, element, ref sErrorMessage);
                    if (artg == DEFAULT_ALLOWRESIZE)
                        Console.WriteLine("OK for :" + ChkResizeID);
                    else
                    {
                        sErrorMessage = ChkResizeID + ": is unfound or state is Off";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalXPositionTextBoxID = "initialXPositionTextBox"; // ------------------XPos 0-----------------
                    string xpos = AUIUtilities.FindTextBoxAndValue(InitalXPositionTextBoxID, element, ref sErrorMessage);
                    if (xpos != null)
                    {
                        Console.WriteLine("OK for :" + InitalXPositionTextBoxID);
                        if (!xpos.Equals(DEFAULT_XPOS))
                        {
                            sErrorMessage = "Xpos not correct";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalYPositionTextBoxID = "initialYPositionTextBox"; // ------------------YPos 0-----------------
                    string ypos = AUIUtilities.FindTextBoxAndValue(InitalYPositionTextBoxID, element, ref sErrorMessage);
                    if (ypos != null)
                    {
                        Console.WriteLine("OK for :" + InitalYPositionTextBoxID);
                        if (!ypos.Equals(DEFAULT_YPOS))
                        {
                            sErrorMessage = "Ypos not correct";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalWidthTextBoxID = "initialWidthTextBox"; // ------------------Width 792-----------------
                    string width = AUIUtilities.FindTextBoxAndValue(InitalWidthTextBoxID, element, ref sErrorMessage);
                    if (width != null)
                    {
                        Console.WriteLine("OK for :" + InitalWidthTextBoxID);
                        if (!width.Equals(DEFAULT_WIDTH))
                        {
                            sErrorMessage = "Width not correct";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalHeightTextBoxID = "initialHeightTextBox"; // ------------------Height 606-----------------
                    string height = AUIUtilities.FindTextBoxAndValue(InitalHeightTextBoxID, element, ref sErrorMessage);
                    if (height != null)
                    {
                        Console.WriteLine("OK for :" + InitalHeightTextBoxID);
                        if (!height.Equals(DEFAULT_HEIGHT))
                        {
                            sErrorMessage = "Height not correct";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    /*string TitleTextBoxID = "titleTextBox"; //-----------Title  "Egemin Shell"-----------------------------
                    string title = AUIUtilities.FindTextBoxAndValue(TitleTextBoxID, element, ref sErrorMessage);
                    if (height != null)
                    {
                        Console.WriteLine("OK for :" + TitleTextBoxID);
                        if (!title.Equals(DEFAULT_TITLE))
                        {
                            sErrorMessage = "title not correct";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }*/

                    string ChkShowRibbonID = "showRibbonCheckBox"; //-------------- Ribbon OFF ---------------------------
                    ToggleState rbtg = AUIUtilities.FindCheckBoxAndToggleState(ChkFullScreenID, element, ref sErrorMessage);
                    if (rbtg == DEFAULT_SHOW_RIBBON)
                        Console.WriteLine("OK for :" + ChkShowRibbonID);
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowMainMenuID = "showMainMenuCheckBox"; //-------------- Main Menu ON---------------------------
                    ToggleState mmtg = AUIUtilities.FindCheckBoxAndToggleState(ChkShowMainMenuID, element, ref sErrorMessage);
                    if (mmtg == DEFAULT_SHOW_MAINENU)
                        Console.WriteLine("OK for :" + ChkShowMainMenuID);
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowToolBarsID = "showToolBarsCheckBox"; //-------------- Tool bars OFF---------------------------
                    ToggleState tbtg = AUIUtilities.FindCheckBoxAndToggleState(ChkShowToolBarsID, element, ref sErrorMessage);
                    if (tbtg == DEFAULT_SHOW_TOOLBARS)
                        Console.WriteLine("OK for :" + ChkShowToolBarsID);
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowNavigatorID = "showNavigatorCheckBox"; //-------------- Navigator ON-------------------------
                    ToggleState nvtg = AUIUtilities.FindCheckBoxAndToggleState(ChkShowNavigatorID, element, ref sErrorMessage);
                    if (nvtg == DEFAULT_SHOW_NAVIGATOR)
                        Console.WriteLine("OK for :" + ChkShowNavigatorID);
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string ChkStackScreensID = "stackScreensCheckBox"; //-------------- Stack Screens OFF-------------------------
                    ToggleState sstg = AUIUtilities.FindCheckBoxAndToggleState(ChkStackScreensID, element, ref sErrorMessage);
                    if (sstg == DEFAULT_STACK_SCREENS)
                        Console.WriteLine("OK for :" + ChkStackScreensID);
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    // Find and Click Cancel Button
                    #region //Cancel
                    // Find and Click Cancel Button
                    Console.WriteLine("Save button ------------:" + BtnCancelID);
                    AutomationElement aeBtnCancel = AUIUtilities.FindElementByID(BtnCancelID, element);
                    if (aeBtnCancel != null)
                    {
                        AUIUtilities.ClickElement(aeBtnCancel);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnCancelID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    sErrorMessage = "OnLayoutValidateDefaultUIAEvent" + ex.Message + " --- " + ex.StackTrace;
                    Console.WriteLine("OnLayoutValidateDefaultUIAEvent" + ex.Message + " --- " + ex.StackTrace);
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnUIACurrentUserEvent
        public static void OnUIACurrentUserEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnUIACurrentUserEvent");
            TestCheck = ConstCommon.TEST_PASS;
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnUIACurrentUserEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);
            if (name.Equals("Open File - Security Warning"))
            {
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                System.Windows.Automation.Condition c = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Run"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                );

                // Find the element.
                AutomationElement aeRun = element.FindFirst(TreeScope.Element | TreeScope.Descendants, c);

                if (aeRun != null)
                {
                    Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeRun));
                }
                else
                {
                    Console.WriteLine("Run Button not Found ------------:" + name);
                    return;
                }
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLogonEpiaAdminUIAEvent
        public static void OnLogonEpiaAdminUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLogonEpiaAdminUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnLogonEpiaAdminUIAEventt:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);

            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Thread.Sleep(5000);
            }
            else
            {
                // Automation Element ID
                string ComboSecurityModesID = "m_ComboSecurityModes";
                string BtnSaveID = "m_btnSave";
                string BtnCancelID = "m_btnCancel";
                Console.WriteLine("Finding Security mode combo");
                try
                {
                    AutomationElement aeCombo = AUIUtilities.FindElementByID(ComboSecurityModesID, element);
                    DateTime mStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - mStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                    while (aeCombo == null && mTime.Seconds <= 120)
                    {
                        Thread.Sleep(2000);
                        mTime = DateTime.Now - mStartTime;
                        Console.WriteLine("find time is:" + mTime.TotalMilliseconds);
                        aeCombo = AUIUtilities.FindElementByID(ComboSecurityModesID, element);
                    }

                    if (aeCombo == null)
                    {
                        Console.WriteLine("ConfigSecurity failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        sErrorMessage = "ConfigSecurity failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    else
                    {
                        Input.MoveTo(aeCombo);
                        Thread.Sleep(2000);
                    }

                    bool select = false;
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        SelectionPattern selectPattern =
                           aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                        AutomationElement item
                            = AUIUtilities.FindElementByName("EpiaMemberOrAnyWindowsUser", aeCombo);
                        if (item != null)
                        {
                            Console.WriteLine("ConfigSecurity item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(2000);

                            SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            itemPattern.Select();
                            select = true;
                        }
                        else
                        {
                            Console.WriteLine("Finding ConfigSecurity item failed");
                            sErrorMessage = "Finding ConfigSecurity item nl failed";
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }

                        if (!select)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            Console.WriteLine("Finding ConfigSecurity combo failed");
                            sErrorMessage = "Finding ConfigSecurity combo failed";
                        }
                    }

                    Thread.Sleep(3000);
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        if (AUIUtilities.FindElementAndClickPoint(BtnSaveID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                            sErrorMessage = "FindElementAndClick failed:" + BtnSaveID;
                            TestCheck = ConstCommon.TEST_FAIL;
                            Thread.Sleep(5000);
                            sEventEnd = true;
                            return;
                        }
                    }
                    else
                    {
                        if (AUIUtilities.FindElementAndClickPoint(BtnCancelID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                            sErrorMessage = "FindElementAndClick failed:" + BtnCancelID;
                            TestCheck = ConstCommon.TEST_FAIL;
                            sEventEnd = true;
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("OnLogonEpiaAdminUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }
        #endregion


        //&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutOpenUIAEvent
        public static void OnLayoutOpenUIAEvent(object src, AutomationEventArgs args)
        {
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;

                string name = "";
                if (element == null)
                {
                    name = "null";
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Open Window name is null";
                    return;
                }
                else
                {
                    name = element.GetCurrentPropertyValue(
                        AutomationElement.NameProperty) as string;
                }

                if (element.Current.AutomationId.StartsWith("Dialog - Egemin.Epia.Modules.RnD.Screens.ShellLayoutDetailsScreen"))
                    Console.WriteLine("element.Current.AutomationId=" + element.Current.AutomationId);
                else
                {
                    return;
                }

                if (name.Length == 0) name = "<NoName>";
                string str = string.Format("LayoutOpen:={0} : {1}", name, args.EventId.ProgrammaticName);
                Console.WriteLine(str);

                if (name.Equals("Open Shell Layout"))
                {
                    string BtnOpenID = "m_BtnOpen";
                    string BtnCancelID = "m_BtnCancel";
                    string ListLayoutID = "m_ListLayoutIds";

                    // Select Shell LayoutID
                    AutomationElement aeLauoutIDs
                        = AUIUtilities.FindElementByID(ListLayoutID, element);
                    if (aeLauoutIDs == null)
                    {
                        Console.WriteLine("FindElementByID failed:" + ListLayoutID);
                    }

                    aeLauoutIDs.SetFocus();

                    // find first listitem
                    System.Windows.Automation.Condition cF = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                        new PropertyCondition(AutomationElement.IsControlElementProperty, true));

                    // Find all children that match the specified conditions.
                    AutomationElementCollection foundCollection = aeLauoutIDs.FindAll(TreeScope.Children, cF);
                    if (foundCollection == null)
                    {
                        Console.WriteLine("Layout is empty:");
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Layout is empty";

                        Thread.Sleep(5000);

                        // Find Cancel Button and click
                        if (AUIUtilities.FindElementAndClick(BtnCancelID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                        }
                        return;

                    }
                    else if (foundCollection.Count == 0)
                    {
                        Console.WriteLine("Layout field is empty:");
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Layout field is empty";

                        Thread.Sleep(5000);

                        // Find Cancel Button and click
                        if (AUIUtilities.FindElementAndClick(BtnCancelID, element))
                            Thread.Sleep(3000);
                        else
                        {
                            Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                        }
                        return;

                    }
                    Console.WriteLine("Layout is not empty:");

                    // Get Opened Config Name
                    string openLayoutName = string.Empty;
                    AutomationElement aeindex = foundCollection[0];
                    if (aeindex != null)
                    {
                        sLayoutName = aeindex.Current.Name;
                        openLayoutName = aeindex.Current.Name;
                        Console.WriteLine("open layout name :" + sLayoutName);
                        SelectionItemPattern sip = (SelectionItemPattern)aeindex.GetCurrentPattern(SelectionItemPattern.Pattern);
                        sip.Select();
                    }

                    Rect rect;
                    rect = (Rect)(aeindex.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
                    System.Windows.Point pt = new System.Windows.Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
                    Input.MoveToAndClick(pt);

                    // Find Open Button and click
                    if (AUIUtilities.FindElementAndClick(BtnOpenID, element))
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                    }
                }
                else if (name.Equals("Error"))
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("Name is ------------:" + name);
                    AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                }
                else
                    Console.WriteLine("LayoutOpen:Name is ------------:" + name);

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("OnLayoutOpenUIAEvent:" + ex.Message + "  --   " + ex.StackTrace);
                return;
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutSaveUIAEvent
        public static void OnLayoutSaveUIAEvent(object src, AutomationEventArgs args)
        {
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;

                string name = "";
                if (element == null)
                {
                    name = "null";
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Open Window name is null";
                    return;
                }
                else
                {
                    name = element.GetCurrentPropertyValue(
                        AutomationElement.NameProperty) as string;
                }

                if (name.Length == 0) name = "<NoName>";
                string str = string.Format("LayoutSave:={0} : {1}", name, args.EventId.ProgrammaticName);
                Console.WriteLine(str);

                if (name.Equals("Open Shell Layout"))
                {
                    string BtnOpenID = "m_BtnOpen";
                    string ListLayoutID = "m_ListLayoutIds";

                    // Select Shell LayoutID
                    AutomationElement aeLauoutIDs
                        = AUIUtilities.FindElementByID(ListLayoutID, element);
                    if (aeLauoutIDs == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Console.WriteLine("FindElementByID failed:" + ListLayoutID);
                        sErrorMessage = "FindElementByID failed:" + ListLayoutID;
                        return;
                    }

                    // find first listitem
                    System.Windows.Automation.Condition cF = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                        new PropertyCondition(AutomationElement.IsControlElementProperty, true));

                    // Find all children that match the specified conditions.
                    AutomationElementCollection foundCollection = aeLauoutIDs.FindAll(TreeScope.Children, cF);
                    if (foundCollection == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Console.WriteLine("Layout is empty:");
                        sErrorMessage = "Layout is empty:";
                        return;
                    }

                    int index = -1;
                    //reopen saved layout
                    for (int i = 0; i < foundCollection.Count; i++)
                    {
                        if (foundCollection[i].Current.Name.EndsWith(sLayoutName))
                        {
                            index = i;
                            break;
                        }
                    }

                    if (index == -1)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Console.WriteLine("No saved layout found:" + sLayoutName);
                        sErrorMessage = "No saved Layout Found:" + sLayoutName;
                        return;
                    }

                    // Get Opened Config Name
                    string openLayoutName = string.Empty;
                    AutomationElement aeindex = foundCollection[index];
                    if (aeindex != null)
                    {
                        sLayoutName = aeindex.Current.Name;
                        openLayoutName = aeindex.Current.Name;
                        Console.WriteLine("open layout name :" + sLayoutName);
                        SelectionItemPattern sip = (SelectionItemPattern)aeindex.GetCurrentPattern(SelectionItemPattern.Pattern);
                        sip.Select();
                    }

                    Rect rect;
                    rect = (Rect)(aeindex.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
                    System.Windows.Point pt = new System.Windows.Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
                    Input.MoveToAndClick(pt);

                    // Find Open Button and click
                    if (AUIUtilities.FindElementAndClick(BtnOpenID, element))
                        Thread.Sleep(3000);
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnOpenID;
                        return;
                    }
                }
                else if (name.Equals("Error"))
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine("Name is ------------:" + name);
                    AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                }
                else
                    Console.WriteLine("Name is ------------:" + name);
            }
            catch (Exception ex)
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine(ex.Message);
                Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + " --- " + ex.StackTrace, sOnlyUITest);
                sErrorMessage = ex.Message;
                return;
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnLayoutSaveAsUIAEvent
        public static void OnLayoutSaveAsUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnLayoutSaveAsUIAEvent");
            TestCheck = ConstCommon.TEST_PASS;
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LayoutSaveAs:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            if (name.Equals("Save Shell Layout"))
            {
                string BtnSaveID = "m_BtnSave";
                string TxtLayoutID = "m_TxtShellLayoutId";
                /*
                 // Find and Click SaveAs Button
                if (AUIUtilities.FindElementAndClick(BtnCancelID, element))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                }
                */
                string getValue = "x";
                // Change Shell ConfigID
                if (AUIUtilities.FindTextBoxAndChangeValue(TxtLayoutID, element, out getValue, "TestLayoutSaveAs@EPIA3TESTPC", ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindTextBoxAndChangeValue failed:" + TxtLayoutID);
                    Epia3Common.WriteTestLogFail(slogFilePath, "FindTextBoxAndChangeValue failed:" + TxtLayoutID + " --- " + sErrorMessage, sOnlyUITest);
                }

                // Find and Click Save Button
                if (AUIUtilities.FindElementAndClick(BtnSaveID, element))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                }

            }
            else if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
            }
            else
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnScreenOpenedUIAEvent
        public static void OnScreenOpenedUIAEvent(object src, AutomationEventArgs args)
        {
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("yyy={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnConfigSaveAsUIAEvent
        public static void OnConfigSaveAsUIAEvent(object src, AutomationEventArgs args)
        {
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch (Exception ex)
            {
                sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("ConfigSaveAs:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            if (name.Equals("Save Shell Configuration"))
            {
                string BtnSaveID = "m_BtnSave";
                string TxtConfigurationID = "m_TxtShellConfigurationID";
                /*
                 // Find and Click SaveAs Button
                if (Utilities.FindElementAndClick(BtnCancelID, element))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                }
                */
                string getValue = "x";
                // Change Shell ConfigID
                if (AUIUtilities.FindTextBoxAndChangeValue(TxtConfigurationID, element, out getValue, "TestConfigSaveAs", ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindTextBoxAndChangeValue failed:" + TxtConfigurationID + " --- " + sErrorMessage);
                }

                // Find and Click Save Button
                if (AUIUtilities.FindElementAndClick(BtnSaveID, element))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnSaveID);
                }

            }
            else
                Console.WriteLine("Name is ------------:" + name);
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnConfigOpenUIAEvent
        public static void OnConfigOpenUIAEvent(object src, AutomationEventArgs args)
        {
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("ConfigOpen:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            if (name.Equals("Open Shell Configuration"))
            {
                string BtnOpenID = "m_BtnOpen";
                string ListConfigurationID = "m_ListConfigurationIds";

                // Select Shell ConfigID
                AutomationElement aeLauoutIDs
                    = AUIUtilities.FindElementByID(ListConfigurationID, element);
                if (aeLauoutIDs == null)
                {
                    Console.WriteLine("FindElementByID failed:" + ListConfigurationID);
                }

                aeLauoutIDs.SetFocus();

                // find first listitem
                System.Windows.Automation.Condition cF = new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                    new PropertyCondition(AutomationElement.IsControlElementProperty, true));

                // Find all children that match the specified conditions.
                AutomationElementCollection foundCollection = aeLauoutIDs.FindAll(TreeScope.Children, cF);
                if (foundCollection == null)
                {
                    Console.WriteLine("Layout is empty:");
                }

                // Get Opened Config Name
                string openLayoutName = string.Empty;
                AutomationElement aeindex = foundCollection[1];
                if (aeindex != null)
                {
                    sConfigurationName = aeindex.Current.Name;
                    openLayoutName = aeindex.Current.Name;
                    Console.WriteLine("open layout name :" + sConfigurationName);
                    SelectionItemPattern sip = (SelectionItemPattern)aeindex.GetCurrentPattern(SelectionItemPattern.Pattern);
                    sip.Select();
                }

                Rect rect;
                rect = (Rect)(aeindex.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
                System.Windows.Point pt = new System.Windows.Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
                Input.MoveToAndClick(pt);

                // Find Open Button and click
                if (AUIUtilities.FindElementAndClick(BtnOpenID, element))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                }
            }
            else
                Console.WriteLine("Name is ------------:" + name);
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnConfigSecurityEpiaUIAEvent
        public static void OnConfigSecurityEpiaUIAEvent(object src, AutomationEventArgs args)
        {
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("ConfigSecurityEpia:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            try
            {
                if (name.Equals("Open Shell Configuration"))
                {
                    string BtnOpenID = "m_BtnOpen";
                    string ListConfigurationID = "m_ListConfigurationIds";

                    // Select Shell ConfigID
                    AutomationElement aeLauoutIDs
                        = AUIUtilities.FindElementByID(ListConfigurationID, element);
                    if (aeLauoutIDs == null)
                    {
                        Console.WriteLine("FindElementByID failed:" + ListConfigurationID);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "FindElementByID failed:" + ListConfigurationID, sOnlyUITest);
                    }
                    //aeLauoutIDs.SetFocus();

                    // find all listitem
                    System.Windows.Automation.Condition cF = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                        new PropertyCondition(AutomationElement.IsControlElementProperty, true));

                    // Find all children that match the specified conditions.
                    AutomationElementCollection foundCollection = aeLauoutIDs.FindAll(TreeScope.Children, cF);
                    if (foundCollection == null)
                    {
                        Console.WriteLine("Shell Configuration is empty:");
                    }

                    int foundIdx = -1;
                    for (int i = 0; i < foundCollection.Count; i++)
                    {
                        Console.WriteLine("Shell Configuration: :" + foundCollection[i].Current.Name);
                        string configName = foundCollection[i].Current.Name;
                        if (configName.Equals("TestConfigEpiaSecurity"))
                        {
                            foundIdx = i;
                            break;
                        }
                    }

                    if (foundIdx < 0)
                    {
                        Console.WriteLine("Shell SecurityEpia Configuration not found :");
                        return;
                    }

                    // Get Opened ConfigSecurityEpia
                    AutomationElement aeindex = foundCollection[foundIdx];
                    if (aeindex != null)
                    {
                        sConfigurationName = aeindex.Current.Name;
                        Console.WriteLine("open layout name :" + sConfigurationName);
                        SelectionItemPattern sip = (SelectionItemPattern)aeindex.GetCurrentPattern(SelectionItemPattern.Pattern);
                        sip.Select();
                    }

                    Rect rect;
                    rect = (Rect)(aeindex.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
                    System.Windows.Point pt = new System.Windows.Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
                    Input.MoveToAndClick(pt);

                    // Find Open Button and click
                    if (AUIUtilities.FindElementAndClick(BtnOpenID, element))
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                    }
                }
                else
                    Console.WriteLine("Name is ------------:" + name);
            }
            catch (Exception ex)
            {
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnConfigSecurityEpiaUIAEvent exception:" + ex.Message + "----" + ex.StackTrace, sOnlyUITest);
                Console.WriteLine("OnConfigSecurityEpiaUIAEvent exception:" + ex.Message + "----" + ex.StackTrace);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnConfigSecurityUnsecuredUIAEvent
        public static void OnConfigSecurityUnsecuredUIAEvent(object src, AutomationEventArgs args)
        {
            string testcase = sTestCaseName[Counter];
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("ConfigSecurityEpia:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            try
            {
                if (name.Equals("Open Shell Configuration"))
                {
                    string BtnOpenID = "m_BtnOpen";
                    string ListConfigurationID = "m_ListConfigurationIds";

                    // Select Shell ConfigID
                    AutomationElement aeLauoutIDs
                        = AUIUtilities.FindElementByID(ListConfigurationID, element);
                    if (aeLauoutIDs == null)
                    {
                        Console.WriteLine("FindElementByID failed:" + ListConfigurationID);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "FindElementByID failed:" + ListConfigurationID, sOnlyUITest);
                    }
                    //aeLauoutIDs.SetFocus();

                    // find all listitem
                    System.Windows.Automation.Condition cF = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                        new PropertyCondition(AutomationElement.IsControlElementProperty, true));

                    // Find all children that match the specified conditions.
                    AutomationElementCollection foundCollection = aeLauoutIDs.FindAll(TreeScope.Children, cF);
                    if (foundCollection == null)
                    {
                        Console.WriteLine("Shell Configuration is empty:");
                    }

                    int foundIdx = -1;
                    for (int i = 0; i < foundCollection.Count; i++)
                    {
                        Console.WriteLine("Shell Configuration: :" + foundCollection[i].Current.Name);
                        string configName = foundCollection[i].Current.Name;
                        if (configName.Equals("TestConfigSecurityUnsecured"))
                        {
                            foundIdx = i;
                            break;
                        }
                    }

                    if (foundIdx < 0)
                    {
                        Console.WriteLine("Shell SecurityUnsecured Configuration not found :");
                        return;
                    }

                    // Get Opened ConfigSecurityEpia
                    AutomationElement aeindex = foundCollection[foundIdx];
                    if (aeindex != null)
                    {
                        sConfigurationName = aeindex.Current.Name;
                        Console.WriteLine("open layout name :" + sConfigurationName);
                        SelectionItemPattern sip = (SelectionItemPattern)aeindex.GetCurrentPattern(SelectionItemPattern.Pattern);
                        sip.Select();
                    }

                    Rect rect;
                    rect = (Rect)(aeindex.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
                    System.Windows.Point pt = new System.Windows.Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
                    Input.MoveToAndClick(pt);

                    // Find Open Button and click
                    if (AUIUtilities.FindElementAndClick(BtnOpenID, element))
                        Thread.Sleep(3000);
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnOpenID);
                    }
                }
                else
                    Console.WriteLine("Name is ------------:" + name);
            }
            catch (Exception ex)
            {
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnConfigSecurityUnsecuredUIAEvent exception:" + ex.Message + "----" + ex.StackTrace, sOnlyUITest);
                Console.WriteLine("OnConfigSecurityUnsecuredUIAEvent exception:" + ex.Message + "----" + ex.StackTrace);
            }
        }
        #endregion
        #endregion Event +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region //------------------------Helper -------------------------------------
        /// <summary>
        /// network related struct
        /// </summary>
        /*private struct NETRESOURCEA
        {
            public int dwScope;
            public int dwType;
            public int dwDisplayType;
            public int dwUsage;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpLocalName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpRemoteName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpComment;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpProvider;
            public override String ToString()
            {
                String str = "LocalName: " + lpLocalName + " RemoteName: " + lpRemoteName
                    + " Comment: " + lpComment + " lpProvider: " + lpProvider;
                return (str);
            }
        }
        */
        /// <summary>
        /// Create a drive mapping to the destination
        /// </summary>
        /// <param name="Destination">Full drive path</param>
        /*static int CreateDriveMap(string Destination)
        {
            if ((Destination == null) || (Destination == ""))
                return -1;

            NETRESOURCEA[] netResource = new NETRESOURCEA[1];
            netResource[0] = new NETRESOURCEA();
            netResource[0].dwType = 1;
            netResource[0].lpLocalName = "";
            netResource[0].lpRemoteName = Destination;
            netResource[0].lpProvider = null;
            int dwFlags = 24; /*CONNECT_INTERACTIVE = 8|CONNECT_PROMPT = 16*/
        //int result = WNetAddConnection2A(netResource, null, null, dwFlags);
        //return result;
        //}*/

        public static AutomationElement FindPanelSelectionButton(AutomationElement root, string Category, string butoonName)
        {
            // Find ToolBar
            System.Windows.Automation.Condition c1 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "stackStrip1"),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar)
            );

            AutomationElement aeToolBar
                = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c1);

            if (aeToolBar == null)
            {
                Console.WriteLine(butoonName + " start find time1: " + System.DateTime.Now.ToString("HH:mm:ss"));
                //result = ConstCommon.TEST_FAIL;
                return null;
            }
            else
                Console.WriteLine(butoonName + " ToolBar found at time1: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Thread.Sleep(2000);

            // Find "Category" Button Element
            System.Windows.Automation.Condition c2 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, Category),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
            );

            // Find "Your Layouts" HyperLink Element
            System.Windows.Automation.Condition c3 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, butoonName),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Hyperlink)
             );

            AutomationElement aeYourLayouts
               = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c3);

            return aeYourLayouts;
        }

        #endregion



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
            ref string buildNr, ref string testApp, ref string buildDef, ref string buildConfig)
        {
            buildBaseDir = getBuildBasePath(installScriptsDir);
            buildNr = getBuildnr(installScriptsDir);

            if (buildNr.StartsWith("Etricc"))
                testApp = "Etricc";
            else if (buildNr.StartsWith("Epia"))
                testApp = "Epia";

            //int ib = buildNr.IndexOf("-");   buildnr = Epia.Development.Dev03.Nightly_20110418.1
            int ie = buildNr.IndexOf("_");

            string b = buildNr.Substring(0, ie);   // b = Epia.Development.Dev03.Nightly
            int ib = b.LastIndexOf(".");

            buildDef = b.Substring(ib + 1, ie - (ib + 1)).Trim();   // buildDef = Nightly
            //buildDef = buildNr.Substring(ib + 1, ie - (ib + 1)).Trim();

            if (installScriptsDir.IndexOf("Debug") > 0)
                buildConfig = "Debug";
            else
                buildConfig = "Release";
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
                if (y.StartsWith("E") && y.IndexOf("_") > 0 && y.IndexOf(".") > 0)
                    nr = y;
                // one level up
                path = path.Substring(0, ib);
                ib = path.LastIndexOf("\\");
            }

            return nr;
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
                if (y.StartsWith("E") && y.IndexOf("_") > 0)
                {
                    return path;
                }
                path = path.Substring(0, ib);
                ib = path.LastIndexOf("\\");
            }

            return string.Empty;
        }

        public static void CreateOutputFileInfo(string[] args, string PCName, ref string outPath, ref string outFilename)
        {

            // out filename 
            outFilename = System.DateTime.Now.ToString("yyyyMMdd-HH-mm-ss") + "-GUITESTS";

            if (args[0].IndexOf("Debug") > 0)
                outFilename = "Debug-" + outFilename + "-" + PCName;
            else
                outFilename = "Release-" + outFilename + "-" + PCName;

            if (args[2].ToLower().StartsWith("false"))
                outFilename = "Manual-" + outFilename;

            // out file path
            try
            {
                string path = args[0];
                Console.WriteLine("path : " + path);
                // save test result to TFS
                if (args[0].IndexOf("Etricc -") > 0)
                {
                    int ib = path.LastIndexOf("\\");
                    path = path.Substring(0, ib);
                    ib = path.LastIndexOf("\\");
                    outPath = path.Substring(0, ib);
                }
                else if (args[0].IndexOf("Epia -") > 0)
                {
                    if (args[0].IndexOf("Weekly") > 0) // only Weekly has folder: Mixed Platforms
                    {
                        int ib = path.LastIndexOf("\\");
                        path = path.Substring(0, ib);
                        ib = path.LastIndexOf("\\");
                        path = path.Substring(0, ib);
                        ib = path.LastIndexOf("\\");
                        outPath = path.Substring(0, ib);
                    }
                    else if (args[0].IndexOf("Version") >= 0) // 
                    {
                        int ib = path.LastIndexOf("\\");
                        path = path.Substring(0, ib);
                        Console.WriteLine("path : " + path);
                        ib = path.LastIndexOf("\\");
                        path = path.Substring(0, ib);
                        Console.WriteLine("path : " + path);
                        ib = path.LastIndexOf("\\");
                        path = path.Substring(0, ib);
                        Console.WriteLine("path : " + path);
                        ib = path.LastIndexOf("\\");
                        outPath = path.Substring(0, ib);
                        Console.WriteLine("outPath : " + outPath);
                    }
                    else  // CI and Nightly removed folder Mixed Platform
                    {
                        int ib = path.LastIndexOf("\\");
                        path = path.Substring(0, ib);
                        ib = path.LastIndexOf("\\");
                        outPath = path.Substring(0, ib);
                    }
                }
                else if (args[0].IndexOf("KC -") > 0)
                {   // X:\CI\Ewcs - Projects\Kimberly Clark\KC - CI_20091202.9
                    int ib = path.LastIndexOf("\\");
                    path = path.Substring(0, ib);
                    ib = path.LastIndexOf("\\");
                    outPath = path.Substring(0, ib);
                }

                outPath = outPath + "\\TestResults";
            }

            catch (Exception ex)
            {
                throw new Exception("TestTools: CreateOutputFileInfo: " + ex.Message + "----: " + ex.StackTrace);
                //MessageBox.Show("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                //WriteTestLogMsg("Save Xls file exception: " + ex.Message + "----: " + ex.StackTrace);
            }

        }

        /*[DllImport("mpr.dll")]
        private static extern int WNetAddConnection2A(
            [MarshalAs(UnmanagedType.LPArray)] NETRESOURCEA[] lpNetResource,
            [MarshalAs(UnmanagedType.LPStr)] string lpPassword,
            [MarshalAs(UnmanagedType.LPStr)] string UserName,
            int dwFlags);
        */
    }

}
