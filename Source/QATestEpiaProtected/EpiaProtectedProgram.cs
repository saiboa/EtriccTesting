using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Security.Principal;
using System.ServiceProcess;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using TFSQATestTools;
using TestTools;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace QATestEpiaProtected
{
    internal class EpiaProtectedProgram
    {
        #region Fields of Program (38)

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
        internal static Logger logger;

        // Test Param. =======================================
        private static readonly string[] sTestCaseName = new string[100];
        private static DateTime sTestStartUpTime = DateTime.Now;
        private static DateTime sStartTime = DateTime.Now;
        private static string sTestApp = string.Empty;
        private static AutomationElement aeForm;
        //private static TimeSpan sTime;
        private static int Counter;
        private static int sTotalCounter;
        private static int sTotalException;
        private static int sTotalFailed;
        private static int sTotalPassed;
        private static int sTotalUntested;
        private static int TestCheck = ConstCommon.TEST_UNDEFINED;
        private static bool sOnlyUITest;
        private static string sParentProgram = string.Empty;
        // PCinfo
        public static string PCName;
        public static string OSName;
        public static string OSVersion;
        public static string UICulture;
        public static string TimeOnPC;
        // Build info
        private static string sBuildDropFolder = string.Empty;
        private static string sBuildConfig = string.Empty;
        private static string sBuildDef = string.Empty;
        private static string sBuildNr = string.Empty;
        private static string sTestToolsVersion = string.Empty;
        private static string m_SystemDrive = string.Empty;
        private static string UserPassword = "Egemin01";
        private static string sTargetPlatform = string.Empty;
        private static string sCurrentPlatform = string.Empty;
        private static string sTestResultFolder = string.Empty;

        private static string sTestDefinitionFile = string.Empty;
        private static string[] mTestDefinitionTypes;
        private static string sInfoFileKey = string.Empty;
        private static string sNetworkMap = string.Empty;

        // Testcase not used =================================
        public static string sConfigurationName = string.Empty;
        private static string sErrorMessage;
        private static bool sEventEnd;
        private static string sExcelVisible = string.Empty;
        private static bool sAutoTest = true;
        private static string sInstallMsiDir = string.Empty;
        public static string sLayoutName = string.Empty;
        private static string sServerRunAs = string.Empty;
        private static bool sDemo;
        private static string sSendMail = "false";
        private static string sTFSServer = "http://Team2010App.TeamSystems.Egemin.Be:8080";


        // LOG=================================================================
        public static string slogFilePath = @"C:\";
        private static string sOutFilename = "OutFilename";
        private static string sOutFilePath = string.Empty;
        private static StreamWriter Writer;
        // Build param ========================================================
        //static BuildStore   buildStore      = null;

        private static IBuildServer m_BuildSvc;
        private static bool TFSConnected = true;
        // excel 	--------------------------------------------------------
        private static Application xApp;
        private static int sHeaderContentsLength = 10;
        //static Excel.Range xRange;
        // default layout

        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
        #region Methods of Program (1)
        /// <summary>
        /// 
        /// </summary>
        /// <param name="args">Application inputs
        ///     1   InstallScripts Directory
        ///     2   Visible or Invisible (Excel )
        ///     3   true or false (auto test) 
        ///     </param>
        [STAThread]
        private static void Main(string[] args)
        {
            try // Get test PC info======================================
            {
                HelpUtilities.SavePCInfo("y");
                HelpUtilities.GetPCInfo(out PCName, out OSName, out OSVersion, out UICulture, out TimeOnPC);
                Console.WriteLine("<PCName : " + PCName + ">, <OSName : " + OSName + ">, <OSVersion : " + OSVersion +
                                  ">");
                Console.WriteLine("<TimeOnPC : " + TimeOnPC + ">, <UICulture : " + UICulture + ">");
                string windir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
                m_SystemDrive = Path.GetPathRoot(windir);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Check PC-info");
            }


            sOnlyUITest = false;
            string x = ConfigurationManager.AppSettings.Get("OnlyUITest");
            if (x.ToLower().StartsWith("true"))
                sOnlyUITest = true;

            Console.WriteLine("sOnlyUITest : " + sOnlyUITest);
            UserPassword = ConfigurationManager.AppSettings.Get("CurrentUserPassword");

            if (!sOnlyUITest)
            {
                try
                {
                    // validate inputs
                    if (args != null)
                    {
                        for (int i = 0; i <= 18; i++)
                        {
                            Console.WriteLine(i + " de args : " + args[i]);
                        }
                        sInstallMsiDir = args[0];
                        sBuildDropFolder = args[1];
                        sBuildNr = args[2];
                        string sProject = args[3];
                        sTestApp = args[4];
                        sTargetPlatform = args[5];
                        sCurrentPlatform = args[6];
                        sBuildDef = args[7];
                        sParentProgram = args[8];
                        sTestToolsVersion = args[9];
                        if (args[10].StartsWith("true"))
                            sAutoTest = true;
                        else
                            sAutoTest = false;

                        sTFSServer = args[11];
                        sServerRunAs = args[12];
                        sExcelVisible = args[13];
                        if (args[14].StartsWith("true"))
                            sDemo = true;
                        else
                            sDemo = false;

                        if (args[15].StartsWith("true"))
                            sSendMail = "true";
                        else
                            sSendMail = "false";

                        sTestDefinitionFile = args[16];
                        sInfoFileKey = args[17];
                        sNetworkMap = args[18];

                        sTestResultFolder = sBuildDropFolder + "\\TestResults";
                        if (!Directory.Exists(sTestResultFolder))
                            Directory.CreateDirectory(sTestResultFolder);

                        //Epia3Common.CreateOutputFileInfo(args, PCName, ref sOutFilePath, ref sOutFilename);
                        //CreateOutputFileInfo(args, sCurrentPlatform, PCName, ref sOutFilePath, ref sOutFilename);
                        sOutFilename = FileManipulation.CreateOutputInfoFileName(sInfoFileKey, sAutoTest);
                        sOutFilePath = Path.GetFullPath(sBuildDropFolder) + "\\TestResults";
                        Console.WriteLine("sOutFilePath : " + sOutFilePath);
                        Console.WriteLine("sOutFilePath2 : " + Path.GetDirectoryName(sOutFilePath));
                        Console.WriteLine("sOutFilename : " + sOutFilename);


                        Epia3Common.CreateTestLog(ref slogFilePath, sOutFilePath, sOutFilename, ref Writer);
                        logger = new Logger(slogFilePath); // logger passed to other object

                        Epia3Common.WriteTestLogMsg(slogFilePath, "0) Install msi file path: " + sInstallMsiDir,
                                                    sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "1) BuildBaseDir: " + sBuildDropFolder, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "2) build nr: " + sBuildNr, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "3) test Project: " + sProject, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "4) test Application: " + sTestApp, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "5) targeted platform: " + sTargetPlatform,
                                                    sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "6) current platform: " + sCurrentPlatform,
                                                    sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "7) test def: " + sBuildDef, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "8) Called by: " + sParentProgram, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "9) TestTool version: " + sTestToolsVersion,
                                                    sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "10) Auto test: " + sAutoTest, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "11) TFSServerUrl: " + sTFSServer, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "12) Server Run As: " + sServerRunAs, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "13) Excel Visible: " + sExcelVisible, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "14) Demo test: " + sDemo, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "15) Mail: " + sSendMail, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "16) TestDefinitionFile: " + sTestDefinitionFile,
                                                    sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "17) InfoFileKey: " + sInfoFileKey, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "18) NetworkMap: " + sNetworkMap, sOnlyUITest);


                        Epia3Common.WriteTestLogMsg(slogFilePath, "m_SystemDrive: " + m_SystemDrive, sOnlyUITest);

                        // sInstallScriptsDir = X:\CI\Epia 3\Epia - CI_20100324.1\Debug\Installation\Setup\Debug
                        // sBuildDropFolder = X:\CI\Epia 3\Epia - CI_20100324.1
                        // sBuildNr = Epia - CI_20100324.1
                        // sTestApp = Epia
                        // sBuildDef = CI
                        // sBuildConfig = Debug
                        Console.WriteLine("0) sInstall msi files Dir : " + sInstallMsiDir);
                        Console.WriteLine("1) sBuildDropFolder : " + sBuildDropFolder);
                        Console.WriteLine("2) sBuildNr : " + sBuildNr);
                        Console.WriteLine("3) sProject : " + sProject);
                        Console.WriteLine("4) sTestApp : " + sTestApp);
                        Console.WriteLine("5) sTargetPlatform : " + sTargetPlatform);
                        Console.WriteLine("6) sCurrentPlatform : " + sCurrentPlatform);
                        Console.WriteLine("7) sBuildDef : " + sBuildDef);
                        Console.WriteLine("8) Called by: " + sParentProgram);
                        Console.WriteLine("9) TestTool version: " + sTestToolsVersion);
                        Console.WriteLine("10) Auto test: " + sAutoTest);
                        Console.WriteLine("11) TFSServerUrl: " + sTFSServer);
                        Console.WriteLine("12) Server Run As: " + sServerRunAs);
                        Console.WriteLine("13) Excel Visible: " + sExcelVisible);
                        Console.WriteLine("14) Demo test: " + sDemo);
                        Console.WriteLine("15) Mail: " + sSendMail);
                        Console.WriteLine("16) TestDefinitionFile: " + sTestDefinitionFile);
                        Console.WriteLine("17) InfoFileKey: " + sInfoFileKey);
                        Console.WriteLine("18) NetworkMap: " + sNetworkMap);

                        mTestDefinitionTypes = File.ReadAllLines(sTestDefinitionFile);

                        for (int i = 0; i < mTestDefinitionTypes.Length; i++)
                        {
                            Console.WriteLine(i + " testdefinition : " + mTestDefinitionTypes[i]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Validate command-line params");
                }
            }

            if (!sOnlyUITest)
            {
                if (sAutoTest)
                {
                    try
                    {
                        // Get TFS Server
                        string serverUrl = "http://team2010app.teamsystems.egemin.be:8080/tfs/Development";
                        var serverUri = new Uri(serverUrl);
                        ICredentials tfsCredentials
                            = new NetworkCredential("TfsBuild", "Egemin01", "TeamSystems.Egemin.Be");

                        var tfsProjectCollection
                            = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                        int kTime = 0;
                        bool conn = false;
                        while (conn == false)
                        {
                            try
                            {
                                m_BuildSvc = (IBuildServer) tfsProjectCollection.GetService(typeof (IBuildServer));
                                conn = true;
                            }
                            catch (TeamFoundationServiceUnavailableException ex)
                            {
                                MessageBoxEx.Show(
                                    "Team Foundation services are not available from server\nWill try to reconnect the Server ...\nException message:"+ex.Message,
                                    kTime++ + " During protected Epia UI Testing, please not touch the screen, time: " +
                                    DateTime.Now.ToLongTimeString(), (uint) Tfs.ReconnectDelay);
                                Thread.Sleep(Tfs.ReconnectDelay);
                                conn = false;
                            }
                            catch (Exception ex)
                            {
                                MessageBoxEx.Show(
                                    "TeamFoundation getService Exception:" + ex.Message + " ----- " + ex.StackTrace,
                                    kTime++ + " This is automatic testing, please not touch the screen: exception time:" +
                                    DateTime.Now.ToLongTimeString(), (uint) Tfs.ReconnectDelay);
                                Thread.Sleep(Tfs.ReconnectDelay);
                                conn = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Get TFS Server:" + ex.Message, sOnlyUITest);
                        TFSConnected = false;
                    }
                }
                else
                    TFSConnected = false;
            }

            Console.WriteLine("Test started:");
            Epia3Common.WriteTestLogMsg(slogFilePath, "Test started: ", sOnlyUITest);
            sTestCaseName[0] = EPIA_SERVICE_START;
            sTestCaseName[1] = EPIA_SERVICE_UNINSTALL;
            sTestCaseName[2] = LAYOUT_INITIAL_Y_POSITION;
            sTestCaseName[3] = LAYOUT_INITIAL_WIDTH;
            sTestCaseName[4] = LAYOUT_INITIAL_HEIGHT;
            sTestCaseName[5] = LAYOUT_ALLOW_RESIZE;
            sTestCaseName[6] = LAYOUT_FULL_SCREEN;
            sTestCaseName[7] = LAYOUT_MAXIMIZED;
            sTestCaseName[8] = LAYOUT_RIBBON_ON;
            //sTestCaseName[9] = LAYOUT_TITLE;
            sTestCaseName[9] = LAYOUT_CANCEL_BUTTON;
            sTestCaseName[22] = LAYOUT_NAVIGATOR_OFF;
            sTestCaseName[10] = SETTING_LANGUAGE;
            sTestCaseName[11] = SHELL_CONFIGURATION_SECURITY;
            sTestCaseName[12] = LOGON_CURRENT_USER;
            //sTestCaseName[14] = LOGON_EPIA_ADMINISTRATOR;
            sTestCaseName[13] = SHELL_SHUTDOWN;
            sTestCaseName[14] = SHELL_LOGOFF;
            sTestCaseName[15] = EPIA4_CLOSEE;
            //=============================================//
            sTestCaseName[34] = LAYOUT_OPEN;
            sTestCaseName[35] = CONFIGURATION_OPEN;
            sTestCaseName[40] = CONFIGURATION_SAVE;
            sTestCaseName[36] = CONFIGURATION_SECURITY_UNSECURED;

            //sTestCaseName[0] = CONFIGURATION_FIND_GRIDVIEW;
            //sTestCaseName[19] = LAYOUT_MAINMENU_ON;

            try
            {
                if (!sOnlyUITest)
                {
                    ProcessUtilities.CloseProcess("EXCEL");
                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    Thread.Sleep(1000);

                    AutomationEventHandler UIAShellEventHandler = OnUIAShellEvent;
                    // Add Open window Event Handler
                    Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                         AutomationElement.RootElement, TreeScope.Descendants,
                                                         UIAShellEventHandler);
                    sEventEnd = false;
                    TestCheck = ConstCommon.TEST_PASS;

                    Console.WriteLine("TestCkeck : " + TestCheck.ToString());
                    Console.WriteLine("Application is started : ");
                    /*aeForm = null;
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
					*/
                }

                // Excel file not for EpiaTestPC3 and EPIATESTSERVER3
                xApp = new Application();
                string[] sHeaderContents =
                    {
                        DateTime.Now.ToString("MMMM-dd") + "*" + "Epia" + " UI Test Scenarios",
                        "Test Machine:" + "*" + PCName,
                        "Tester::" + "*" + WindowsIdentity.GetCurrent().Name,
                        "OSName:" + "*" + OSName,
                        "OS Version:" + "*" + OSVersion,
                        "UI Culture:" + "*" + UICulture,
                        "Time On PC:" + "*" + "local time:" + TimeOnPC,
                        "Test Tool Version:" + "*" + sTestToolsVersion,
                        "NetworkMap:" + "*" + sNetworkMap,
                        "Build Location:" + "*" + sInstallMsiDir,
                    };

                sHeaderContentsLength = sHeaderContents.Length;
                Epia3Common.WriteTestLogMsg(slogFilePath, "sHeaderContents.length: " + sHeaderContentsLength,
                                            sOnlyUITest);
                FileManipulation.WriteExcelHeader(ref xApp, sExcelVisible, sHeaderContents);

                //FileManipulation.WriteExcelHeader(ref xApp, "Epia", sExcelVisible, PCName, OSName, OSVersion,
                //   UICulture, TimeOnPC, sTestToolsVersion, sNetworkMap, sInstallMsiDir);

                // start test----------------------------------------------------------
                int sResult = ConstCommon.TEST_UNDEFINED;
                int aantal = 1;
                if (sDemo)
                    aantal = 1;

                if (sOnlyUITest)
                {
                    string testType = ConfigurationManager.AppSettings.Get("TestType");
                    Console.WriteLine("--------------------------------------testType : " + testType);

                    if (testType.ToLower().StartsWith("all"))
                        aantal = 16;
                    else
                    {
                        int thisTest = Convert.ToInt16(testType);
                        aantal = 1;
                        sTestCaseName[0] = sTestCaseName[thisTest - 1];
                    }
                }
                else
                {
                    /*if (TFSConnected)
					{
						Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
							TestTools.TfsUtilities.GetProjectName("Epia"), sBuildNr);
						string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

						if (quality.Equals("GUI Tests Failed"))
						{
							Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has build quality: " + quality + " , no update needed", sOnlyUITest);
						}
						else
						{
							TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
								TestTools.TfsUtilities.GetProjectName("Epia"),
								"GUI Tests Started", m_BuildSvc, sDemo ? "true" : "false");
						}

						if (sAutoTest)
						{
							if (sInstallScriptsDir.IndexOf("Protected") > 0)
								//TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", "Epia4+" + sCurrentPlatform + "Protected");
							else
								//TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", "Epia4+" + sCurrentPlatform + "Normal");
							Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.EPIA, sOnlyUITest);
						}
					}*/
                }

                while (Counter < aantal)
                {
                    sResult = ConstCommon.TEST_UNDEFINED;
                    switch (sTestCaseName[Counter])
                    {
                        case EPIA_SERVICE_START:
                            EpiaServiceStart(EPIA_SERVICE_START, aeForm, out sResult);
                            break;
                        case EPIA_SERVICE_UNINSTALL:
                            EpiaServiceUninstall(EPIA_SERVICE_UNINSTALL, aeForm, out sResult);
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
                            ShellConfigSecurity(SHELL_CONFIGURATION_SECURITY, aeForm, out sResult);
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
                        case EPIA4_CLOSEE:
                            Epia4Close(EPIA4_CLOSEE, aeForm, out sResult);
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

                    // Write this testcase result to Excel
                    FileManipulation.WriteExcelTestCaseResult(xApp, sResult, sHeaderContentsLength, Counter,
                                                              sTestCaseName[Counter], sErrorMessage);

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
                }

                Epia3Common.WriteTestLogMsg(slogFilePath, "Total Tests: " + Counter, sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "Total Passed: " + sTotalPassed, sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "Total Failed: " + sTotalFailed, sOnlyUITest);

                // Write Excel Foot
                FileManipulation.WriteExcelFoot(xApp, sHeaderContentsLength, Counter, sTotalCounter, sTotalPassed,
                                                sTotalFailed);

                #region // save Excel to Local machine and remote machine

                string sXLSPath = Path.Combine(Directory.GetCurrentDirectory(), sOutFilename + ".xls");
                Epia3Common.WriteTestLogMsg(slogFilePath, "Save to local machine : " + sXLSPath, sOnlyUITest);
                if (FileManipulation.SaveExcel(xApp, sXLSPath, ref sErrorMessage) == false)
                {
                    string sTXTPath = Path.Combine(Directory.GetCurrentDirectory(), sOutFilename + ".txt");
                    StreamWriter write = File.CreateText(sTXTPath);
                    write.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                    write.Close();
                }

                // Save to remote machine
                if (WindowsIdentity.GetCurrent().Name.ToUpper().StartsWith("TEAMSYSTEMS\\JIEMINSHI"))
                {
                    Console.WriteLine("\n   not write to remote machine");
                }
                else
                {
                    string sXLSPath2 = Path.Combine(sOutFilePath, sOutFilename + ".xls");
                    Console.WriteLine("Save2 : " + sXLSPath2);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "sXLSPath2 =: " + sXLSPath2, sOnlyUITest);
                    if (FileManipulation.SaveExcel(xApp, sXLSPath2, ref sErrorMessage) == false)
                    {
                        string sTXTPath = Path.Combine(Directory.GetCurrentDirectory(), sOutFilename + ".txt");
                        StreamWriter write = File.CreateText(sTXTPath);
                        write.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        write.Close();
                    }
                }
                // quit Excel.
                xApp.Quit();

                // Send Result via Email
                if (!sOnlyUITest)
                    SendEmail(sXLSPath);

                #endregion

                if (!sOnlyUITest)
                {
                    string msgX = "update epia build quality test status to Passed if needed";
                    TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                    while (TFSConnected == false)
                    {
                        MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                          "update build quality Deployment Failed2", (uint) Tfs.ReconnectDelay);
                        Thread.Sleep(Tfs.ReconnectDelay);
                        Console.WriteLine(" Reconnect TFS Server:::::: ");
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                    }

                    if (TFSConnected)
                    {
                        // added check sTestResultFolder exist; some time during testing this build can be completely deleted by WVB
                        if (Directory.Exists(sTestResultFolder))
                        {
                            #region  // update testinfo file first and then update build quality

                            if (sAutoTest)
                            {
                                Epia3Common.WriteTestLogMsg(slogFilePath,
                                                            "TfsUtilities.GetBuildUriFromBuildNumber: " + sTotalFailed,
                                                            sOnlyUITest);
                                string prjName = TfsUtilities.GetProjectName(TestApp.EPIA4);
                                Epia3Common.WriteTestLogMsg(slogFilePath, "prjName: " + prjName, sOnlyUITest);
                                Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                                                                                  TfsUtilities.GetProjectName(
                                                                                      TestApp.EPIA4), sBuildNr);
                                string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;
                                Epia3Common.WriteTestLogMsg(slogFilePath,
                                                            "m_BuildSvc.GetMinimalBuildDetails(uri).Quality: " + quality,
                                                            sOnlyUITest);
                                if (sTotalFailed == 0)
                                {
                                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Passed",
                                                                                 "Tests OK", sInfoFileKey);
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Passed:" + TestApp.EPIA4,
                                                                sOnlyUITest);

                                    Console.WriteLine(" Update build quality:  quality: " + quality);
                                    if (quality.Equals("GUI Tests Failed"))
                                    {
                                        Console.WriteLine("Quality is GUI Tests Failed : No Update Needed ");
                                        Epia3Common.WriteTestLogMsg(slogFilePath,
                                                                    sBuildNr + " has failed quality, no update needed :" +
                                                                    quality, sOnlyUITest);
                                    }
                                    else if (
                                        TestListUtilities.IsAllTestDefinitionsTested(mTestDefinitionTypes,
                                                                                     sTestResultFolder,
                                                                                     ref sErrorMessage) == false)
                                        //else if ( TestListUtilities.IsAllTestDefinitionsTested( mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage ) == false )
                                    {
                                        Console.WriteLine("NOT All Test definitions tested " + sErrorMessage);
                                        Epia3Common.WriteTestLogMsg(slogFilePath,
                                                                    "NOT All Test definitions tested " +
                                                                    sErrorMessage, sOnlyUITest);
                                    }
                                    else
                                    {
                                        //if (TestListUtilities.IsAllTestStatusPassed(mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage) == true)
                                        if (TestListUtilities.IsAllTestStatusPassed(slogFilePath,
                                                                                    mTestDefinitionTypes,
                                                                                    sTestResultFolder,
                                                                                    ref sErrorMessage))
                                        {
                                            // update quality to GUI Tests Passed
                                            TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                                                                  TfsUtilities.GetProjectName(
                                                                                      TestApp.EPIA4),
                                                                                  "GUI Tests Passed", m_BuildSvc,
                                                                                  sDemo ? "true" : "false");

                                            Console.WriteLine("update quality to true -----  ");
                                            Epia3Common.WriteTestLogMsg(slogFilePath,
                                                                        "update quality to true ----- ", sOnlyUITest);
                                            Thread.Sleep(1000);
                                        }
                                        else
                                        {
                                            Epia3Common.WriteTestLogMsg(slogFilePath,
                                                                        "IsAllTestStatusPassed failed ----- " +
                                                                        sErrorMessage, sOnlyUITest);
                                        }
                                    }
                                }
                                else
                                {
                                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Failed",
                                                                                 "--->" + sOutFilename + ".log",
                                                                                 sInfoFileKey);
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Failed:" + TestApp.EPIA4,
                                                                sOnlyUITest);

                                    Console.WriteLine(" Update build quality:  quality: " + quality);
                                    if (quality.Equals("GUI Tests Failed"))
                                    {
                                        Console.WriteLine("Quality is GUI Tests Failed : No Update Needed ");
                                        Epia3Common.WriteTestLogMsg(slogFilePath,
                                                                    sBuildNr + " has failed quality, no update needed :" +
                                                                    quality, sOnlyUITest);
                                    }
                                    else
                                    {
                                        // update quality to GUI Tests Passed
                                        TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                                                              TfsUtilities.GetProjectName(TestApp.EPIA4),
                                                                              "GUI Tests Failed", m_BuildSvc,
                                                                              sDemo ? "true" : "false");

                                        Console.WriteLine("update quality to GUI Tests Failed -----  ");
                                        Thread.Sleep(1000);
                                    }
                                }
                            }

                            #endregion
                        }
                    }
                }

                // Close LogFile
                Epia3Common.CloseTestLog(slogFilePath, sOnlyUITest);

                Console.WriteLine("\nClosing application in 10 seconds");
                if (sOnlyUITest)
                    Thread.Sleep(10000);
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
                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                Thread.Sleep(10000);
                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                ProcessUtilities.CloseProcess("cmd");
                Console.WriteLine("\nEnd test run\n");
                //Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Thread.Sleep(2000);
                if (sAutoTest)
                {
                    #region // test exception : update infofile and build quality

                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Exception",
                                                                 " -->" + sOutFilename + ".log", sInfoFileKey);
                    Epia3Common.WriteTestLogMsg(slogFilePath,
                                                "GUI Tests Exception -->" + sOutFilename + ".log:" + ConstCommon.EPIA,
                                                sOnlyUITest);

                    Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace, sOnlyUITest);

                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                    ProcessUtilities.CloseProcess("cmd");

                    string msgX = "epia exception build quality test status to Failed if needed";
                    TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                    while (TFSConnected == false)
                    {
                        MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                          "update build quality Deployment Failed2", (uint) Tfs.ReconnectDelay);
                        Thread.Sleep(Tfs.ReconnectDelay);
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                    }

                    if (TFSConnected)
                    {
                        Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                                                                          TfsUtilities.GetProjectName("Epia"), sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                        if (quality.Equals("GUI Tests Failed"))
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath,
                                                        sBuildNr + " has failed quality, no update needed :" + quality,
                                                        sOnlyUITest);
                        }
                        else
                        {
                            TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName("Epia"),
                                                                  "GUI Tests Failed", m_BuildSvc,
                                                                  sDemo ? "true" : "false");
                        }
                    }

                    #endregion
                }
            }
        }

        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        #region TestCase Name

        private const string EPIA_SERVICE_UNINSTALL = "EpiaServiceUninstall";
        private const string EPIA_SERVICE_START = "EpiaServiceStart";
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
        private const string SHELL_CONFIGURATION_SECURITY = "ShellConfigSecurity";
        private const string LOGON_CURRENT_USER = "LogonCurrentUser";
        private const string LOGON_EPIA_ADMINISTRATOR = "LogonEpiaAdmin";
        private const string SHELL_SHUTDOWN = "ShellShutdown";
        private const string SHELL_LOGOFF = "ShellLogOff";
        private const string EPIA4_CLOSEE = "Epia4Close";
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        private const string LAYOUT_NAVIGATOR_OFF = "LayoutNavigatorOff";

        private const string LAYOUT_OPEN = "LayoutOpen";
        private const string CONFIGURATION_OPEN = "ConfigOpen";
        private const string CONFIGURATION_SAVE = "ConfigSave";
        private const string CONFIGURATION_SECURITY_UNSECURED = "ConfigSecurityUnsecured";
        //private const string CONFIGURATION_SECURITY_EPIA = "ConfigSecurityEpia";
        private const string CONFIGURATION_FIND_GRIDVIEW = "ConfigFindGridView";

        #endregion TestCase Name

        #region Test Cases -------------------------------------------------------------------------------------------

        #region EpiaServiceUninstall

        public static void EpiaServiceUninstall(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIAFindLayoutPanelEventHandler = OnWIBUSYSTEMSWindowOpenEvent;

            try
            {
                // Add Open MyLayoutScreen window Event Handler
                //AutomationEventHandler UIAFindLayoutPanelEventHandler = new AutomationEventHandler(OnWIBUSYSTEMSWindowOpenEvent);
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIAFindLayoutPanelEventHandler);

                // uninstall Egemin.Epia.server Service
                Console.WriteLine("UNINSTALL EPIA SERVER Service : ");
                ProcessUtilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
                                                         ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /u");
                Thread.Sleep(2000);

                sEventEnd = false;
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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

        #region EpiaServiceStart

        public static void EpiaServiceStart(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIAFindLayoutPanelEventHandler = OnWIBUSYSTEMSWindowOpenEvent;

            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 3, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;

                }

                // Add Open MyLayoutScreen window Event Handler
                //AutomationEventHandler UIAFindLayoutPanelEventHandler = new AutomationEventHandler(OnWIBUSYSTEMSWindowOpenEvent);
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIAFindLayoutPanelEventHandler);

                // uninstall Egemin.Epia.server Service
                Console.WriteLine("START EPIA SERVER Service : ");
                ProcessUtilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
                                                         ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /start");
                Thread.Sleep(2000);

                sEventEnd = false;
                sErrorMessage = string.Empty;
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

                while (sEventEnd == false && mTime.TotalSeconds <= 60)
                {
                    Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime.TotalSeconds);
                    mTime = DateTime.Now - mStartTime;
                }

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                        AutomationElement.RootElement,
                                                        UIAFindLayoutPanelEventHandler);


                var svcEpia = new ServiceController("Egemin Epia Server");
                Console.WriteLine(svcEpia.ServiceName + " has status " + svcEpia.Status.ToString());

                string epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                Thread.Sleep(5000);
                if (svcEpia.Status != ServiceControllerStatus.Running)
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Epia Service start up failed: " + epiaServiceStatus,
                                                sOnlyUITest);
                    //throw new Exception("Epia service start up failed:"); //   get message from log file sErrorMessage//
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Epia Service  is running";
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
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
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                        AutomationElement.RootElement,
                                                        UIAFindLayoutPanelEventHandler);
            }
        }

        #endregion

        #region LayoutInitialYPosition

        public static void LayoutInitialYPosition(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutYPosEventHandler = OnLayoutYPosUIAEvent;

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                                                                   ConstCommon.MY_LAYOUT,
                                                                                   ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutYPosEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                             sOnlyUITest);
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

        #region LayoutInitialWidth

        public static void LayoutInitialWidth(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutWidthEventHandler = OnLayoutWidthUIAEvent;

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                                                                   ConstCommon.MY_LAYOUT,
                                                                                   ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutWidthEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                             sOnlyUITest);
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

        #region LayoutInitialHeight

        public static void LayoutInitialHeight(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutHeightEventHandler = OnLayoutHeightUIAEvent;

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                                                                   ConstCommon.MY_LAYOUT,
                                                                                   ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutHeightEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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

                    if (Math.Abs(height - 500) < 10)
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
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                             sOnlyUITest);
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

        #region LayoutTitle

        public static void LayoutTitle(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutTitleEventHandler = OnLayoutTitleUIAEvent;

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                                                                   ConstCommon.MY_LAYOUT,
                                                                                   ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutTitleEventHandler);

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                             sOnlyUITest);
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

        #region LayoutAllowResize

        public static void LayoutAllowResize(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutResizeEventHandler = OnLayoutResizeUIAEvent;

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                                                                   ConstCommon.MY_LAYOUT,
                                                                                   ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutResizeEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);

                while (sEventEnd == false && mTime.Seconds <= 600)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                }
                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
                    var tranform =
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
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                             sOnlyUITest);
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

        #region LayoutFullScreen

        public static void LayoutFullScreen(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutFullScreenEventHandler = OnLayoutFullScreenUIAEvent;

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                                                                   ConstCommon.MY_LAYOUT,
                                                                                   ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutFullScreenEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                             sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                        AutomationElement.RootElement,
                                                        UIALayoutFullScreenEventHandler);
            }
        }

        #endregion

        #region LayoutMaximized

        public static void LayoutMaximized(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutMaximizedScreenEventHandler = OnLayoutMaximizedScreenUIAEvent;

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                                                                   ConstCommon.MY_LAYOUT,
                                                                                   ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MaximizedScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutMaximizedScreenEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                             sOnlyUITest);
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

        #region LayoutRibbonOn

        public static void LayoutRibbonOn(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;

            AutomationEventHandler UIALayoutRibbonOnEventHandler = OnLayoutRibbonOnUIAEvent;

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                                                                   ConstCommon.MY_LAYOUT,
                                                                                   ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutRibbonOnEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
                            Epia3Common.WriteTestLogMsg(slogFilePath,
                                                        "ribbon height is: " + aeRibbon.Current.BoundingRectangle.Height,
                                                        sOnlyUITest);
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
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                             sOnlyUITest);
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

        #region LayoutCancelButton

        public static void LayoutCancelButton(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;

            AutomationEventHandler UIALayoutCancelButtonEventHandler = OnLayoutCancelButtonUIAEvent;

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
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutCancelButtonEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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

        #region LayoutNavigatorOff

        public static void LayoutNavigatorOff(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;

            AutomationEventHandler UIALayoutNavigatorOffEventHandler = OnLayoutNavigatorOffUIAEvent;
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
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutNavigatorOffEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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

                Input.SendKeyboardInput(Key.Space, true);
            }
        }

        #endregion

        #region LanguageSetting

        public static void LanguageSetting(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALanguageSettingEventHandler = OnLanguageSettingUIAEvent;

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
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALanguageSettingEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeMySettings);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my settings :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
                    aeNLnode = AUIUtilities.FindTreeViewNodeByName(testname, aeTreeView, "Mijn instellingen",
                                                                   ref sErrorMessage);
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
        }

        #endregion

        #region ShellConfigSecurity

        public static void ShellConfigSecurity(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIAConfigSecurityEventHandler = OnConfigSecurityUIAEvent;

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
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom)/2;
                var shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                //while (root.Current.IsEnabled)
                //{
                Console.WriteLine("re click Shell Config Security :");
                Input.MoveToAndClick(shellPoint);
                Thread.Sleep(5000);
                //}

                // Add Open MyLanguageSetting window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIAConfigSecurityEventHandler);
                sEventEnd = false;

                double y2 = y + 15;
                var securityPoint = new Point(x, y2);
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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
        }

        #endregion

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
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom)/2;
                var shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);

                Console.WriteLine("aaaaaaaaaaaaaaaaaaaaaaaa ");
                Thread.Sleep(3000);

                double y2 = y + 40;
                var securityPoint = new Point(x, y2);
                Input.MoveTo(securityPoint);

                Console.WriteLine("bbbbbbbbbbbbbbbbbbb ");
                Thread.Sleep(3000);

                Input.MoveToAndClick(securityPoint);
                Thread.Sleep(3000);

                Console.WriteLine("After log off Shell, wait 2 second : ");
                Thread.Sleep(2000);
                Console.WriteLine("=== Test " + testname + " ===");
                Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
                ;
                result = ConstCommon.TEST_UNDEFINED;
                AutomationElement aeSecurityForm = null;


                /*Console.WriteLine("Starting : ");
				Thread.Sleep(3000);
			
					AutomationEventHandler UIACurrentUserEventHandler = new AutomationEventHandler(OnUIACurrentUserEvent);
					// Add Open window Event Handler
					Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						AutomationElement.RootElement, TreeScope.Descendants, UIACurrentUserEventHandler);
					sEventEnd = false;
					TestCheck = ConstCommon.TEST_PASS;
			  
					//string path = Path.Combine(sInstallScriptsDir, Constants.SHELL_BAT);
					string path = Path.Combine(m_SystemDrive+ConstCommon.EPIA_CLIENT_ROOT, 
						ConstCommon.EGEMIN_EPIA_SHELL_EXE);
					System.Diagnostics.Process proc = System.Diagnostics.Process.Start(path);
					Console.WriteLine("*****" + proc.Id);
					Thread.Sleep(9000);

					Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						   AutomationElement.RootElement,
						  UIACurrentUserEventHandler);

					Console.WriteLine("Application is started : ");*/

                Console.WriteLine("After Logoff, wait until LogonForm displaying... : ");
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
                    Console.WriteLine("Find Application aeSecurityForm : " + DateTime.Now);
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
                string tester = WindowsIdentity.GetCurrent().Name;


                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser, UserPassword,
                                                           ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + PasswordID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester,
                                                           ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + UserNameID;
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
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                // Set a property condition that will be used to find the control.
                Condition c = new PropertyCondition(
                    AutomationElement.AutomationIdProperty, "MainForm", PropertyConditionFlags.IgnoreCase);

                Console.WriteLine(" find total mainForm :");

                // Find the element.
                AutomationElementCollection aes =
                    AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                Thread.Sleep(10000);

                DateTime mAppTime = DateTime.Now;
                TimeSpan Time = DateTime.Now - mAppTime;
                while (aes.Count != 1 && Time.Minutes < 2)
                {
                    Console.WriteLine("Find Application aeForm : " + DateTime.Now);
                    aes = AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                    mTime = DateTime.Now - mAppTime;
                    Console.WriteLine(" find time is :" + Time.TotalMilliseconds/1000);
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
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                            sOnlyUITest);
            }
        }

        #endregion LogonCurrentUser

        #region LogonEpiaAdministrator

        public static void LogonEpiaAdministrator(string testname, AutomationElement root, out int result)
        {
            ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIAConfigSecurityEventHandler = OnLogonEpiaAdminUIAEvent;

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
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom)/2;
                var shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Thread.Sleep(3000);

                // Add Open MyLanguageSetting window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIAConfigSecurityEventHandler);
                sEventEnd = false;

                double y2 = y + 15;
                var securityPoint = new Point(x, y2);
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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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

                var logoffPoint = new Point(shellPoint.X + 5, shellPoint.Y + 45);
                Input.MoveTo(logoffPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(logoffPoint);
                Thread.Sleep(3000);

                // logon with hidden user
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
                    Console.WriteLine("Find Application aeSecurityForm : " + DateTime.Now);
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
                    sErrorMessage = "Failed find " + "Logon BUTTon " + " at time: " + DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(sErrorMessage);
                    return;
                }
                else
                {
                    Console.WriteLine("Logon button " + ": found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                    Point LogonPoint = AUIUtilities.GetElementCenterPoint(aeLogonButton);
                    Thread.Sleep(1000);
                    Input.MoveToAndClick(LogonPoint);
                    Thread.Sleep(2000);
                    var withGenericPoint = new Point(LogonPoint.X, LogonPoint.Y + 35);
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
                string origPassword = string.Empty;
                string tester = TFSQATestTools.Constants.HIDDEN_USERNAME;

                if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester,
                                                           ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + UserNameID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origPassword, TFSQATestTools.Constants.HIDDEN_PASSWORD,
                                                           ref sErrorMessage))
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
                Condition c = new PropertyCondition(
                    AutomationElement.AutomationIdProperty, "MainForm", PropertyConditionFlags.IgnoreCase);

                Console.WriteLine(" find total mainForm :");

                // Find the element.
                AutomationElementCollection aes =
                    AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                Thread.Sleep(10000);

                DateTime mAppTime = DateTime.Now;
                TimeSpan Time = DateTime.Now - mAppTime;
                while (aes.Count != 1 && Time.Minutes < 2)
                {
                    Console.WriteLine("Find Application aeForm : " + DateTime.Now);
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

        #region ShellShutdown

        public static void ShellShutdown(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;

            try
            {
                Process ShellProcess = null;
                int pID = ProcessUtilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out ShellProcess);
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
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom)/2;
                var shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Thread.Sleep(3000);

                //Point securityPoint = new Point(x, y + 90);
                var securityPoint = new Point(x, y + 75);
                Input.MoveTo(securityPoint);

                Thread.Sleep(2000);
                Input.MoveToAndClick(securityPoint);

                Thread.Sleep(3000);

                Epia3Common.WriteTestLogMsg(slogFilePath, "Epia shutdown:", sOnlyUITest);
                // Check total number of main form
                // Set a property condition that will be used to find the control.
                Condition c = new PropertyCondition(
                    AutomationElement.AutomationIdProperty, "MainForm", PropertyConditionFlags.IgnoreCase);

                Console.WriteLine(" find total mainForm :");

                // Find the element.
                AutomationElementCollection aes =
                    AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                Thread.Sleep(10000);

                DateTime mAppTime = DateTime.Now;
                TimeSpan Time = DateTime.Now - mAppTime;
                while (aes.Count != 0 && Time.Minutes < 2)
                {
                    Console.WriteLine("Find Application aeForm aes.Count: " + aes.Count);
                    aes = AutomationElement.RootElement.FindAll(TreeScope.Element | TreeScope.Children, c);
                    Time = DateTime.Now - mAppTime;
                    Console.WriteLine(" find time is :" + Time.TotalMilliseconds/1000);
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

        #region ShellLogoff

        public static void ShellLogoff(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;

            Thread.Sleep(5000);

            try
            {
                AutomationEventHandler UIACurrentUserEventHandler = OnUIACurrentUserEvent;
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIACurrentUserEventHandler);
                sEventEnd = false;
                TestCheck = ConstCommon.TEST_PASS;

                //string path = Path.Combine(sInstallScriptsDir, Constants.SHELL_BAT);
                string path = Path.Combine(m_SystemDrive + ConstCommon.EPIA_CLIENT_ROOT,
                                           ConstCommon.EGEMIN_EPIA_SHELL_EXE);
                Process proc = Process.Start(path);
                Console.WriteLine("*****" + proc.Id);
                Thread.Sleep(9000);

                // Start Shell
                //TestTools.Utilities.StartProcessNoWait(m_SystemDrive + ConstCommon.EPIA_CLIENT_ROOT,
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
                    Console.WriteLine("Find Application aeSecurityForm : " + DateTime.Now);
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
                string tester = WindowsIdentity.GetCurrent().Name;


                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser, UserPassword,
                                                           ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + PasswordID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester,
                                                           ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + UserNameID;
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

                Thread.Sleep(5000);
                AutomationElement aeMainForm = null;
                Process ShellProcess = null;

                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                int pID = 0;
                while (pID == 0 && mTime.TotalMilliseconds < 120000)
                {
                    pID = ProcessUtilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out ShellProcess);
                    Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                    mTime = DateTime.Now - mStartTime;
                }

                Console.WriteLine("Proc ID:" + pID);
                Thread.Sleep(3000);

                //aeMainForm = AutomationElement.FromHandle(ShellProcess.MainWindowHandle);

                string formID = "MainForm";
                DateTime mAppTime = DateTime.Now;
                AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeMainForm, formID, mAppTime,
                                                       300);
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
                double y = (aeShell.Current.BoundingRectangle.Top + aeShell.Current.BoundingRectangle.Bottom)/2;
                var shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Thread.Sleep(3000);

                double y2 = y + 50;
                var securityPoint = new Point(x, y2);
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
                tester = WindowsIdentity.GetCurrent().Name;

                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser, UserPassword,
                                                           ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + PasswordID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester,
                                                           ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + UserNameID;
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
        }

        #endregion

        #region Epia4Close

        public static void Epia4Close(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            string BtnCloseID = "Close";
            try
            {
                Thread.Sleep(5000);
                // Close the other mainForms
                Process[] pShell = Process.GetProcessesByName(ConstCommon.EGEMIN_EPIA_SHELL);
                for (int i = 0; i < pShell.Length; i++)
                {
                    AutomationElement aeMainForm = AutomationElement.FromHandle(pShell[i].MainWindowHandle);
                    AutomationElement aeCloses = AUIUtilities.FindElementByID(BtnCloseID, aeMainForm);
                    if (aeCloses == null)
                    {
                        sErrorMessage = "failed to find aeCloses of mainForm";
                        Console.WriteLine(testname + " failed to find aeCloses at time: " +
                                          DateTime.Now.ToString("HH:mm:ss"));
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                    else
                    {
                        Console.WriteLine(testname + " aeClose found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                        var ipc =
                            aeCloses.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        ipc.Invoke();
                    }

                    Thread.Sleep(5000);
                }

                // Close the other LogonForms
                Process[] pLogon = Process.GetProcessesByName(ConstCommon.EGEMIN_EPIA_SHELL);
                for (int i = 0; i < pLogon.Length; i++)
                {
                    AutomationElement aeLogonForm = AutomationElement.FromHandle(pLogon[i].MainWindowHandle);
                    AutomationElement aeCancel = AUIUtilities.FindElementByID("m_BtnCancel", aeLogonForm);
                    if (aeCancel == null)
                    {
                        sErrorMessage = "failed to find aeCancel of LogonForm";
                        Console.WriteLine(testname + " failed to find aeCloses at time: " +
                                          DateTime.Now.ToString("HH:mm:ss"));
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                    else
                    {
                        Console.WriteLine(testname + " aeCancel found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                        var ipc =
                            aeCancel.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        ipc.Invoke();
                    }

                    Thread.Sleep(5000);
                }

                Process proc = null;
                int pID = ProcessUtilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out proc);
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

        #endregion Epia4Close

        #region ReturnDefaultLayout

        public static void ReturnDefaultLayout(AutomationElement root, out int result)
        {
            string testname = "ReturnDefaultLayout";
            Console.WriteLine("\n=== Test ReturnDefaultLayout ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutStandardScreenEventHandler
                = OnLayoutStandardScreenUIAEvent;

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                                                                   ConstCommon.MY_LAYOUT,
                                                                                   ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutStandardScreenEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                             sOnlyUITest);
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

        #region ValidateDefaultLayout

        public static void ValidateDefaultLayout(AutomationElement root, out int result)
        {
            string testname = "ValidateDefaultLayout";
            Console.WriteLine("\n=== Test ValidateDefaultLayout ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            AutomationEventHandler UIALayoutValidateDefaultEventHandler
                = OnLayoutValidateDefaultUIAEvent;

            try
            {
                AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
                                                                                   ConstCommon.MY_LAYOUT,
                                                                                   ref sErrorMessage);
                if (aeYourLayouts == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutValidateDefaultEventHandler);

                sEventEnd = false;
                Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(5000);
                while (root.Current.IsEnabled)
                {
                    Console.WriteLine("re click my layout :");
                    Input.MoveToAndClick(Pnt);
                    Thread.Sleep(5000);
                }

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

                Console.WriteLine("time is:" + mTime.TotalMilliseconds/1000);
                Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds/1000, sOnlyUITest);

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
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                             sOnlyUITest);
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

        #region LayoutMainMenuOn

        public static void LayoutMainMenuOn(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
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
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                            sOnlyUITest);
            }
        }

        #endregion

        #region LayoutOpen

        public static void LayoutOpen(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            string BtnOpenID = "m_BtnOpen";
            TestCheck = ConstCommon.TEST_PASS;

            try
            {
                // Find Open Button Element
                Condition c = new PropertyCondition(
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
                AutomationEventHandler UIALayoutOpenEventHandler = OnLayoutOpenUIAEvent;
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                                     AutomationElement.RootElement, TreeScope.Descendants,
                                                     UIALayoutOpenEventHandler);

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
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                            sOnlyUITest);
            }
        }

        #endregion

        #region ConfigOpen

        public static void ConfigOpen(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            string BtnOpenID = "m_BtnOpen";
            try
            {
                // Add Open  window Event Handler
                AutomationEventHandler UIAConfigOpenEventHandler = OnConfigOpenUIAEvent;
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
                    Console.WriteLine("\nName should be " + sConfigurationName + " , But it is: " +
                                      aeLbLConfiguration.Current.Name);
                    Console.WriteLine("\nTest scenario Open: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                            sOnlyUITest);
            }
        }

        #endregion

        #region ConfigSave

        public static void ConfigSave(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
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
                    Console.WriteLine("ap[i].ToString():" + ap[i]);
                }
                Thread.Sleep(3000000);

                // Add Open  window Event Handler
                AutomationEventHandler UIAConfigOpenEventHandler = OnConfigOpenUIAEvent;
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
                    Console.WriteLine("\nName should be " + sConfigurationName + " , ButôÎit is: " +
                                      aeLbLConfiguration.Current.Name);
                    Console.WriteLine("\nTest scenario Open: *FAIL*");
                    result = ConstCommon.TEST_FAIL;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                            sOnlyUITest);
            }
        }

        #endregion

        #region ConfigSecurityUnsecured

        public static void ConfigSecurityUnsecured(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
            result = ConstCommon.TEST_UNDEFINED;
            string BtnOpenID = "m_BtnOpen";
            try
            {
                // Add Open  window Event Handler
                AutomationEventHandler UIAConfigSecurityUnsecuredEventHandler = OnConfigSecurityUnsecuredUIAEvent;
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

                var wpCloseForm =
                    (WindowPattern) root.GetCurrentPattern(WindowPattern.Pattern);
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

                string dir = m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Epia Shell";
                string procShellExe = "Egemin.Epia.Shell.exe";
                string procName = "Egemin.Epia.Shell";
                bool startApp = ProcessUtilities.StartProcessAndWaitUntilUIWindowFound(dir, procShellExe, null, procName,
                                                                                       2, ref aeForm);

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
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                            sOnlyUITest);
            }
        }

        #endregion

        #region ConfigFindGridView

        public static void ConfigFindGridView(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine(Constants.TestLogHeader + testname);
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            ;
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
                    var patterns = (GridPattern) agv.GetCurrentPattern(GridPattern.Pattern);

                    Console.WriteLine("current pattern is " + patterns);

                    //patterns.Invoke();
                    Thread.Sleep(300000);

                    var gP = (TablePattern) agv.GetCurrentPattern(TablePattern.Pattern);
                    gP.GetItem(1, 1).ToString();

                    Console.WriteLine("GRIDVIEW (1,1)is " + gP.GetItem(1, 1));
                }
                Thread.Sleep(300000);

                // Add Open  window Event Handler
                AutomationEventHandler UIAConfigSecurityEpiaEventHandler = OnConfigSecurityEpiaUIAEvent;
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

                var wpCloseForm =
                    (WindowPattern) root.GetCurrentPattern(WindowPattern.Pattern);
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
                    Console.WriteLine("\nName should be " + LogonScreenTitle + " , But it is: " +
                                      aeLogonScreen.Current.Name);
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
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace,
                                            sOnlyUITest);
            }
        }

        #endregion ConfigFindGridView

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        #endregion Test Cases ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        #region Event ------------------------------------------------------------------------------------------------

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
                Condition c = new PropertyCondition(
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
                            var tp = (TextPattern) aeTxt.GetCurrentPattern(TextPattern.Pattern);
                            Thread.Sleep(1000);
                            sErrorMessage = tp.DocumentRange.GetText(-1).Trim();
                            Console.WriteLine("Error Message Catched ------------:");
                            //Console.WriteLine("Error Message is ------------:" + sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, "start shell failed: " + sErrorMessage,
                                                         sOnlyUITest);
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
                Condition c = new AndCondition(
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
            else if (name.Equals("Egemin.Epia.Server"))
            {
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Condition c = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Don't Send"),
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
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Epia Server startup failed: " + name, sOnlyUITest);
                }
                else
                {
                    Condition c2 = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, "Close"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                        );

                    // Find the element.
                    aeRun = element.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                    if (aeRun != null)
                    {
                        Point pt = AUIUtilities.GetElementCenterPoint(aeRun);
                        Input.MoveTo(pt);
                        Thread.Sleep(1000);
                        Input.ClickAtPoint(pt);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Epia Server startup failed: " + name, sOnlyUITest);
                    }
                }

                #region update quality

                if (!sOnlyUITest)
                {
                    if (TFSConnected)
                    {
                        Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                                                                          TfsUtilities.GetProjectName(ConstCommon.EPIA),
                                                                          sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                        if (quality.Equals("GUI Tests Failed"))
                        {
                            //
                            // check testinfo file line count if Line count = 2, and if sTotalFailed == 0 ---> update status to QUI Test Passed
                            // this means I have deleted testinfo file due to the previous test failure; and rerun this test again. and if OK I should
                            // change this satus to 'Pass'
                            //_---
                            string path = sBuildDropFolder + "\\TestResults";
                            string testPC = Environment.MachineName;

                            // read allline from test info file
                            int count = 0;
                            StreamReader reader = File.OpenText(Path.Combine(path, ConstCommon.TESTINFO_FILENAME));
                            string infoline = reader.ReadLine();
                            Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " read line :" + infoline, sOnlyUITest);

                            while (infoline != null && infoline.Length > 0)
                            {
                                Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " read line while :" + infoline,
                                                            sOnlyUITest);
                                count++;
                                infoline = reader.ReadLine();
                            }
                            reader.Close();

                            Epia3Common.WriteTestLogMsg(slogFilePath,
                                                        sBuildNr + " number of line in testinfo.txt file :" + count,
                                                        sOnlyUITest);


                            /*if (count == 2)
							{    
								if (sTotalFailed == 0)
									TestTools.TfsUtilities.UpdateBuildQualityStatusEvenHasFailedStatus(logger, uri,
									TestTools.TfsUtilities.GetProjectName(ConstCommon.EPIA),
									"GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
							}*/
                            //---------
                            Epia3Common.WriteTestLogMsg(slogFilePath,
                                                        sBuildNr + " has failed quality, no update needed :" + quality,
                                                        sOnlyUITest);
                        }
                        else
                        {
                            //if (sTotalFailed == 0)
                            //    TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                            //    TestTools.TfsUtilities.GetProjectName(ConstCommon.EPIA),
                            //    "GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
                            //else
                            TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                                                  TfsUtilities.GetProjectName(ConstCommon.EPIA),
                                                                  "GUI Tests Failed", m_BuildSvc,
                                                                  sDemo ? "true" : "false");
                        }

                        // update testinfo file
                        string testout = "-->" + sOutFilename + ".xls";
                        if (sAutoTest)
                        {
                            if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
                                || PCName.ToUpper().Equals("EPIATESTSRV3V1"))
                                testout = "-->" + sOutFilename + ".log";

                            //if (sTotalFailed == 0)
                            //{
                            //    FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Passed" + testout, ConstCommon.EPIA);
                            //    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Passed:" + ConstCommon.EPIA, sOnlyUITest);
                            //}
                            //else
                            //{

                            TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Failed", testout,
                                                                         sInfoFileKey);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Failed:" + ConstCommon.EPIA,
                                                        sOnlyUITest);
                            //}
                        }
                    }

                    if (sAutoTest)
                        FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                }

                #endregion

                Environment.Exit(1);
            }
            else
            {
                Console.WriteLine("Do ELSE OTHER is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
            }
            sEventEnd = true;
        }

        #endregion

        #region OnWIBUSYSTEMSWindowOpenEvent

        public static void OnWIBUSYSTEMSWindowOpenEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnWIBU-SYSTEMS-Event");
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
            string str = string.Format("OnWIBU-SYSTEMS-Event:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);
            if (name.Equals("WIBU-SYSTEMS protected application"))
            {
                Console.WriteLine("Name is ------------:" + name);
                TestCheck = ConstCommon.TEST_PASS;
                string BtnCancelID = "buttonRight";
                try
                {
                    Thread.Sleep(1000);
                    AUIUtilities.FindElementAndClick(BtnCancelID, element);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("OnWIBUSYSTEMSWindowOpenEvent :" + ex.Message + " --- " + ex.StackTrace);
                }
                sEventEnd = true;
            }
            else if (name.Equals("Error"))
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
                Console.WriteLine("-------   OTHER Name is ------------:" + name);
                Thread.Sleep(5000);
            }
            sEventEnd = true;
        }

        #endregion

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
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalXPositionTextBoxID, element, out origXPos, "200",
                                                               ref sErrorMessage))
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
                    // Find and Click Save Button
                    if (AUIUtilities.FindElementAndClickPoint(BtnSaveID, element))
                        Thread.Sleep(5000);
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
                    Console.WriteLine("OnLayoutXPosUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }

        #endregion

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
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalYPositionTextBoxID, element, out origYPos, "100",
                                                               ref sErrorMessage))
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
                    Console.WriteLine("OnLayoutYPosUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }

        #endregion

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
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalWidthTextBoxID, element, out origWidth, "600",
                                                               ref sErrorMessage))
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
                    Console.WriteLine("OnLayoutWidthUIAEvent:" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                }
            }
            sEventEnd = true;
        }

        #endregion

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
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalHeightTextBoxID, element, out origWidth, "500",
                                                               ref sErrorMessage))
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
                    if (AUIUtilities.FindTextBoxAndChangeValue(titleTextBoxID, element, out origTitle, "abcdefg",
                                                               ref sErrorMessage))
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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkResizeID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkResizeID;
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
                    Console.WriteLine("OnLayoutFullScreenUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }

        #endregion

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
            {
                // Automation Element ID
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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkFullScreenID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkFullScreenID;
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
                    Console.WriteLine("OnLayoutFullScreenUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }

        #endregion

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
            {
                // Automation Element ID
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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkFullScreenID,
                                                     sOnlyUITest);
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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkMaximizedScreenID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkMaximizedScreenID;
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
                    Console.WriteLine("OnLayoutFullScreenUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }

        #endregion

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
            {
                // Automation Element ID   
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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowRibbonID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowRibbonID;
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
                    Console.WriteLine("OnLayoutRibbonOnUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }

        #endregion

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
            {
                // Automation Element ID   
                string BtnSaveID = "m_btnSave";
                try
                {
                    string ChkShowNavigatorID = "showNavigatorCheckBox";
                        //-------------- Navigator Off-------------------------
                    bool checkNav = AUIUtilities.FindElementAndToggle(ChkShowNavigatorID, element, ToggleState.Off);
                    if (checkNav)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowNavigatorID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowNavigatorID,
                                                     sOnlyUITest);
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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkFullScreenID,
                                                     sOnlyUITest);
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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkMaximizedID,
                                                     sOnlyUITest);
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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkResizeID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkResizeID;
                        sEventEnd = true;
                        return;
                    }

                    string InitalXPositionTextBoxID = "initialXPositionTextBox";
                        // ------------------XPos 200-----------------
                    string origXPos = "";
                    // Change XPos TxtBox
                    string getValue = string.Empty;
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalXPositionTextBoxID, element, out origXPos, "200",
                                                               ref sErrorMessage))
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindTextBoxAndChangeValue failed:" + InitalXPositionTextBoxID);
                        sErrorMessage = "FindTextBoxAndChangeValue failed:" + InitalXPositionTextBoxID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    string InitalYPositionTextBoxID = "initialYPositionTextBox";
                        // ------------------YPos 100-----------------
                    string origYPos = "";
                    // Change YPos TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalYPositionTextBoxID, element, out origYPos, "100",
                                                               ref sErrorMessage))
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
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalWidthTextBoxID, element, out origWidth, "600",
                                                               ref sErrorMessage))
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

                    string InitalHeightTextBoxID = "initialHeightTextBox";
                        // ------------------Height 700-----------------
                    string origHeight = "";
                    // Change Height TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalHeightTextBoxID, element, out origHeight, "700",
                                                               ref sErrorMessage))
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

                    string ChkShowRibbonID = "showRibbonCheckBox";
                        //-------------- Ribbon On ---------------------------
                    bool checkRibbon = AUIUtilities.FindElementAndToggle(ChkShowRibbonID, element, ToggleState.On);
                    if (checkRibbon)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowRibbonID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowRibbonID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowRibbonID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowMainMenuID = "showMainMenuCheckBox";
                        //-------------- Main Menu Off---------------------------
                    bool checkMm = AUIUtilities.FindElementAndToggle(ChkShowMainMenuID, element, ToggleState.Off);
                    if (checkMm)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowMainMenuID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowMainMenuID,
                                                     sOnlyUITest);
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

                    string ChkShowNavigatorID = "showNavigatorCheckBox";
                        //-------------- Navigator Off-------------------------
                    bool checkNav = AUIUtilities.FindElementAndToggle(ChkShowNavigatorID, element, ToggleState.Off);
                    if (checkNav)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowNavigatorID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowNavigatorID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowNavigatorID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkStackScreensID = "stackScreensCheckBox";
                        //-------------- Stack Screens On-------------------------
                    bool checkSs = AUIUtilities.FindElementAndToggle(ChkStackScreensID, element, ToggleState.On);
                    if (checkSs)
                        Thread.Sleep(500);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkStackScreensID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkStackScreensID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkStackScreensID;
                        sEventEnd = true;
                        return;
                    }

                    // Find and Click Cancel Button
                    if (AUIUtilities.FindElementAndClick(BtnCancelID, element))
                    {
                        Thread.Sleep(500);
                        sEventEnd = true;
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnCancelID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
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
                        Console.WriteLine("LanguageSettings failed to find aeCombo at time: " +
                                          DateTime.Now.ToString("HH:mm:ss"));
                        sErrorMessage = "LanguageSettings failed to find aeCombo at time: " +
                                        DateTime.Now.ToString("HH:mm:ss");
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }

                    bool select = false; //Utilities.SelectItemFromList("nl", aeCombo);
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        var selectPattern =
                            aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                        AutomationElement item
                            = AUIUtilities.FindElementByName("Nederlands", aeCombo);
                        if (item != null)
                        {
                            Console.WriteLine("LanguageSettings item found at time: " +
                                              DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(2000);

                            var itemPattern =
                                item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
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

        #region OnConfigSecurityUIAEvent

        public static void OnConfigSecurityUIAEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnConfigSecurityUIAEvent");
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
            string str = string.Format("OnConfigSecurityUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
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
                Console.WriteLine("Finding Language combo");
                try
                {
                    AutomationElement aeCombo = AUIUtilities.FindElementByID(ComboSecurityModesID, element);
                    if (aeCombo == null)
                    {
                        Console.WriteLine("ConfigSecurity failed to find aeCombo at time: " +
                                          DateTime.Now.ToString("HH:mm:ss"));
                        sErrorMessage = "ConfigSecurity failed to find aeCombo at time: " +
                                        DateTime.Now.ToString("HH:mm:ss");
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
                        var selectPattern =
                            aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                        AutomationElement item
                            = AUIUtilities.FindElementByName("Windows gebruiker", aeCombo);
                        if (item != null)
                        {
                            Console.WriteLine("ConfigSecurity item found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(2000);

                            var itemPattern =
                                item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
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
                    Console.WriteLine("OnConfigSecurityUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
                    sErrorMessage = ex.Message + " --- " + ex.StackTrace;
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }

        #endregion

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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkFullScreenID,
                                                     sOnlyUITest);
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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkMaximizedID,
                                                     sOnlyUITest);
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
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkResizeID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkResizeID;
                        sEventEnd = true;
                        return;
                    }

                    string InitalXPositionTextBoxID = "initialXPositionTextBox";
                        // ------------------XPos 0-----------------
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

                    string InitalYPositionTextBoxID = "initialYPositionTextBox";
                        // ------------------YPos 0-----------------
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

                    string InitalHeightTextBoxID = "initialHeightTextBox";
                        // ------------------Height 606-----------------
                    string origHeight = "";
                    // Change Height TxtBox
                    if (AUIUtilities.FindTextBoxAndChangeValue(InitalHeightTextBoxID,
                                                               element, out origHeight, DEFAULT_HEIGHT,
                                                               ref sErrorMessage))
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

                    string ChkShowRibbonID = "showRibbonCheckBox";
                        //-------------- Ribbon OFF ---------------------------
                    bool checkRibbon = AUIUtilities.FindElementAndToggle(ChkShowRibbonID, element, ToggleState.Off);
                    if (checkRibbon)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowRibbonID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowRibbonID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowRibbonID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowMainMenuID = "showMainMenuCheckBox";
                        //-------------- Main Menu ON---------------------------
                    bool checkMm = AUIUtilities.FindElementAndToggle(ChkShowMainMenuID, element, ToggleState.On);
                    if (checkMm)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowMainMenuID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowMainMenuID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowMainMenuID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowToolBarsID = "showToolBarsCheckBox";
                        //-------------- Tool bars OFF---------------------------
                    bool checktb = AUIUtilities.FindElementAndToggle(ChkShowToolBarsID, element, ToggleState.Off);
                    if (checktb)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowToolBarsID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowToolBarsID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowToolBarsID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkShowNavigatorID = "showNavigatorCheckBox";
                        //-------------- Navigator ON-------------------------
                    bool checkNav = AUIUtilities.FindElementAndToggle(ChkShowNavigatorID, element, ToggleState.On);
                    if (checkNav)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkShowNavigatorID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkShowNavigatorID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkShowNavigatorID;
                        sEventEnd = true;
                        return;
                    }

                    string ChkStackScreensID = "stackScreensCheckBox";
                        //-------------- Stack Screens OFF-------------------------
                    bool checkSs = AUIUtilities.FindElementAndToggle(ChkStackScreensID, element, ToggleState.Off);
                    if (checkSs)
                        Thread.Sleep(300);
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + ChkStackScreensID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkStackScreensID,
                                                     sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "FindElementAndToggle failed:" + ChkStackScreensID;
                        sEventEnd = true;
                        return;
                    }

                    // Find and Click Save Button
                    if (AUIUtilities.FindElementAndClickPoint(BtnSaveID, element))
                    {
                        Thread.Sleep(300);
                        sEventEnd = true;
                    }
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
                    sErrorMessage = "OnLayoutStandardScreenUIAEvent" + ex.Message + " --- " + ex.StackTrace;
                    Console.WriteLine("OnLayoutStandardScreenUIAEvent" + ex.Message + " --- " + ex.StackTrace);
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                }
            }
            sEventEnd = true;
        }

        #endregion

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
                    ToggleState fstg = AUIUtilities.FindCheckBoxAndToggleState(ChkFullScreenID, element,
                                                                               ref sErrorMessage);
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
                    ToggleState mstg = AUIUtilities.FindCheckBoxAndToggleState(ChkMaximizedID, element,
                                                                               ref sErrorMessage);
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

                    string InitalXPositionTextBoxID = "initialXPositionTextBox";
                        // ------------------XPos 0-----------------
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

                    string InitalYPositionTextBoxID = "initialYPositionTextBox";
                        // ------------------YPos 0-----------------
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

                    string InitalHeightTextBoxID = "initialHeightTextBox";
                        // ------------------Height 606-----------------
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

                    string ChkShowRibbonID = "showRibbonCheckBox";
                        //-------------- Ribbon OFF ---------------------------
                    ToggleState rbtg = AUIUtilities.FindCheckBoxAndToggleState(ChkFullScreenID, element,
                                                                               ref sErrorMessage);
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

                    string ChkShowMainMenuID = "showMainMenuCheckBox";
                        //-------------- Main Menu ON---------------------------
                    ToggleState mmtg = AUIUtilities.FindCheckBoxAndToggleState(ChkShowMainMenuID, element,
                                                                               ref sErrorMessage);
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

                    string ChkShowToolBarsID = "showToolBarsCheckBox";
                        //-------------- Tool bars OFF---------------------------
                    ToggleState tbtg = AUIUtilities.FindCheckBoxAndToggleState(ChkShowToolBarsID, element,
                                                                               ref sErrorMessage);
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

                    string ChkShowNavigatorID = "showNavigatorCheckBox";
                        //-------------- Navigator ON-------------------------
                    ToggleState nvtg = AUIUtilities.FindCheckBoxAndToggleState(ChkShowNavigatorID, element,
                                                                               ref sErrorMessage);
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

                    string ChkStackScreensID = "stackScreensCheckBox";
                        //-------------- Stack Screens OFF-------------------------
                    ToggleState sstg = AUIUtilities.FindCheckBoxAndToggleState(ChkStackScreensID, element,
                                                                               ref sErrorMessage);
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
                    if (AUIUtilities.FindElementAndClickPoint(BtnCancelID, element))
                    {
                        Thread.Sleep(500);
                        sEventEnd = true;
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndClick failed:" + BtnCancelID);
                        sErrorMessage = "FindElementAndClick failed:" + BtnCancelID;
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
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
                Condition c = new AndCondition(
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
                        Console.WriteLine("ConfigSecurity failed to find aeCombo at time: " +
                                          DateTime.Now.ToString("HH:mm:ss"));
                        sErrorMessage = "ConfigSecurity failed to find aeCombo at time: " +
                                        DateTime.Now.ToString("HH:mm:ss");
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
                        var selectPattern =
                            aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                        AutomationElement item
                            = AUIUtilities.FindElementByName("EpiaMemberOrAnyWindowsUser", aeCombo);
                        if (item != null)
                        {
                            Console.WriteLine("ConfigSecurity item found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(2000);

                            var itemPattern =
                                item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
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
                    Condition cF = new AndCondition(
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
                        var sip = (SelectionItemPattern) aeindex.GetCurrentPattern(SelectionItemPattern.Pattern);
                        sip.Select();
                    }

                    Rect rect;
                    rect = (Rect) (aeindex.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
                    var pt = new Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
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
                MessageBox.Show("OnLayoutOpenUIAEvent:" + ex.Message + "  --   " + ex.StackTrace);
                return;
            }
        }

        #endregion

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
                    Condition cF = new AndCondition(
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
                        var sip = (SelectionItemPattern) aeindex.GetCurrentPattern(SelectionItemPattern.Pattern);
                        sip.Select();
                    }

                    Rect rect;
                    rect = (Rect) (aeindex.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
                    var pt = new Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
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
                if (AUIUtilities.FindTextBoxAndChangeValue(TxtLayoutID, element, out getValue,
                                                           "TestLayoutSaveAs@EPIA3TESTPC", ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    Console.WriteLine("FindTextBoxAndChangeValue failed:" + TxtLayoutID);
                    Epia3Common.WriteTestLogFail(slogFilePath,
                                                 "FindTextBoxAndChangeValue failed:" + TxtLayoutID + " --- " +
                                                 sErrorMessage, sOnlyUITest);
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
                if (AUIUtilities.FindTextBoxAndChangeValue(TxtConfigurationID, element, out getValue, "TestConfigSaveAs",
                                                           ref sErrorMessage))
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
                Condition cF = new AndCondition(
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
                    var sip = (SelectionItemPattern) aeindex.GetCurrentPattern(SelectionItemPattern.Pattern);
                    sip.Select();
                }

                Rect rect;
                rect = (Rect) (aeindex.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
                var pt = new Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
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
                        Epia3Common.WriteTestLogMsg(slogFilePath, "FindElementByID failed:" + ListConfigurationID,
                                                    sOnlyUITest);
                    }
                    //aeLauoutIDs.SetFocus();

                    // find all listitem
                    Condition cF = new AndCondition(
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
                        var sip = (SelectionItemPattern) aeindex.GetCurrentPattern(SelectionItemPattern.Pattern);
                        sip.Select();
                    }

                    Rect rect;
                    rect = (Rect) (aeindex.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
                    var pt = new Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
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
                Epia3Common.WriteTestLogMsg(slogFilePath,
                                            "OnConfigSecurityEpiaUIAEvent exception:" + ex.Message + "----" +
                                            ex.StackTrace, sOnlyUITest);
                Console.WriteLine("OnConfigSecurityEpiaUIAEvent exception:" + ex.Message + "----" + ex.StackTrace);
            }
        }

        #endregion

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
                        Epia3Common.WriteTestLogMsg(slogFilePath, "FindElementByID failed:" + ListConfigurationID,
                                                    sOnlyUITest);
                    }
                    //aeLauoutIDs.SetFocus();

                    // find all listitem
                    Condition cF = new AndCondition(
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
                        var sip = (SelectionItemPattern) aeindex.GetCurrentPattern(SelectionItemPattern.Pattern);
                        sip.Select();
                    }

                    Rect rect;
                    rect = (Rect) (aeindex.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
                    var pt = new Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
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
                Epia3Common.WriteTestLogMsg(slogFilePath,
                                            "OnConfigSecurityUnsecuredUIAEvent exception:" + ex.Message + "----" +
                                            ex.StackTrace, sOnlyUITest);
                Console.WriteLine("OnConfigSecurityUnsecuredUIAEvent exception:" + ex.Message + "----" + ex.StackTrace);
            }
        }
        #endregion

        
        #endregion Event +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        public static void SendEmail(string resultFile)
        {
            string str1 = DeployUtilities.GetTestReportContentString(sTotalCounter, sTotalPassed, sTotalFailed, sTotalException, sTotalUntested,
              sCurrentPlatform, sInstallMsiDir); // AnyCPU 

            ProcessUtilities.SendTestResultToDevelopers(resultFile, sTestApp, sBuildDef, logger, sTotalFailed,
                                                        sBuildNr /*used for email title*/, str1 /*content*/, sSendMail);
        }
    }
}