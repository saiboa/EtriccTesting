using System;
using System.IO;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Xml;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using TestTools;
using Excel = Microsoft.Office.Interop.Excel;
using TFSQATestTools;

namespace QATestEtriccUI
{
    class EtriccUIProgram
    {
        [DllImport("mpr.dll")]
        public static extern int WNetCancelConnection2A(string sharename, int dwFlags, int fForce);

        static TestTools.Logger logger;
        static public string m_CurrentDrive = @"C:\";
        // PCinfo
        static public string PCName;
        static public string OSName;
        static public string OSVersion;
        static public string UICulture;
        // Build param ========================================================
        static IBuildServer m_BuildSvc = null;
        static bool TFSConnected = true;
        //static BuildStore buildStore = null;
        static string sInstallMsiDir =  @"C:\LocalTest";
        static string sBuildDropFolder = string.Empty;
        static string sTestApp = string.Empty;
        static string sBuildDef = string.Empty;
        static string sBuildConfig = string.Empty;
        static string sBuildNr = string.Empty;
        static string sTestToolsVersion = string.Empty;
        static string sEtriccServerRoot = string.Empty;
        static string sParentProgram = string.Empty;
        static string sTestResultFolder = string.Empty;
        static string sTargetPlatform = string.Empty;
        static string sCurrentPlatform = string.Empty;
        // Testcase not used ==================================================
        public static string sConfigurationName = string.Empty;
        public static string sLayoutName = string.Empty;
        // LOG=================================================================
        public static string slogFilePath = @"C:\EtriccTests";
        private static string sErrorMessage = string.Empty;
        static string sExcelVisible = string.Empty;
        static string sServerRunAs = "Service";
        static string sOutFilename = string.Empty;
        static string sOutFilePath = string.Empty;
        static StreamWriter Writer;

        // Test Param. =========================================================
        static string sTFSServer = "http://Team2010App.TeamSystems.Egemin.Be:8080";
        static string sProjectFile = "demo.xml";
        static AutomationElement aeForm = null;
        static int Counter = 0;
        static string[] sTestCaseName = new string[100];
        static DateTime sTestStartUpTime = DateTime.Now;
        static int sTotalCounter = 0;
        static int sTotalException = 0;
        static int sTotalFailed = 0;
        static int sTotalPassed = 0;
        static int sTotalUntested = 0;
        static int TestCheck = ConstCommon.TEST_UNDEFINED;
        static public string TimeOnPC;
        static bool sEventEnd = false;
        static bool sAutoTest = true;
        static bool sFunctionalTest = true;
        static bool sDemo = false;
        static string sSendMail = "false";
        static string m_SystemDrive = @"C:\";

        static DateTime sStartTime = DateTime.Now;
        static TimeSpan sTime;

        static string sFuncTotalFailed = "0";
        private static int sNumAgvs = 2;

        static bool sOnlyUITest = false;
        static string sTestType = "all";
        static string sTestDefinitionFile = string.Empty;
        static string[] mTestDefinitionTypes;
        static string sInfoFileKey = string.Empty;
        static string sNetworkMap = "LocalTest";

        // excel 	--------------------------------------------------------
        static Excel.Application xApp;
        static int sHeaderContentsLength = 11;
        static Excel.Application xAppFunc;

        #region TestCase Name
        private const string SHELL_CLOSEE_WITHIN_MIN_AFTER_SERVER_DOWN = "ShellCloseWithinMinuteAfterServerDown";
        private const string START_EPIA_ETRICC_SERVER_SHELL = "StartEpiaEtriccServerAndShell";
        private const string RESOURCEID_INTEGRITY_CHECK = "ResourceIDIntegrityCheck";
        private const string PLAYBACK_INSTALL_AND_STARTUP = "PlaybackInstallStartup";
        private const string HOSTTEST_INSTALL_AND_STARTUP = "HostTestInstallStartup";
        private const string DISPLAY_SYSTEM_OVERVIEW = "SystemOverviewDisplay";
        private const string DISPLAY_AGV_OVERVIEW = "AgvOverviewDisplay";
        private const string DISPLAY_LOCATION_OVERVIEW_START_NODE = "LocationOverviewScreenSelectNodeAsStart";
        private const string DISPLAY_LOCATION_OVERVIEW_END_NODE = "LocationOverviewScreenSelectNodeAsEnd";
        private const string DISPLAY_STATION_OVERVIEW_START_NODE = "StationOverviewScreenSelectNodeAsStart";
        private const string DISPLAY_STATION_OVERVIEW_END_NODE = "StationOverviewScreenSelectNodeAsEnd";
        private const string DISPLAY_TRANSPORT_OVERVIEW = "TransportOverviewDisplay";
        private const string MULTI_LANGUAGE_CHECK = "MultiLanguageCheck";
        private const string LOCATION_OVERVIEW_OPEN_DETAIL = "LocationOverviewOpenDetail";
        private const string LOCATION_MODE_MANUAL = "LocationModeManual";
        private const string AGV_OVERVIEW_OPEN_DETAIL = "AgvOverviewOpenDetail";
        private const string AGV_JOB_OVERVIEW = "AgvJobsOverview";
        private const string AGV_JOB_OVERVIEW_OPEN_DETAIL = "AgvJobOverviewOpenDetail";
        private const string AGV_ACTION_MENUITEMS_VALIDATION = "AgvActionMenuitemsValidation";
        private const string STATION_ACTION_MENUITEMS_VALIDATION = "StationActionMenuitemsValidation";
        private const string STATION_MODE_CONTROL = "StationModeControl";
        private const string AGV_RESTART = "AgvRestart";
        private const string AGV_MODE = "AgvMode";
        private const string AGV_ENGINEERING_SIMULATION = "AgvEngineringSimulation";
        private const string CREATE_NEW_TRANSPORT = "TransportCreateNew";
        private const string EDIT_TRANSPORT = "TransportEdit";
        private const string SUSPEND_TRANSPORT = "TransportSuspend";
        private const string RELEASE_TRANSPORT = "TransportRelease";
        private const string CANCEL_TRANSPORT = "TransportCancel";
        private const string TRANSPORT_OVERVIEW_OPEN_DETAIL = "TransportOverviewOpenDetail";
        private const string AGV_OVERVIEW_REMOVE_ALL = "AgvsAllModeRemoved";
        private const string AGV_OVERVIEW_ID_SORTING = "AgvsIdSorting";
        private const string SYSTEM_OVERVIEW_QUERY = "SystemOverviewQuery";
        private const string EPIA4_CLOSE = "Epia4Close";
        private const string ETRICC_PLAYBACK_CHECK = "EtriccPlaybackCheck";
        private const string ETRICC_HOSTTEST_CHECK = "EtriccHostTest";
        private const string SCRIPT_BEFORE_AFTER_ACTIVATE = "BeforeAfterActivateScript";
        private const string SCRIPT_BEFORE_AFTER_DEACTIVATE = "BeforeAfterDeactivateScript";
        private const string ETRICC_EXPLORER_OVERVIEW = "EtriccExplorerOverview";
        private const string ETRICC_EXPLORER_LOAD_PROJECT = "EtriccExplorerLoadProject";
        private const string ETRICC_EXPLORER_EDIT_SAVE_PROJECT = "EtriccExplorerEditSaveSampleProject";
        private const string ETRICC_EXPLORER_BUILD_PROJECT = "EtriccExplorerBuildProject";
        private const string ETRICC_EXPLORER_CLOSE = "EtriccExplorerClose";
        private const string ETRICC_EXPLORER_VALIDATE_NEW_BUILD = "EtriccExplorerValidateNewProject";
        #endregion TestCase Name

        private const string INFRASTRUCTURE = "E'tricc®";
        private const string SYSTEM_OVERVIEW = "System Overview";
        private const string AGV_OVERVIEW = "Agvs";
        private const string LOCATION_OVERVIEW = "Locations";
        private const string STATION_OVERVIEW = "Stations";
        private const string TRANSPORT_OVERVIEW = "Transports";
        private const string NEW_TRANSPORT = "New Transport";

        private const string SYSTEM_OVERVIEW_TITLE = "System overview";
        private const string AGV_OVERVIEW_TITLE = "Agvs";
        private const string LOCATION_OVERVIEW_TITLE = "Locations";
        private const string STATION_OVERVIEW_TITLE = "Stations";
        private const string TRANSPORT_OVERVIEW_TITLE = "Transports";

        private const string DATAGRIDVIEW_ID = "m_GridData";
        private const string AGV_GRIDDATA_ID = "m_GridData";
        private const string MESSAGESCREEN_ID = "MessageScreen";
        private static string sScreenResolution = string.Empty;

        // Test Case Status. =========================================================
        private static bool sEtriccServerStartupOK = true;
        private static bool sEtriccExplorerStartupOK = true;
        private static bool sTransportSuspendOK = false;
        private static bool sBeforeAfterActivateScriptOK = true;

        private static bool sExploreProjectLoadOK = true;

        private static TfsTeamProjectCollection tfsProjectCollection;
        static AutomationEventHandler sUIAShellEventHandler;

        [STAThread]
        static void Main(string[] args)
        {
            try  // Get test PC info======================================
            {
                m_CurrentDrive = Path.GetPathRoot(Directory.GetCurrentDirectory());
                HelpUtilities.SavePCInfo("y");
                HelpUtilities.GetPCInfo(out PCName, out OSName, out OSVersion, out UICulture, out TimeOnPC);
                Console.WriteLine("<PCName : " + PCName + ">, <OSName : " + OSName + ">, <OSVersion : " + OSVersion + ">");
                Console.WriteLine("<TimeOnPC : " + TimeOnPC + ">, <UICulture : " + UICulture + ">");

                int w = System.Windows.Forms.SystemInformation.VirtualScreen.Width;
                int h = System.Windows.Forms.SystemInformation.VirtualScreen.Height;
                sScreenResolution = w + "x" + h;
                //System.Windows.MessageBox.Show("screen resolution: " + w + "x" + h);

                #region // unzip EticcTests.zip
                if (System.IO.Directory.Exists(@"C:\EtriccTests\EtriccUI"))
                {
                    Console.WriteLine(@"C:\EtriccTests\EtriccUI folder exist, do nothing: ");
                }
                else
                {
                    Console.WriteLine(@"C:\EtriccTests\EtriccUI folder not exist, unzip test scripts data: ");
                    try
                    {
                        //string zipFile = EticcTests.zip;
                        string zipFile = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "EtriccTests.zip");
                        FastZip fz = new FastZip();
                        fz.ExtractZip(zipFile, @"C:\", "");
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show("EtriccTests.zip unzip exception: " + ex.Message + "--------------" + ex.StackTrace);
                    }
                    Thread.Sleep(5000);
                }
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            sOnlyUITest = false;
            string x = System.Configuration.ConfigurationManager.AppSettings.Get("OnlyUITest");
            if (x.ToLower().StartsWith("true"))
            {
                sOnlyUITest = true;
                ProcessUtilities.CloseProcess("EXCEL");
            }
                

            if (!sOnlyUITest)
            {
                try
                {
                    // validate inputs
                    if (args != null)
                    {
                        for (int i = 0; i <= 21; i++)
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

                        if (args[16].StartsWith("true"))
                            sFunctionalTest = true;
                        else
                            sFunctionalTest = false;

                        sProjectFile = args[17];
                        sEtriccServerRoot = args[18];

                        sTestDefinitionFile = args[19];
                        sInfoFileKey = args[20];
                        sNetworkMap = args[21];

                        sTestResultFolder = sBuildDropFolder + "\\TestResults";
                        if (!System.IO.Directory.Exists(sTestResultFolder))
                            System.IO.Directory.CreateDirectory(sTestResultFolder);

                        sOutFilename = FileManipulation.CreateOutputInfoFileName(sInfoFileKey, sAutoTest);
                        sOutFilePath = Path.Combine(sBuildDropFolder, "TestResults");
                        Console.WriteLine("sOutFilePath : " + sOutFilePath);

                        Epia3Common.CreateTestLog(ref slogFilePath, sOutFilePath, sOutFilename, ref Writer);

                        Epia3Common.WriteTestLogMsg(slogFilePath, "sReportDirectory : " + sTestResultFolder, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "sOutFilePath : " + sOutFilePath, sOnlyUITest);

                        Epia3Common.WriteTestLogMsg(slogFilePath, "0) Install msi file path: " + sInstallMsiDir, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "1) BuildBaseDir: " + sBuildDropFolder, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "2) build nr: " + sBuildNr, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "3) test Project: " + sProject, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "4) test Application: " + sTestApp, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "5) targeted platform: " + sTargetPlatform, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "6) current platform: " + sCurrentPlatform, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "7) test def: " + sBuildDef, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "8) Called by: " + sParentProgram, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "9) TestTool version: " + sTestToolsVersion, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "10) Auto test: " + sAutoTest, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "11) TFSServerUrl: " + sTFSServer, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "12) Server Run As: " + sServerRunAs, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "13) Excel Visible: " + sExcelVisible, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "14) Demo test: " + sDemo, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "15) Mail: " + sSendMail, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "16 Functional test: " + sFunctionalTest, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "17 Project File: " + sProjectFile, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "18 Server root: " + sEtriccServerRoot, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "19) TestDefinitionFile: " + sTestDefinitionFile, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "20) InfoFileKey: " + sInfoFileKey, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "21) NetworkMap: " + sNetworkMap, sOnlyUITest);

                        Console.WriteLine("slogFilePath : " + slogFilePath);
                        Console.WriteLine("sOutFilePath : " + sOutFilePath);
                        Console.WriteLine("sOutFilename : " + sOutFilename);
                        logger = new Logger(slogFilePath);

                        string windir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
                        m_SystemDrive = Path.GetPathRoot(windir);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "m_SystemDrive: " + m_SystemDrive, sOnlyUITest);

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
                        Console.WriteLine("16) Functional test: " + sFunctionalTest);
                        Console.WriteLine("17) Project File: " + sProjectFile);
                        Console.WriteLine("18) Server root: " + sEtriccServerRoot);
                        Console.WriteLine("19) TestDefinitionFile: " + sTestDefinitionFile);
                        Console.WriteLine("20) InfoFileKey: " + sInfoFileKey);
                        Console.WriteLine("21) NetworkMap: " + sNetworkMap);

                        mTestDefinitionTypes = System.IO.File.ReadAllLines(sTestDefinitionFile);

                        for (int i = 0; i < mTestDefinitionTypes.Length; i++)
                        {
                            Console.WriteLine(i + " testdefinition : " + mTestDefinitionTypes[i]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "----" + ex.StackTrace);
                }
            }
            else
            { //log message to C:\EtriccTests
                logger = new Logger(slogFilePath);
                Epia3Common.CreateTestLog(ref slogFilePath, @"C:\EtriccTests", "x.log", ref Writer);
            }

            Console.WriteLine("Test started:");
            sTestCaseName[0] = SHELL_CLOSEE_WITHIN_MIN_AFTER_SERVER_DOWN;
            sTestCaseName[1] = START_EPIA_ETRICC_SERVER_SHELL;
            sTestCaseName[2] = RESOURCEID_INTEGRITY_CHECK;
            sTestCaseName[3] = PLAYBACK_INSTALL_AND_STARTUP;
            sTestCaseName[4] = HOSTTEST_INSTALL_AND_STARTUP;
            sTestCaseName[5] = DISPLAY_SYSTEM_OVERVIEW;
            sTestCaseName[6] = SYSTEM_OVERVIEW_QUERY;
            sTestCaseName[7] = DISPLAY_AGV_OVERVIEW;
            sTestCaseName[8] = DISPLAY_LOCATION_OVERVIEW_START_NODE;
            sTestCaseName[9] = DISPLAY_LOCATION_OVERVIEW_END_NODE;
            sTestCaseName[10] = DISPLAY_STATION_OVERVIEW_START_NODE;
            sTestCaseName[11] = DISPLAY_STATION_OVERVIEW_END_NODE;
            sTestCaseName[12] = STATION_ACTION_MENUITEMS_VALIDATION;
            sTestCaseName[13] = STATION_MODE_CONTROL;
            sTestCaseName[14] = DISPLAY_TRANSPORT_OVERVIEW;
            sTestCaseName[15] = LOCATION_OVERVIEW_OPEN_DETAIL;
            sTestCaseName[16] = LOCATION_MODE_MANUAL;
            sTestCaseName[17] = AGV_OVERVIEW_OPEN_DETAIL;
            sTestCaseName[18] = AGV_ACTION_MENUITEMS_VALIDATION;
            sTestCaseName[19] = AGV_RESTART;
            sTestCaseName[20] = AGV_MODE;
            sTestCaseName[21] = AGV_JOB_OVERVIEW;
            sTestCaseName[22] = AGV_JOB_OVERVIEW_OPEN_DETAIL;
            sTestCaseName[23] = AGV_ENGINEERING_SIMULATION;
            sTestCaseName[24] = CREATE_NEW_TRANSPORT;
            sTestCaseName[25] = EDIT_TRANSPORT;
            sTestCaseName[26] = SUSPEND_TRANSPORT;
            sTestCaseName[27] = RELEASE_TRANSPORT;
            sTestCaseName[28] = CANCEL_TRANSPORT;
            sTestCaseName[29] = TRANSPORT_OVERVIEW_OPEN_DETAIL;  
            sTestCaseName[30] = AGV_OVERVIEW_REMOVE_ALL;
            sTestCaseName[31] = AGV_OVERVIEW_ID_SORTING;
            sTestCaseName[32] = SCRIPT_BEFORE_AFTER_ACTIVATE;
            sTestCaseName[33] = SCRIPT_BEFORE_AFTER_DEACTIVATE;
            sTestCaseName[34] = MULTI_LANGUAGE_CHECK;
            sTestCaseName[35] = EPIA4_CLOSE;
            sTestCaseName[36] = ETRICC_EXPLORER_OVERVIEW;
            sTestCaseName[37] = ETRICC_EXPLORER_LOAD_PROJECT;
            sTestCaseName[38] = ETRICC_EXPLORER_EDIT_SAVE_PROJECT;
            sTestCaseName[39] = ETRICC_EXPLORER_BUILD_PROJECT;
            sTestCaseName[40] = ETRICC_EXPLORER_CLOSE;

            try
            {
                // Write Excel Header
                xApp = new Excel.Application();
                string[] sHeaderContents = { System.DateTime.Now.ToString("MMMM-dd") + "*" + "Etricc" +  " UI Test Scenarios",
                                              "Test Machine:" + "*" + PCName,
                                               "Tester::" + "*" + System.Security.Principal.WindowsIdentity.GetCurrent().Name,
                                               "OSName:" + "*" + OSName,
                                               "OS Version:" + "*" + OSVersion,
                                               "UI Culture:" + "*" + UICulture,
                                               "ProjectFile:" + "*" + sProjectFile,
                                               "Time On PC:" + "*" + "local time:" + TimeOnPC,
                                               "Test Tool Version:" + "*" +sTestToolsVersion,
                                               "NetworkMap:" + "*" +sNetworkMap,
                                                "Build Location:" + "*" +sInstallMsiDir.Substring(3),
                                          };
                
                sHeaderContentsLength = sHeaderContents.Length;
                Epia3Common.WriteTestLogMsg(slogFilePath, "sHeaderContents.length: " + sHeaderContentsLength, sOnlyUITest);
                FileManipulation.WriteExcelHeader(ref xApp, sExcelVisible, sHeaderContents);

                // start test----------------------------------------------------------
                int sResult = ConstCommon.TEST_UNDEFINED;
                int aantal = 41;

                if (sDemo)
                    aantal = 2;

                if (sOnlyUITest)   // read test case from config file, if false, test all
                {
                    aantal = 39;
                    sTestType = System.Configuration.ConfigurationManager.AppSettings.Get("TestType");
                    if (sTestType.ToLower().StartsWith("all"))
                    {
                        aantal = 41;
                    }
                    else
                    {
                        int thisTest = 0;
                        if (sTestType.IndexOf("-") > 0)
                        {
                            Console.WriteLine("first num: " + (sTestType.Substring(0, sTestType.IndexOf("-"))));
                            Console.WriteLine("second num: " + (sTestType.Substring(sTestType.IndexOf("-") + 1)));

                            thisTest = Convert.ToInt16(sTestType.Substring(0, sTestType.IndexOf("-")));
                            Counter = thisTest - 1;
                            aantal = Convert.ToInt16(sTestType.Substring(sTestType.IndexOf("-") + 1));
                        }
                        else
                        {
                            thisTest = Convert.ToInt16(sTestType);
                            aantal = 1;
                            sTestCaseName[0] = sTestCaseName[thisTest - 1];
                        }
                        sTestCaseName[0] = sTestCaseName[thisTest - 1];

                        Console.WriteLine("counter: " + Counter);
                        if (Counter < 32 && Counter > 1)
                            aeForm = EtriccUtilities.GetMainWindow("MainForm");

                    }
                }

                if (sProjectFile.IndexOf("TestProject") >= 0)
                {
                    sNumAgvs = 11;
                }
                else if (sProjectFile.ToLower().IndexOf("eurobaltic") >= 0)
                    sNumAgvs = 12;
                else
                    sNumAgvs = 2;

                while (Counter < aantal)
                {
                    //aeForm = EtriccUtilities.GetMainWindow("MainForm");
                    sResult = ConstCommon.TEST_UNDEFINED;
                    switch (sTestCaseName[Counter])
                    {
                        case SHELL_CLOSEE_WITHIN_MIN_AFTER_SERVER_DOWN:
                            ShellCloseWithinOneMinuteAfterServerDown(SHELL_CLOSEE_WITHIN_MIN_AFTER_SERVER_DOWN, aeForm, out sResult);
                            break;
                        case START_EPIA_ETRICC_SERVER_SHELL:
                            StartEpiaEtriccServerShell(START_EPIA_ETRICC_SERVER_SHELL, aeForm, out sResult);
                            break;
                        case RESOURCEID_INTEGRITY_CHECK:
                            ResourceIdIntegrityCheck(RESOURCEID_INTEGRITY_CHECK, aeForm, out sResult);
                            break;
                        case PLAYBACK_INSTALL_AND_STARTUP:
                            sErrorMessage = string.Empty;
                            //PlaybackInstallStartup(PLAYBACK_INSTALL_AND_STARTUP, aeForm, out sResult);
                            break;
                        case HOSTTEST_INSTALL_AND_STARTUP:
                            sErrorMessage = string.Empty;
                            //HostTestInstallStartup(HOSTTEST_INSTALL_AND_STARTUP, aeForm, out sResult);
                            break;
                        case DISPLAY_SYSTEM_OVERVIEW:
                            SystemOverviewDisplay(DISPLAY_SYSTEM_OVERVIEW, aeForm, out sResult);
                            break;
                        case DISPLAY_AGV_OVERVIEW:
                            AgvOverviewDisplay(DISPLAY_AGV_OVERVIEW, aeForm, out sResult);
                            break;
                        case DISPLAY_LOCATION_OVERVIEW_START_NODE:
                            LocationOverviewDisplay(DISPLAY_LOCATION_OVERVIEW_START_NODE, aeForm, out sResult);
                            break;
                        case DISPLAY_LOCATION_OVERVIEW_END_NODE:
                            LocationOverviewDisplayEndNode(DISPLAY_LOCATION_OVERVIEW_END_NODE, aeForm, out sResult);
                            break;
                        case DISPLAY_STATION_OVERVIEW_START_NODE:
                            StationOverviewDisplay(DISPLAY_STATION_OVERVIEW_START_NODE, aeForm, out sResult);
                            break;
                        case DISPLAY_STATION_OVERVIEW_END_NODE:
                            StationOverviewDisplayEndNode(DISPLAY_STATION_OVERVIEW_END_NODE, aeForm, out sResult);
                            break;
                        case STATION_ACTION_MENUITEMS_VALIDATION:
                            StationActionMenuitemsValidation(STATION_ACTION_MENUITEMS_VALIDATION, aeForm, out sResult);
                            break;
                        case STATION_MODE_CONTROL:
                            StationModeControl(STATION_MODE_CONTROL, aeForm, out sResult);
                            break;
                        case DISPLAY_TRANSPORT_OVERVIEW:
                            TransportOverviewDisplay(DISPLAY_TRANSPORT_OVERVIEW, aeForm, out sResult);
                            break;
                        case MULTI_LANGUAGE_CHECK:
                            MultiLanguageCheck(MULTI_LANGUAGE_CHECK, aeForm, out sResult);
                            break;
                        case AGV_OVERVIEW_OPEN_DETAIL:
                            AgvOverviewOpenDetail(AGV_OVERVIEW_OPEN_DETAIL, aeForm, out sResult);
                            break;
                        case LOCATION_OVERVIEW_OPEN_DETAIL:
                            LocationOverviewOpenDetail(LOCATION_OVERVIEW_OPEN_DETAIL, aeForm, out sResult);
                            break;
                        case LOCATION_MODE_MANUAL:
                            LocationModeManual(LOCATION_MODE_MANUAL, aeForm, out sResult);
                            break;
                        case AGV_JOB_OVERVIEW:
                            AgvJobOverview(AGV_JOB_OVERVIEW, aeForm, out sResult);
                            break;
                        case AGV_JOB_OVERVIEW_OPEN_DETAIL:
                            AgvJobOverviewOpenDetail(AGV_JOB_OVERVIEW_OPEN_DETAIL, aeForm, out sResult);
                            break;
                        case AGV_ACTION_MENUITEMS_VALIDATION:
                            AgvActionMenuitemsValidation(AGV_ACTION_MENUITEMS_VALIDATION, aeForm, out sResult);
                            break;
                        case AGV_RESTART:
                            RestartAgv(AGV_RESTART, aeForm, out sResult);
                            break;
                        case AGV_MODE:
                            AgvModeSemiAutomatic(AGV_MODE, aeForm, out sResult);
                            break;
                        case AGV_OVERVIEW_REMOVE_ALL:
                            AgvsAllModeRemoved(AGV_OVERVIEW_REMOVE_ALL, aeForm, out sResult);
                            break;
                        case AGV_OVERVIEW_ID_SORTING:
                            AgvsIdSorting(AGV_OVERVIEW_ID_SORTING, aeForm, out sResult);
                            //MessageBox.Show("AGV_OVERVIEW_ID_SORTING", "Click OK to continue test activescript", MessageBoxButtons.OK);
                            break;
                        case CREATE_NEW_TRANSPORT:
                            CreateNewTransport1(CREATE_NEW_TRANSPORT, aeForm, out sResult);
                            break;
                        case EDIT_TRANSPORT:
                            EditTransport(EDIT_TRANSPORT, aeForm, out sResult);
                            break;
                        case SUSPEND_TRANSPORT:
                            SuspendTransport(SUSPEND_TRANSPORT, aeForm, out sResult);
                            break;
                        case RELEASE_TRANSPORT:
                            ReleaseTransport(RELEASE_TRANSPORT, aeForm, out sResult);
                            break;
                        case CANCEL_TRANSPORT:
                            CancelTransport(CANCEL_TRANSPORT, aeForm, out sResult);
                            break;
                        case TRANSPORT_OVERVIEW_OPEN_DETAIL:
                            TransportOverviewOpenDetail(TRANSPORT_OVERVIEW_OPEN_DETAIL, aeForm, out sResult);
                            break;
                        case SCRIPT_BEFORE_AFTER_ACTIVATE:
                            ScriptBeforeAfterActivate(SCRIPT_BEFORE_AFTER_ACTIVATE, aeForm, out sResult);
                            break;
                        case SCRIPT_BEFORE_AFTER_DEACTIVATE:
                            ScriptBeforeAfterDeactivate(SCRIPT_BEFORE_AFTER_DEACTIVATE, aeForm, out sResult);
                            break;
                        case EPIA4_CLOSE:
                            Epia3Close(EPIA4_CLOSE, aeForm, out sResult);
                            MessageBox.Show("EPIA4_CLOSE", "Click OK to continue", MessageBoxButtons.OK);
                            break;
                        case SYSTEM_OVERVIEW_QUERY:
                            SystemOverviewQuery(SYSTEM_OVERVIEW_QUERY, aeForm, out sResult);
                            break;
                        case ETRICC_EXPLORER_OVERVIEW:
                            EtriccExplorerOverview(ETRICC_EXPLORER_OVERVIEW, aeForm, out sResult);
                            //MessageBox.Show("EtriccExplorerOverview", "Click OK to continue test...", MessageBoxButtons.OK);
                            break;
                        case ETRICC_EXPLORER_LOAD_PROJECT:
                            EtriccExplorerLoadProject(ETRICC_EXPLORER_LOAD_PROJECT, aeForm, out sResult);
                            //MessageBox.Show("EtriccExplorerOverview", "Click OK to continue test...", MessageBoxButtons.OK);
                            break;
                        case ETRICC_EXPLORER_EDIT_SAVE_PROJECT:
                            EtriccExplorerEditSaveProject(ETRICC_EXPLORER_EDIT_SAVE_PROJECT, aeForm, out sResult);
                            //MessageBox.Show("EtriccExplorerOverview", "Click OK to continue test...", MessageBoxButtons.OK);
                            break;
                        case ETRICC_EXPLORER_BUILD_PROJECT:
                            EtriccExplorerBuildProject(ETRICC_EXPLORER_BUILD_PROJECT, aeForm, out sResult);
                            //MessageBox.Show("EtriccExplorerOverview", "Click OK to continue test...", MessageBoxButtons.OK);
                            break;
                        case ETRICC_EXPLORER_CLOSE:
                            EtriccExplorerClose(ETRICC_EXPLORER_CLOSE, aeForm, out sResult);
                            //MessageBox.Show("EtriccExplorerOverview", "Click OK to continue test...", MessageBoxButtons.OK);
                            break;
                        case AGV_ENGINEERING_SIMULATION:
                            AgvEngeeringSimulation(AGV_ENGINEERING_SIMULATION, aeForm, out sResult);
                            break;
                        default:
                            break;
                    }

                    
                    FileManipulation.WriteExcelTestCaseResult(xApp, sResult, sHeaderContentsLength, Counter, sTestCaseName[Counter], sErrorMessage);

                    ++sTotalCounter;
                    if (sResult == ConstCommon.TEST_PASS)
                    {
                        sErrorMessage = string.Empty;
                        ++sTotalPassed;
                    }
                    if (sResult == ConstCommon.TEST_FAIL)
                        ++sTotalFailed;
                    if (sResult == ConstCommon.TEST_EXCEPTION)
                        ++sTotalException;
                    if (sResult == ConstCommon.TEST_UNDEFINED)
                        ++sTotalUntested;

                    ++Counter;
                }

                MessageBox.Show("end", "Click OK to continue sace execel.", MessageBoxButtons.OK);
                FileManipulation.WriteExcelFoot(xApp, sHeaderContentsLength, Counter, sTotalCounter, sTotalPassed, sTotalFailed);
                Epia3Common.WriteTestLogMsg(slogFilePath, "sTotalCounter: " + sTotalCounter, sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "sTotalPassed: " + sTotalPassed, sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "sTotalFailed: " + sTotalFailed, sOnlyUITest);

                #region Save excel file and send email
                // save Excel to Local machine
                string sXLSPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                    sOutFilename + ".xls");

                Epia3Common.WriteTestLogMsg(slogFilePath, "Save to local machine : " + sXLSPath, sOnlyUITest);
                if (FileManipulation.SaveExcel(xApp, sXLSPath, ref sErrorMessage) == false)
                {
                    string sTXTPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), sOutFilename + ".txt");
                    StreamWriter write = File.CreateText(sTXTPath);
                    write.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                    write.Close();
                }

                // Save to remote machine
                if (System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().StartsWith("TEAMSYSTEMS\\JIEMINSHI"))
                {
                    Console.WriteLine("\n   not write to remote machine");
                }
                else    //better use copy local to remote
                {
                    string sXLSPath2 = System.IO.Path.Combine(sOutFilePath, sOutFilename + ".xls");
                    Console.WriteLine("Save2 : " + sXLSPath2);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "sXLSPath2 =: " + sXLSPath2, sOnlyUITest);
                    if (FileManipulation.SaveExcel(xApp, sXLSPath2, ref sErrorMessage) == false)
                    {
                        string sTXTPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), sOutFilename + ".txt");
                        StreamWriter write = File.AppendText(sTXTPath);
                        write.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        write.Close();
                    }
                }

                // quit Excel.
                xApp.Quit();

                // Send Result via Email
                SendEmail(sXLSPath);

                Thread.Sleep(5000);
                ProcessUtilities.CloseProcess("EXCEL");
                TestTools.ProcessUtilities.CloseProcess("EPIA.Explorer");

                #endregion

                if (!sOnlyUITest)
                {
                    string msgX = "update etricc build quality test status to Passed if needed";
                    TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX, ref tfsProjectCollection);

                    while (TFSConnected == false)
                    {
                        TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay);
                        System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX, ref tfsProjectCollection);
                    }

                    if (TFSConnected)
                    {
                        m_BuildSvc = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));
                        // added check sTestResultFolder exist; some time during testing this build can be completely deleted by WVB
                        if (Directory.Exists(sTestResultFolder))
                        {
                            #region  // update testinfo file first and then update build quality
                            string testout = "-->" + sOutFilename + ".xls";
                            if (sAutoTest)
                            {
                                Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI), sBuildNr);

                                string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                                if (sTotalFailed == 0)
                                {
                                    if (TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Passed", "Tests OK", sInfoFileKey) == false)
                                    { 
                                        // build is deleted by Wim, exit this app 
                                        System.Environment.Exit(0);
                                    }
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Passed:" + ConstCommon.ETRICCUI, sOnlyUITest);

                                    Console.WriteLine(" Update build quality:  quality: " + quality);
                                    if (quality.Equals("GUI Tests Failed"))
                                    {
                                        Console.WriteLine("Quality is GUI Tests Failed : No Update Needed ");
                                        Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                                    }
                                    else
                                    {   // check all def tested 
                                        Epia3Common.WriteTestLogMsg(slogFilePath, "Quality is:" + quality + " now check IsAllTestDefinitionsTested", sOnlyUITest);
                                        for (int i = 0; i < mTestDefinitionTypes.Length; i++)
                                        {
                                            Epia3Common.WriteTestLogMsg(slogFilePath, " testdefinition[" + i + "]: " + mTestDefinitionTypes[i], sOnlyUITest);
                                        }

                                        try
                                        {
                                            if (TestListUtilities.IsAllTestDefinitionsTested(mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage) == false)
                                            {
                                                Console.WriteLine("NOT All Test definitions tested " + sErrorMessage);
                                                Epia3Common.WriteTestLogMsg(slogFilePath, "NOT All Test definitions tested " + sErrorMessage, sOnlyUITest);
                                            }
                                            else
                                            {
                                                if (TestListUtilities.IsAllTestStatusPassed(mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage) == true)
                                                {
                                                    // update quality to GUI Tests Passed
                                                    TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI),
                                                    "GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");

                                                    Console.WriteLine("update quality to true -----  ");
                                                    Thread.Sleep(1000);
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Windows.Forms.MessageBox.Show("exception  " + ex.Message + "---" + ex.StackTrace);
                                        }
                                    }
                                }
                                else
                                {
                                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Failed", "--->" + sOutFilename + ".log", sInfoFileKey);
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Failed:" + ConstCommon.ETRICC_UI, sOnlyUITest);

                                    Console.WriteLine(" Update build quality:  quality: " + quality);
                                    if (quality.Equals("GUI Tests Failed"))
                                    {
                                        Console.WriteLine("Quality is GUI Tests Failed : No Update Needed ");
                                        Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                                    }
                                    else
                                    {
                                        // update quality to GUI Tests Passed
                                        TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI),
                                        "GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");

                                        Console.WriteLine("update quality to GUI Tests Failed -----  ");
                                        Thread.Sleep(1000);
                                    }
                                }
                            }
                            #endregion
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Thread.Sleep(2000);
                if (sAutoTest)
                {
                    #region // test exception : update infofile and build quality
                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Exception", " -->" + sOutFilename + ".log", sInfoFileKey);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Exception -->" + sOutFilename + ".log:" + ConstCommon.ETRICCUI, sOnlyUITest);

                    Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace, sOnlyUITest);

                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                    ProcessUtilities.CloseProcess("cmd");
                    ProcessUtilities.CloseProcess("EXCEL");
                    TestTools.ProcessUtilities.CloseProcess("EPIA.Explorer");

                    string msgX = "etricc test exception build quality test status to Failed if needed";
                    TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX, ref tfsProjectCollection);
                    while (TFSConnected == false)
                    {
                        TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay);
                        System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX, ref tfsProjectCollection);
                    }

                    if (TFSConnected)
                    {
                        m_BuildSvc = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));
                        Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI), sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                        if (quality.Equals("GUI Tests Failed"))
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                        }
                        else
                        {
                            TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI),
                                "GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
                        }
                    }
                    #endregion
                }
                return;
            }
            finally
            {
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                             AutomationElement.RootElement,
                            sUIAShellEventHandler);
            }

            try
            {
                #region // Functional Testing
                if (sTotalFailed == 0 && sFunctionalTest == true && sBuildDef.ToLower().StartsWith("nightly"))
                {
                    Console.WriteLine("Start Functional Testing: ");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Start Functional Testing... ", sOnlyUITest);

                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);

                    // unzip project file
                    //string zipFile = @"C:\Testing\EurobalticWorker.zip";
                    string zipFile = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "EurobalticWorker.zip");
                    string workerPathValue = "\"C:\\EtriccTests\\EurobalticWorker.xml\"";
                    if (sProjectFile.ToLower().StartsWith("testproject"))
                    {
                        zipFile = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\AutomaticTesting\TestProjectWorker.zip";
                        workerPathValue = "\"C:\\EtriccTests\\TestProjectWorker.xml\"";
                    }

                    FastZip fz = new FastZip();
                    fz.ExtractZip(zipFile, @"C:\EtriccTests", "");

                    Thread.Sleep(5000);

                    XmlServerConfigUpdate(workerPathValue);
                    Thread.Sleep(10000);

                    #region // Recompile TestRuns
                    //
                    // Recompile TestRuns
                    //

                    Epia3Common.WriteTestLogMsg(slogFilePath, "Start TestRun Recompiled  ", sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Start TestRun Recompiled  ", !sOnlyUITest);
                    string deletePathDll = @"C:\EtriccTests\TestRuns\bin\debug\Egemin*.dll";
                    string msg = "Recompile TestRuns:delete files:" + deletePathDll;
                    if (!FileManipulation.DeleteFilesWithWildcards(deletePathDll, ref msg))
                        throw new Exception(msg);

                    string deletePathPdb = @"C:\EtriccTests\TestRuns\bin\debug\Egemin*.pdb";
                    string msg1 = "Recompile TestRuns:delete files:" + deletePathDll;
                    if (!FileManipulation.DeleteFilesWithWildcards(deletePathPdb, ref msg1))
                        throw new Exception(msg1);

                    Thread.Sleep(3000);

                    string origPath = sEtriccServerRoot + @"\Egemin*.dll";
                    string destPath = @"C:\EtriccTests\TestRuns\bin\debug\";
                    string msg2 = "recompile TestRuns: ";
                    if (!FileManipulation.CopyFilesWithWildcards(origPath, destPath, ref msg2))
                        throw new Exception(msg2);

                    //=========================================
                    // Compile TestRuns
                    //-========================================
                    Thread.Sleep(3000);

                    string space = " ";
                    char Qmark = '"';
                    string dllPath = string.Empty;
                    string arg = string.Empty;

                    dllPath = sEtriccServerRoot + @"\";
                    // DOTNET Version 3.5
                    if (System.IO.Directory.GetCurrentDirectory().Equals(m_CurrentDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))  //  Deployed Testing
                    {
                        arg = "/debug /target:library /out:" + Qmark + ConstCommon.ETRICC_TESTS_DIRECTORY + @"\TestRuns\bin\debug\TestRuns.dll" + Qmark;
                        arg = arg + space + Qmark + m_CurrentDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestWorker.cs" + Qmark
                            //+ space + getRootPath() + @"TestRuns\TestOpstellingWorker.cs"
                             + space + Qmark + m_CurrentDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestConstants.cs" + Qmark
                             + space + Qmark + m_CurrentDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\Logger.cs" + Qmark
                             + space + Qmark + m_CurrentDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestData.cs" + Qmark
                             + space + Qmark + m_CurrentDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestUtility.cs" + Qmark;
                        arg = arg + space + "/reference:";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Design.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Interfaces.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Security.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Security.SSPI.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.UI.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Definitions.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.WCS.dll" + '"' + ";";
                        //arg = arg + '"' + System.IO.Path.Combine(m_CurrentDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "Interop.Excel.dll") + '"' + ";";
                        arg = arg + '"' + System.IO.Path.Combine(m_CurrentDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "Microsoft.Office.Interop.Excel.dll") + '"' + ";";

                        if (!File.Exists(System.IO.Path.Combine(destPath, "Interop.Excel.dll")))
                            File.Copy(System.IO.Path.Combine(m_CurrentDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "Microsoft.Office.Interop.Excel.dll"),
                                System.IO.Path.Combine(destPath, "Interop.Excel.dll"), true);

                        //if (!File.Exists(System.IO.Path.Combine(destPath, "Interop.VBIDE.dll")))
                        //    File.Copy(System.IO.Path.Combine(m_CurrentDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "Interop.VBIDE.dll"),
                        //        System.IO.Path.Combine(destPath, "Interop.VBIDE.dll"), true);
                    }
                    else    //  test in development environment 
                    {
                        /*arg = "/debug /target:library /out:" + Qmark + Constants.ETRICC_TESTS_DIRECTORY + @"\TestRuns\bin\debug\TestRuns.dll" + Qmark;
                        arg = arg + space + Qmark + getRootPath() + @"TestRuns\TestWorker.cs" + Qmark
                            //+ space + getRootPath() + @"TestRuns\TestOpstellingWorker.cs"
                             + space + Qmark + getRootPath() + @"TestRuns\TestConstants.cs" + Qmark
                             + space + Qmark + getRootPath() + @"TestRuns\Logger.cs" + Qmark
                             + space + Qmark + getRootPath() + @"TestRuns\TestData.cs" + Qmark
                             + space + Qmark + getRootPath() + @"TestRuns\TestUtility.cs" + Qmark;
                        arg = arg + space + "/reference:";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Design.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Interfaces.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Security.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Security.SSPI.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Common.UI.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.Definitions.dll" + '"' + ";";
                        arg = arg + '"' + dllPath + "Egemin.EPIA.WCS.dll" + '"' + ";";
                        arg = arg + '"' + @"C:\Epia 3\Testing\Automatic\AutomaticTests\Source\TestRuns\bin\Debug\Interop.Excel.dll" + '"' + ";";
                        arg = arg + '"' + @"C:\Epia 3\Testing\Automatic\AutomaticTests\Source\TestRuns\bin\Debug\Interop.VBIDE.dll" + '"' + ";";

                        if (!File.Exists(Path.Combine(destPath, "Interop.Excel.dll")))
                            File.Copy(@"C:\Epia 3\Testing\Automatic\AutomaticTests\OEM\Microsoft\Interop.Excel.dll",
                                Path.Combine(destPath, "Interop.Excel.dll"), true);

                        if (!File.Exists(Path.Combine(destPath, "Interop.VBIDE.dll")))
                            File.Copy(@"C:\Epia 3\Testing\Automatic\AutomaticTests\OEM\Microsoft\Interop.VBIDE.dll",
                                Path.Combine(destPath, "Interop.VBIDE.dll"), true);
                         */

                        Thread.Sleep(1000);
                    }

                    Epia3Common.WriteTestLogMsg(slogFilePath, "arg: " + arg, sOnlyUITest);

                    //string DotnetVersionPath = m_CurrentDrive + @"WINDOWS\Microsoft.NET\Framework\v2.0.50727";
                    string DotnetVersionPath = m_CurrentDrive + @"WINDOWS\Microsoft.NET\Framework\v3.5";
                    string exePath = System.IO.Path.Combine(DotnetVersionPath, "csc.exe");
                    // Run recompile Process
                    string output = ProcessUtilities.RunProcessAndGetOutput(exePath, arg);
                    if (output.IndexOf("error") >= 0)
                        throw new Exception(output);

                    Epia3Common.WriteTestLogMsg(slogFilePath, "recompile exePath: " + exePath, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "recompile arg: " + arg, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "recompile output: " + output, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "TestRun Recompiled ", sOnlyUITest);

                    Epia3Common.WriteTestLogMsg(slogFilePath, "!recompile exePath: " + exePath, !sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "!recompile arg: " + arg, !sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "!recompile output: " + output, !sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "!TestRun Recompiled ", !sOnlyUITest);

                    Thread.Sleep(2000);
                    #endregion

                    // Start Worker
                    Console.WriteLine("Start Worker: ");

                    Epia3Common.WriteTestLogMsg(slogFilePath, "Start Worker... ", sOnlyUITest);

                    //----------------------
                    try
                    {
                        ProcessUtilities.CloseProcess("EXCEL");
                        //========================   SERVER =================================================
                        #region SERVER
                        if (sServerRunAs.ToLower().IndexOf("service") >= 0)
                        {
                            if (!sOnlyUITest)
                            {
                                if (TFSConnected)
                                {
                                    try
                                    {
                                        Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI), sBuildNr);
                                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                                        //if (quality.Equals("GUI Tests Failed"))
                                        //{
                                        //    Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                                        //}
                                        //else
                                        //{
                                        TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI),
                                            "Functional Tests Started", m_BuildSvc, sDemo ? "true" : "false");
                                        //}
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message + "---" + ex.StackTrace, "TFSConnected Exception");
                                        Console.WriteLine(ex.Message);
                                        Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "-final check -" + ex.StackTrace, sOnlyUITest);
                                    }
                                }
                            }

                            // uninstall Egemin.Epia.server Service
                            Console.WriteLine("UNINSTALL EPIA SERVER Service : ");
                            TestTools.ProcessUtilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
                                ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /u");
                            Thread.Sleep(2000);

                            // uninstall Egemin.Etricc.server Service
                            Console.WriteLine("UNINSTALL ETRICC SERVER Service : ");
                            TestTools.ProcessUtilities.StartProcessWaitForExit(sEtriccServerRoot,
                                ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /u");
                            Thread.Sleep(2000);

                            Console.WriteLine("INSTALL EPIA SERVER Service : ");
                            TestTools.ProcessUtilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
                                ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /i");
                            Thread.Sleep(2000);

                            Console.WriteLine("INSTALL ETRICC SERVER Service : ");
                            TestTools.ProcessUtilities.StartProcessWaitForExit(sEtriccServerRoot,
                                ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /i");
                            Thread.Sleep(2000);

                            Console.WriteLine("Start EPIA SERVER as Service : ");
                            TestTools.ProcessUtilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
                                ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /start");
                            Thread.Sleep(2000);

                            Console.WriteLine("Start ETRICC SERVER as Service : ");
                            TestTools.ProcessUtilities.StartProcessWaitForExit(sEtriccServerRoot,
                                ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /start");
                            Thread.Sleep(2000);

                            ServiceController svcEpia = new ServiceController("Egemin Epia Server");
                            Console.WriteLine(svcEpia.ServiceName + " has status " + svcEpia.Status.ToString());
                            Thread.Sleep(2000);
                            svcEpia.WaitForStatus(ServiceControllerStatus.Running);

                            if (!svcEpia.Status.ToString().ToLower().StartsWith("running"))
                                throw new Exception("Egemin Epia Server Service not running:"
                                    + svcEpia.Status.ToString().ToLower());

                            ServiceController svcEtricc = new ServiceController("Egemin Etricc Server");
                            Console.WriteLine(svcEtricc.ServiceName + " has status " + svcEtricc.Status.ToString());
                            Thread.Sleep(2000);
                            svcEtricc.WaitForStatus(ServiceControllerStatus.Running);

                            if (!svcEtricc.Status.ToString().ToLower().StartsWith("running"))
                                throw new Exception("Egemin Etricc Server Service not running"
                                    + svcEtricc.Status.ToString().ToLower());
                            /*while (svcEtricc.Status != ServiceControllerStatus.Running)
                            {
                                Console.WriteLine(svcEtricc.ServiceName + " ==has status == " + svcEtricc.Status.ToString());
                                Thread.Sleep(2000);
                            }*/

                            Console.WriteLine("EPIA and ETRICC SERVER Service Started : ");
                            Thread.Sleep(2000);
                        }

                        if (sServerRunAs.ToLower().IndexOf("console") >= 0)
                        {
                            // Start Epia SERVER as Console
                            TestTools.ProcessUtilities.StartProcessNoWait(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
                                ConstCommon.EGEMIN_EPIA_SERVER_EXE, string.Empty);

                            // Start Etricc SERVER as Console
                            TestTools.ProcessUtilities.StartProcessNoWait(sEtriccServerRoot, ConstCommon.EGEMIN_ETRICC_SERVER_EXE, string.Empty);
                            Thread.Sleep(90000);
                        }
                        #endregion
                        Console.WriteLine("----- time 90 seconds .....");
                        Thread.Sleep(15000);

                        sStartTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - sStartTime;

                        //========================   SHELL =================================================
                        #region  Shell
                        sUIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
                        // Add Open window Event Handler
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement, TreeScope.Descendants, sUIAShellEventHandler);
                        sEventEnd = false;
                        TestCheck = ConstCommon.TEST_PASS;

                        Thread.Sleep(45000);

                        // Start Shell
                        TestTools.ProcessUtilities.StartProcessNoWait(m_SystemDrive + ConstCommon.EPIA_CLIENT_ROOT,
                            ConstCommon.EGEMIN_EPIA_SHELL_EXE, string.Empty);
                        //--------------------------
                        sStartTime = DateTime.Now;
                        sTime = DateTime.Now - sStartTime;
                        int wt = 0;
                        Console.WriteLine(" time is :" + sTime.TotalSeconds);
                        while (sEventEnd == false && wt < 60)
                        {
                            Thread.Sleep(2000);
                            //sTime = DateTime.Now - sStartTime;
                            wt = wt + 2;
                            Console.WriteLine("wait shell start up time is (sec) : " + wt);
                        }

                        Console.WriteLine("Shell started after (sec) : " + 2 * wt);
                       
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
                        MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Start Server and Shell");
                    }
                    //------------------------
                    System.Diagnostics.Process proc = null;
                    int pIDx = ProcessUtilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out proc);
                    Console.WriteLine("Proc ID:" + pIDx);

                    Thread MyNewThread = new Thread(new ThreadStart(ClickScreenThreadProc));
                    MyNewThread.Start();

                    //-------------------------------- Check Functional Testing End?
                    string directoryPath = ConstCommon.ETRICC_TESTS_DIRECTORY;
                    string fileNameExcel = "*.xls";
                    string[] allFilesExcel = System.IO.Directory.GetFiles(directoryPath, fileNameExcel);
                    while (allFilesExcel.Length == 0)
                    {
                        Thread.Sleep(30000);
                        Console.WriteLine("No xls files found   :" + System.DateTime.Now);
                        allFilesExcel = System.IO.Directory.GetFiles(directoryPath, fileNameExcel);
                    }

                    Console.WriteLine("xls file is:" + allFilesExcel[0]);
                    MyNewThread.Abort();

                    Epia3Common.WriteTestLogMsg(slogFilePath, "xls file is:" + allFilesExcel[0], sOnlyUITest);


                    //
                    // copy excel file to testresult folder 
                    //  Debug-20081110-17-45-17-GUITESTS-EPIATESTPC1.xls  become
                    //  Debug-20081110-17-45-17-FUNCTESTS-EPIATESTPC1
                    int lastIndex = sOutFilename.LastIndexOf('-');
                    string prefix = sOutFilename.Substring(0, lastIndex);
                    lastIndex = prefix.LastIndexOf('-');
                    prefix = prefix.Substring(0, lastIndex);
                    string resultFile = prefix + "-FUNCTESTS-" + System.Environment.MachineName;
                    Epia3Common.WriteTestLogMsg(slogFilePath, "resultFile is:" + resultFile, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "resultFullFile is:"
                        + System.IO.Path.Combine(sOutFilePath, resultFile + ".xls"), sOnlyUITest);


                    FileInfo fileExcel = new FileInfo(allFilesExcel[0]);
                    fileExcel.CopyTo(System.IO.Path.Combine(sOutFilePath, resultFile + ".xls"));

                    // copy txt file to testresult folder 
                    string fileNameTxt = "TestEur*.txt";
                    string[] allFilesTxt = System.IO.Directory.GetFiles(directoryPath, fileNameTxt);

                    FileInfo fileTxt = new FileInfo(allFilesTxt[0]);
                    fileTxt.CopyTo(System.IO.Path.Combine(sOutFilePath, resultFile + ".log"));

                    Epia3Common.WriteTestLogMsg(slogFilePath, "Copy Finished:", sOnlyUITest);

                    // Check Functional Test Result
                    string Path = System.IO.Path.Combine(sOutFilePath, resultFile + ".xls");
                    // initialize the Excel Application class
                    //Excel.ApplicationClass app = new Excel.ApplicationClass();
                    xAppFunc = new Excel.Application();
                    // create the workbook object by opening the excel file.
                    Epia3Common.WriteTestLogMsg(slogFilePath, "check test result by open Result File" + Path, sOnlyUITest);

                    try
                    {
                        Excel.Workbook workBook = xAppFunc.Workbooks.Open(Path,
                                                                 0,
                                                                 true,
                                                                 5,
                                                                 "",
                                                                 "",
                                                                 true,
                                                                 Excel.XlPlatform.xlWindows,
                                                                 "\t",
                                                                 false,
                                                                 false,
                                                                 0,
                                                                 true,
                                                                 1,
                                                                 0);
                        // get the active worksheet using sheet name or active sheet
                        Excel.Worksheet workSheet = (Excel.Worksheet)workBook.ActiveSheet;
                        // This row,column index should be changed as per your need.
                        // i.e. which cell in the excel you are interesting to read.


                        object rowIndex = 86;
                        //object rowIndex = 4;
                        object colIndex2 = 2;

                        sFuncTotalFailed = ((Excel.Range)workSheet.Cells[rowIndex, colIndex2]).Value2.ToString();
                        Console.WriteLine("cell value is:" + sFuncTotalFailed);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "---Failed Func Tests Count is:" + sFuncTotalFailed, sOnlyUITest);
                    }
                    catch (Exception ex)
                    {
                        sFuncTotalFailed = "1";
                        xAppFunc.Quit();
                        Console.WriteLine(ex.Message);
                        Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "-final check -" + ex.StackTrace, sOnlyUITest);
                    }

                    xAppFunc.Quit();

                    if (!sOnlyUITest)
                    {
                        if (TFSConnected)
                        {
                            try
                            {
                                Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI), sBuildNr);
                                Epia3Common.WriteTestLogMsg(slogFilePath, "uri -" + uri.ToString(), sOnlyUITest);

                                string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;
                                Epia3Common.WriteTestLogMsg(slogFilePath, "quality -" + quality, sOnlyUITest);

                                //if (quality.Equals("GUI Tests Failed"))
                                //{
                                //    Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                                //}
                                //else
                                //{
                                if (sFuncTotalFailed.Equals("0"))
                                    TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI),
                                        "Functional Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
                                else
                                    TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI),
                                        "Functional Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
                                //}


                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                                Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "-final check -" + ex.StackTrace, sOnlyUITest);
                            }

                        }
                    }

                }  //end sTotalFailed == 0 || sFunctionalTest == true
                #endregion

                #region Close LogFile
                Epia3Common.CloseTestLog(slogFilePath, sOnlyUITest);

                Console.WriteLine("\nClosing application in 10 seconds");
                if (sOnlyUITest)
                    Thread.Sleep(10000000);
                else
                    Thread.Sleep(10000);

                // close CommandHost
                Thread.Sleep(10000);
                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);
                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                Console.WriteLine("\nEnd test run\n");

                if (sAutoTest)
                {
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Working file updated now, Functional testing finished ", sOnlyUITest);
                    //FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                }
                #endregion
            }
            catch (Exception ex)
            {
                #region Functional Testing Exception
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                if (sAutoTest)
                {
                    if (sParentProgram.StartsWith("TFS"))
                    {
                        TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "FUNCTIONAL Tests Exception", " -->" + sOutFilename + ".log", sInfoFileKey);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.ETRICCUI, sOnlyUITest);
                    }
                    else
                    {
                        TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "FUNCTIONAL Tests Exception", " -->" + sOutFilename + ".log", sInfoFileKey);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.ETRICC_UI, sOnlyUITest);
                    }


                    Epia3Common.WriteTestLogFail(slogFilePath, "FUNCTIONAL Tests Exception -->" + sOutFilename + ".log:" + ConstCommon.ETRICC_UI, sOnlyUITest);

                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                    ProcessUtilities.CloseProcess("cmd");
                    FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");

                    if (TFSConnected)
                    {
                        Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI), sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;
                        if (quality.Equals("Functional Tests Failed"))
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                        }
                        else
                        {
                            TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(TFSQATestTools.Constants.ETRICCUI),
                                "Functional Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
                        }
                    }
                }
                #endregion
            }
        }

        /*private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;

            }
            finally
            {
                GC.Collect();
            }
        }*/

        public static void ClickUiScreenActionToAvoidScreenStandBy()
        {
            System.Windows.Point point = new Point(1, 1);

            Input.MoveToAndRightClick(point);

            Thread.Sleep(2000);

            point.Y = point.Y + 300;
            Input.MoveToAndClick(point);

            Thread.Sleep(2000);
        }

        protected static void ClickScreenThreadProc()
        {
            while (true)
            {
                ClickUiScreenActionToAvoidScreenStandBy();
                Thread.Sleep(60000);
            }

        }

        /// <summary>
        /// Cancel a drive mapping to the destination
        /// </summary>
        /// <param name="Destination">Full drive path</param>
        public static int CancelDriveMap(string driveLetter)
        {
            int result = WNetCancelConnection2A(driveLetter, 1, 1);
            return result;
        }

        private static void XmlServerConfigUpdate(string WorkerPathValue)
        {
            var xDoc = new XmlDocument();
            //xDoc.Load("C:\\Etricc\\Server\\Egemin.Etricc.Server.exe.config");
            xDoc.Load(Path.Combine(sEtriccServerRoot, "Egemin.Etricc.Server.exe.config"));

            var xPathNav = xDoc.CreateNavigator();
            xPathNav.MoveToFirstChild();
            xPathNav.MoveToFirstChild();
            while (!xPathNav.LocalName.StartsWith("epia.componentconfiguration"))
            {
                xPathNav.MoveToNext();
            }

            xPathNav.MoveToFirstChild();
            while (!xPathNav.LocalName.Equals("parameter"))
            {
                xPathNav.MoveToFirstChild();
            }

            if (xPathNav.GetAttribute("name", "").ToLower().StartsWith("xmlfile"))
            {
                xPathNav.ReplaceSelf("<parameter name=\"XmlFile\" value=" + WorkerPathValue + " />");
            }
            xDoc.Save(Path.Combine(sEtriccServerRoot, "Egemin.Etricc.Server.exe.config"));

            return;
        }
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region StartEpiaEtriccServerShell
        public static int StartEpiaEtriccServerShell(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            try
            {
                if (sEtriccServerStartupOK == false)
                {
                    sErrorMessage = "StartEpiaEtriccServerShell failed:" + sErrorMessage;
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                //========================   SERVER =================================================
                AutomationEventHandler UIAServerEventHandler = new AutomationEventHandler(OnUIAServerEvent);
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                  AutomationElement.RootElement, TreeScope.Descendants, UIAServerEventHandler);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (ProjServerOrShellStartup.ServerStartup("Epia Server", sServerRunAs, ref sErrorMessage, slogFilePath, sOnlyUITest) == false)
                    {
                        sEtriccServerStartupOK = false;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }

                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        if (ProjServerOrShellStartup.ServerStartup("Etricc Server", sServerRunAs, ref sErrorMessage, slogFilePath, sOnlyUITest) == false)
                        {
                            sEtriccServerStartupOK = false;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }
               
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                          AutomationElement.RootElement, UIAServerEventHandler);

                //========================   SHELL =================================================
                sUIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
                sStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - sStartTime;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(5000);
                    Console.WriteLine("EPIA ETRICC SERVER Service Started : ");
                    // Add Open window Event Handler
                    Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement, TreeScope.Descendants, sUIAShellEventHandler);
                    sEventEnd = false;
                    #region  Shell
                    Console.WriteLine(" --------------------Start Epia Shell ...................... :" + m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() +
                        @"\Dematic\Epia Shell");
                    TestTools.ProcessUtilities.StartProcessNoWait(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() +
                        @"\Dematic\Epia Shell",
                        ConstCommon.EGEMIN_EPIA_SHELL_EXE, string.Empty);
                    //--------------------------
                    sStartTime = DateTime.Now;
                    sTime = DateTime.Now - sStartTime;
                    Console.WriteLine(" time is :" + sTime.TotalSeconds);
                    while (sEventEnd == false && sTime.TotalSeconds <120)
                    {
                        Thread.Sleep(2000);
                        sTime = DateTime.Now - sStartTime;
                        Console.WriteLine("wait shell start up time is (sec) : " + sTime.TotalSeconds);
                        Console.WriteLine("TestCkeck : " + TestCheck.ToString());
                    }

                    Console.WriteLine("Shell started after (sec) : " + sTime.TotalSeconds);
                    //Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    //       AutomationElement.RootElement,
                    //      UIAShellEventHandler);

                    Console.WriteLine("TestCkeck : " + TestCheck.ToString());
                    Thread.Sleep(10000);
                    if (TestCheck == ConstCommon.TEST_FAIL)
                    {
                        Thread.Sleep(10000);
                        throw new Exception("shell start up failed:" + sErrorMessage);
                    }
                    #endregion
                }
                Console.WriteLine("Shell started after (sec) : " + mTime.Seconds);


                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Application is started : ");
                    aeForm = EtriccUtilities.GetMainWindow("MainForm", 120);
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
                    {
                        Console.WriteLine("Application maeForm name : " + aeForm.Current.Name);
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    sEtriccServerStartupOK = false;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine(testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "  ---  " + ex.StackTrace;
                Console.WriteLine(testname + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
                sEtriccServerStartupOK = false;
                //System.Windows.Forms.MessageBox.Show("where ");
                return ConstCommon.TEST_FAIL;
            }
            finally
            {
                Thread.Sleep(3000);
                //Input.SendKeyboardInput(System.Windows.Input.Key.Space, true);
            }

            return TestCheck;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ShellCloseWithinOneMinuteAfterServerDown
        public static void ShellCloseWithinOneMinuteAfterServerDown(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            try
            {
                /*if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 9, 26), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }*/

                StartEpiaEtriccServerShell(START_EPIA_ETRICC_SERVER_SHELL, root, out result);

                if (sUIAShellEventHandler != null)
                {
                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                             AutomationElement.RootElement, sUIAShellEventHandler);
                }

                if (sEtriccServerStartupOK == false)
                {
                    sErrorMessage = "StartEpiaEtriccServerShell failed:" + sErrorMessage;
                    TestCheck = ConstCommon.TEST_FAIL;
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    //if (sOnlyUITest)
                    root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                    if (root != null)
                    {
                        #region open system overview
                        AutomationElement aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, SYSTEM_OVERVIEW, ref sErrorMessage);
                        if (aePanelLink == null)
                        {
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Input.MoveToAndClick(aePanelLink);
                            Thread.Sleep(10000);
                        }
                        #endregion
                    }
                    else
                    {
                        sErrorMessage = "Main Form notfound---";
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }

                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        #region // stop Server
                        Console.WriteLine("Stop EPIA SERVER as Service : ");
                        ProcessUtilities.StartProcessWaitForExit(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() +
                            @"\Dematic\Epia Server",
                            ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /stop");
                        Thread.Sleep(2000);

                        ServiceController svcEpia = new ServiceController("Egemin Epia Server");
                        Console.WriteLine(svcEpia.ServiceName + " has status " + svcEpia.Status.ToString());
                        // wait until epia service status is stopped
                        sStartTime = DateTime.Now;
                        sTime = DateTime.Now - sStartTime;
                        int wat = 0;
                        string epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                        Console.WriteLine(" time is :" + sTime.TotalSeconds);
                        while (!epiaServiceStatus.StartsWith("stopped") && wat < 60)
                        {
                            Thread.Sleep(2000);
                            wat = wat + 2;
                            epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                            Console.WriteLine("wait epia service status stopped:  time is (sec) : " + wat + "  and status is:" + epiaServiceStatus);
                        }

                        //svcEpia.WaitForStatus(ServiceControllerStatus.Running);
                        if (svcEpia.Status != ServiceControllerStatus.Stopped)
                        {
                            sErrorMessage = "Epia Service stop failed:";
                            Epia3Common.WriteTestLogMsg(slogFilePath, "Epia Service stop failed: " + epiaServiceStatus, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        #endregion
                    }
                }
                else
                {
                    sErrorMessage = "Epia Server and Shell startup failed";
                    Console.WriteLine(testname + sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                // wait until dialog box 
                AutomationElement aeWindow = null;
                AutomationElement aeLicenseServiceShutdownDialogBox = null;
                AutomationElement aeShellShutdownButton = null;
                System.Windows.Automation.Condition cWindowLicenseShutdown = new AndCondition(
                    //new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen")
                    );

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region close LicenseServiceShutdownDialogBox
                    sStartTime = DateTime.Now;
                    sTime = DateTime.Now - sStartTime;
                    Console.WriteLine(" time is :" + sTime.TotalSeconds);
                    while (aeLicenseServiceShutdownDialogBox == null && sTime.TotalSeconds < 300)
                    {
                        aeWindow = EtriccUtilities.GetMainWindow("MainForm");
                        if (aeWindow == null)
                        {
                            sErrorMessage = "MainForm is not found";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                            break;
                        }
                        else
                        {
                            //我想Microshaoft的意思应该是
/*var ptr = IntPtr.Zero;
try
{
ptr = ...
}
finally
{
Marshal.ReleaseCom(ptr);
}
*/
//ComException不是由于垃圾回收或者内存的问题,ComException实际上是针对几乎所有的COM错误码的一个封装,你这个问题,应该是UI Automation返回了错误:

// The window style or class attribute is invalid for this operation.
//ERROR_INVALID_WINDOW_STYLE 


                            Console.WriteLine(" wait until dialog box :" + sTime.TotalSeconds);
                            Epia3Common.WriteTestLogMsg(slogFilePath, " wait until dialog box :" + sTime.TotalSeconds, sOnlyUITest);

                            WindowPattern windowPattern = aeWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                            // bool inputIdel = false;
                            while (windowPattern.WaitForInputIdle(100) == false)
                            {
                                Epia3Common.WriteTestLogMsg(slogFilePath, " windowPattern.WaitForInputIdle(100) == false:", sOnlyUITest);
                                Console.WriteLine("windowPattern.WaitForInputIdle(100):" + windowPattern.WaitForInputIdle(100));
                                Thread.Sleep(1000);
                            }

                            int ikk = 0;
                            while (aeLicenseServiceShutdownDialogBox == null && ikk++ < 10)
                            {
                                try
                                {
                                    aeLicenseServiceShutdownDialogBox = aeWindow.FindFirst(TreeScope.Descendants, cWindowLicenseShutdown);
                                }
                                catch (Exception)
                                {
                                    aeLicenseServiceShutdownDialogBox = null;
                                    Thread.Sleep(60000);
                                    Console.WriteLine("----  exception "+ikk);
                                }
                            }
                            Thread.Sleep(2000);
                            sTime = DateTime.Now - sStartTime;
                            Console.WriteLine("wait aeLicenseServiceShutdownDialogBox displayed time is (sec) : " + sTime.TotalSeconds);
                        }
                    }

                    if (aeLicenseServiceShutdownDialogBox == null)
                    {
                        sErrorMessage = "aeLicenseServiceShutdownDialogBox not displayed after 2 min";
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        Console.WriteLine(testname + sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Epia3Common.WriteTestLogMsg(slogFilePath, "aeLicenseServiceShutdownDialogBox IS displayed IN 2 min", sOnlyUITest);
                        System.Windows.Automation.Condition c = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Shell shutdown"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                        );

                        // Find the BUTTON element.
                        aeShellShutdownButton = aeLicenseServiceShutdownDialogBox.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                        if (aeShellShutdownButton != null)
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, "aeShellShutdownButton != null", sOnlyUITest);
                            Point pt = AUIUtilities.GetElementCenterPoint(aeShellShutdownButton);
                            Input.MoveTo(pt);
                            Thread.Sleep(5000);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "ClickAtPoint aeShellShutdownButton", sOnlyUITest);
                            Input.ClickAtPoint(pt);
                        }
                        else
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, "aeShellShutdownButton == null", sOnlyUITest);
                        }
                    }
                    #endregion
                }

                // After two minute shell should be closed
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sStartTime = DateTime.Now;
                    sTime = DateTime.Now - sStartTime;
                    Console.WriteLine(" time is :" + sTime.TotalSeconds);
                    aeWindow = EtriccUtilities.GetMainWindow("MainForm", 5);
                    while (aeWindow != null && sTime.TotalSeconds < 121)
                    {
                        aeWindow = EtriccUtilities.GetMainWindow("MainForm", 5);
                        if (aeWindow == null)
                        {
                            Console.WriteLine("OK MainForm is shutdown after click  button");
                            break;
                        }
                        else
                        {
                            Thread.Sleep(6000);
                            sTime = DateTime.Now - sStartTime;
                            Console.WriteLine(" time is :" + sTime.TotalSeconds);
                        }
                    }

                    //after two minute
                    if (aeWindow != null)
                    {
                        // not wait any longer, kill the shell process
                        ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                        //sErrorMessage = "Shell is still open after two minute";
                        //Console.WriteLine(sErrorMessage);
                        //TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);

                // wait 2 min and try to close dialog screen if exist
                Thread.Sleep(300000);
                // if dialog window still open,should close it, following test cases can continue
                #region close LicenseServiceShutdownDialogBox
                AutomationElement aeWindow = null;
                AutomationElement aeLicenseServiceShutdownDialogBox = null;
                System.Windows.Automation.Condition cWindowLicenseShutdown = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)//,
                    //new PropertyCondition(AutomationElement.AutomationIdProperty, "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen")
                   );
               
                aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm",10);
                if (aeWindow != null)
                {
                    Epia3Common.WriteTestLogMsg(slogFilePath, "MainForm exist , check  dialog box is still open : ---", sOnlyUITest);
                    Console.WriteLine(" check  dialog box is still open :");
                    aeLicenseServiceShutdownDialogBox = aeWindow.FindFirst(TreeScope.Descendants, cWindowLicenseShutdown);
                    Thread.Sleep(2000);
                    if (aeLicenseServiceShutdownDialogBox != null)
                    {
                        Epia3Common.WriteTestLogMsg(slogFilePath, "dialog box exist ---", sOnlyUITest);
                        System.Windows.Automation.Condition cButton = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Shell shutdown"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                            );

                        // Find the BUTTON element.
                        AutomationElement aeShellShutdownButton = aeLicenseServiceShutdownDialogBox.FindFirst(TreeScope.Element | TreeScope.Descendants, cButton);
                        if (aeShellShutdownButton != null)
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, "Shell shutdown button exist ---", sOnlyUITest);
                            Point pt = AUIUtilities.GetElementCenterPoint(aeShellShutdownButton);
                            Input.MoveTo(pt);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(pt);
                            Thread.Sleep(3000);
                        }
                        else
                            Epia3Common.WriteTestLogMsg(slogFilePath, "Shell shutdown button NOT exist ---", sOnlyUITest);
                    }
                    else
                        Epia3Common.WriteTestLogMsg(slogFilePath, "dialog box NOT exist ---", sOnlyUITest);
                }
                else
                    Epia3Common.WriteTestLogMsg(slogFilePath, "MainForm NOT exist ---", sOnlyUITest);

                #endregion
            }
            finally
            {
                // Add Open window Event Handler
                //Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                //  AutomationElement.RootElement, TreeScope.Descendants, sUIAShellEventHandler);
            }
        }
        #endregion 
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ResourceIdIntegrityCheck
        public static void ResourceIdIntegrityCheck(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 8, 14), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                string shellServiceLogFile = Path.Combine(OSVersionInfoClass.ProgramFilesx86() + @"\Dematic\Epia Server\Log", "ShellServices.log");
                string shellServiceDestLogFile = Path.Combine(OSVersionInfoClass.ProgramFilesx86() + @"\Dematic\Epia Server\Log", "ShellServicesDest.log");

                Console.WriteLine("shellServiceLogFile:  " + shellServiceLogFile);
                if (File.Exists(shellServiceLogFile))
                {
                    System.IO.File.Copy(shellServiceLogFile, shellServiceDestLogFile, true);
                    string[] loglines = System.IO.File.ReadAllLines(shellServiceDestLogFile);

                    for (int i = 0; i < loglines.Length; i++)
                    {
                        //Console.WriteLine("loglines[i]:  " + loglines[i]);
                        if (loglines[i].IndexOf("Warning") > 0 && loglines[i].IndexOf("ResourceId") > 0 && loglines[i].IndexOf("Etricc") > 0)
                        {
                            sErrorMessage = "-------- resource file warning log  --> " + loglines[i];
                            TestCheck = ConstCommon.TEST_FAIL;
                            Epia3Common.WriteTestLogMsg(slogFilePath, testname+ " error message: " + sErrorMessage, sOnlyUITest);
                            //Console.WriteLine(sErrorMessage);
                            break;
                        }
                    }
                }

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
                sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
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
        #region OnInstallEtricc5UIEvent
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region PlaybackInstallStartup
        public static void PlaybackInstallStartup(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sErrorMessage = string.Empty;

            try
            {
                // uninstall playback if already installed:
                if (ProjAppInstall.UninstallApplication(EgeminApplication.ETRICC_PLAYBACK, ref sErrorMessage) == false)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = EgeminApplication.ETRICC_PLAYBACK + " Uninstall failed:" + sErrorMessage;
                    Console.WriteLine(sErrorMessage);
                }
                //AutomationEventHandler UIALayoutXPosEventHandler = new AutomationEventHandler(OnInstallEtricc5UIEvent);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    //string path = System.IO.Directory.GetCurrentDirectory() + @"\Setup\Etricc\Current";
                    string path = OSVersionInfoClass.ProgramFilesx86() + @"\Egemin\AutomaticTesting\Setup\Etricc\Current";
                    if ( logger != null) logger.LogMessageToFile(" *** start Etricc installing :" + EgeminApplication.ETRICC_PLAYBACK, 0, 0);
                    if (ProjAppInstall.InstallApplication(path, EgeminApplication.ETRICC_PLAYBACK, EgeminApplication.SetupType.Default, ref sErrorMessage, null/*logger*/))
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        sErrorMessage = EgeminApplication.ETRICC_PLAYBACK + " Uninstall failed:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }         
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\nInstall Etricc Playback.: Pass");
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine(sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(8000);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region HostTestInstallStartup
        public static void HostTestInstallStartup(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sErrorMessage = string.Empty;

            try
            {
                // uninstall hosttest if already installed:
                if (ProjAppInstall.UninstallApplication(EgeminApplication.ETRICC_HOSTTEST, ref sErrorMessage) == false)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = EgeminApplication.ETRICC_HOSTTEST+ " Uninstall failed:" + sErrorMessage;
                    Console.WriteLine(sErrorMessage);
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    string path = OSVersionInfoClass.ProgramFilesx86() + @"\Dematic\AutomaticTesting\Setup\Etricc\Current";
                    if (logger != null)  logger.LogMessageToFile(" *** start Etricc installing :" + EgeminApplication.ETRICC_HOSTTEST, 0, 0);
                    if (ProjAppInstall.InstallApplication(path, EgeminApplication.ETRICC_HOSTTEST, EgeminApplication.SetupType.Default, ref sErrorMessage, null/*logger*/))
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        sErrorMessage = EgeminApplication.ETRICC_HOSTTEST + " Uninstall failed:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                Thread.Sleep(5000);
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\nInstall Etricc HostTest: Pass");
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);

                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
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
        #region HostTestInstallStartupXXX
        public static void HostTestInstallStartupXXX(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            try
            {
                // uninstall hostetst if already installed:
                if (ProjAppInstall.UninstallApplication(EgeminApplication.ETRICC_HOSTTEST, ref sErrorMessage) == false)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = EgeminApplication.ETRICC_HOSTTEST + " Uninstall failed:" + sErrorMessage;
                    Console.WriteLine(sErrorMessage);
                }

                //AutomationEventHandler UIALayoutXPosEventHandler = new AutomationEventHandler(OnInstallEtricc5UIEvent);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Add Open MyLayoutScreen window Event Handler
                    //Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    //    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutXPosEventHandler);

                    System.Threading.Thread.Sleep(15000);
                    string InstallerSource = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory() + @"\Setup\Etricc\Current", "Etricc HostTest.msi");
                    Console.WriteLine("start:" + InstallerSource);
                    System.Diagnostics.Process Proc = new System.Diagnostics.Process();
                    Proc.StartInfo.FileName = InstallerSource;
                    Proc.StartInfo.CreateNoWindow = false;
                    Proc.Start();
                    Console.WriteLine("started:" + InstallerSource);

                    // ====================================================================================
                    AutomationElement rootElement = AutomationElement.RootElement;
                    #region Install EtriccCore
                    Console.WriteLine("Searching for main window");

                    PropertyCondition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
                    AutomationElement appElement = rootElement.FindFirst(TreeScope.Children, condition);

                    DateTime startTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - startTime;
                    while (appElement == null && mTime.TotalSeconds < 60)
                    {
                        EtriccUtilities.Wait(2);
                        appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                        mTime = DateTime.Now - startTime;
                        if (mTime.TotalSeconds > 60)
                        {
                            System.Windows.Forms.MessageBox.Show("After one minute no Installer Window Form found");
                            return;
                        }
                    }

                    // (1) Welcom Main window
                    Console.WriteLine("EtriccProgram Main Form found ...");
                    Console.WriteLine("1 Searching next button...");
                    AutomationElement btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                    if (btnNext != null)
                    {   // (2) Components
                        AUIUtilities.ClickElement(btnNext);
                        appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                        Console.WriteLine("Welcom Etricc Core window opend...");
                        Console.WriteLine("2 Searching next button...");
                        btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                        if (btnNext != null)
                        {
                            AUIUtilities.ClickElement(btnNext);
                            EtriccUtilities.Wait(3);
                            appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                            Console.WriteLine("Confirm Etricc Core window opend...");
                            Console.WriteLine("3 Searching Next button...");
                            btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                            if (btnNext != null)
                            {
                                AUIUtilities.ClickElement(btnNext);
                                EtriccUtilities.Wait(3);
                                appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                Console.WriteLine("Confirm Etricc Core window opend...");
                                Console.WriteLine("4 Searching close button...");

                                // Wait until Close button Found
                                appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);

                                System.Windows.Automation.Condition c2 = new AndCondition(
                                        new PropertyCondition(AutomationElement.NameProperty, "Close"),
                                        new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                                    );
                                AutomationElement aeBtnClose
                                    = appElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);

                                while (aeBtnClose == null)
                                {
                                    EtriccUtilities.Wait(5);
                                    Console.WriteLine("Wait until Close button found...");
                                    appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                    if (appElement == null)
                                        Console.WriteLine("Installer Window  not found");
                                    else
                                    {
                                        aeBtnClose = appElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                                        Console.WriteLine("Installer Window found: " + appElement.Current.Name);
                                    }
                                }
                                Console.WriteLine("Close button found... ---> Close Installer Window");
                                AUIUtilities.ClickElement(aeBtnClose);
                            }
                        }
                    }
                    #endregion
                }

                Thread.Sleep(5000);
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\nInstall Etricc HostTest: Pass");
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);

                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
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
        #region SystemOverviewQuery
        public static void SystemOverviewQuery(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            TestCheck = ConstCommon.TEST_PASS;
            result = ConstCommon.TEST_UNDEFINED;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            try
            {
                #region open system overview
               
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, SYSTEM_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {
                    Input.MoveToAndClick(aePanelLink);
                    Thread.Sleep(10000);
                    // wait extra one minute is project is TestProject.zip
                    if (sProjectFile.IndexOf("TestProject") >= 0)
                    {
                        sStartTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - sStartTime;
                        while (mTime.Seconds < 20)
                        {
                            Thread.Sleep(2000);
                            mTime = DateTime.Now - sStartTime;
                            Console.WriteLine("wait extra time(sec) : " + mTime.Seconds);
                        }
                    }
                }
                #endregion

                // Find System Overview Window
                Condition cWindow = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, SYSTEM_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the SystemOverview element.
                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cWindow);
                if (aeOverview == null)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = SYSTEM_OVERVIEW + " Window not found";
                    Console.WriteLine(sErrorMessage);
                }
                else
                {
                    string ms = SYSTEM_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                }

                Thread.Sleep(2000);

                #region resize root
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    System.Drawing.Rectangle rec = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
                    int Width = rec.Width;
                    int Height = rec.Height;

                    TransformPattern tranform =
                    root.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                    if (tranform != null)
                        tranform.Move(0, 0);

                    Thread.Sleep(3000);
                    tranform.Resize((Width - Width * 0.3), (Height - Height*0.2));

                    Thread.Sleep(3000);
                    /*
                    if (root.Current.BoundingRectangle.Width == Width &&
                        root.Current.BoundingRectangle.Height == (Height - 60))
                    {
                        Console.WriteLine("\nTest scenario SystemOverviewQuery: Pass1");
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        sErrorMessage = "current width=" + root.Current.BoundingRectangle.Width
                            + " --- "
                            + "current height=" + root.Current.BoundingRectangle.Height;
                        Console.WriteLine("current width=" + root.Current.BoundingRectangle.Width);
                        Console.WriteLine("current height=" + root.Current.BoundingRectangle.Height);
                        Console.WriteLine("\nTest scenario Resize: *FAIL*");
                        TestCheck = ConstCommon.TEST_FAIL;
                    }*/
                }
                #endregion

                #region open query screen
                Console.WriteLine("open query screen....");
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Condition Cond = new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.HelpTextProperty, "Show the Agv query window")
                    );

                    // Find the element.
                    AutomationElement aeQuery = root.FindFirst(TreeScope.Element | TreeScope.Descendants, Cond);
                    if (aeQuery == null)
                    {
                        result = ConstCommon.TEST_FAIL;
                        sErrorMessage = SYSTEM_OVERVIEW_QUERY + ": query button not found";
                        Console.WriteLine("FindElementByID failed:" + SYSTEM_OVERVIEW_QUERY);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        return;
                    }
                    else
                    {
                        string ms = SYSTEM_OVERVIEW_QUERY + ": query button found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        Console.WriteLine(ms);
                        Epia3Common.WriteTestLogMsg(slogFilePath, ms, sOnlyUITest);

                        Point pt = aeQuery.GetClickablePoint();
                        Input.MoveTo(pt);
                        Console.WriteLine("moved to X: " + pt.X);
                        Thread.Sleep(500);
                        TestTools.AUIUtilities.ClickElement(aeQuery);
                        Thread.Sleep(500);
                        TestTools.AUIUtilities.ClickElement(aeQuery);
                        Thread.Sleep(3000);
                    }
                }
                #endregion

                #region Find Query window
                AutomationElement aeQueryWindow = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find System Overview Window
                    Condition cQueryWindow = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Agv query"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                   );

                    // Find the SystemOverview element.
                    DateTime mStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - mStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalSeconds + "    testcheck = " + TestCheck);
                    aeQueryWindow = null;
                    while (aeQueryWindow == null && mTime.TotalSeconds < 60)
                    {
                        aeQueryWindow = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cQueryWindow);
                        Thread.Sleep(2000);
                        mTime = DateTime.Now - mStartTime;
                        Console.WriteLine(" time2 is :" + mTime.TotalSeconds);
                    }
                    if (aeQueryWindow == null)
                    {
                        result = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Query window not found";
                        Console.WriteLine("FindElementByID failed:" + SYSTEM_OVERVIEW_QUERY);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        return;
                    }
                    else
                    {
                        string ms = SYSTEM_OVERVIEW_QUERY + " Query window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        Console.WriteLine(ms);
                        Epia3Common.WriteTestLogMsg(slogFilePath, ms, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                }
                #endregion

                #region Check route cost tracking
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find System Overview Window
                    System.Windows.Automation.Condition cQueryRouteCostTrackingItem = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Route cost/tracking"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem)
                   );
                   
                    // Find the SystemOverview element.
                    //AutomationElement aeQueryRouteCostTrackingItem = aeQueryWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cQueryRouteCostTrackingItem);
                    AutomationElement aeQueryRouteCostTrackingItem = aeQueryWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, TestTools.ShellUIConst.CButtonByName("Route cost/tracking"));
                    if (aeQueryRouteCostTrackingItem == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "QueryRouteCostTrackingItem not found";
                        Console.WriteLine("FindElementByID failed:" + SYSTEM_OVERVIEW_QUERY);
                    }
                    else
                    {
                        string ms = SYSTEM_OVERVIEW_QUERY + " QueryRouteCostTrackingItem found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        Console.WriteLine(ms);
                        TestCheck = ConstCommon.TEST_PASS;
                        Thread.Sleep(3000);
                        Point pt = AUIUtilities.GetElementCenterPoint(aeQueryRouteCostTrackingItem);
                        Input.MoveToAndClick(pt);
                        //AUIUtilities.ClickElement(aeQueryRouteCostTrackingItem);
                    }
                }
                #endregion

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    result = ConstCommon.TEST_PASS;
                    string ms = testname + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                MessageBox.Show(sErrorMessage, "Exception", MessageBoxButtons.OK);
            }
        }
        #endregion SystemOverviewQuery
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region SystemOverviewDisplay
        public static void SystemOverviewDisplay(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sErrorMessage = string.Empty;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            AutomationElement aePanelLink = null;
            AutomationElement aeWindow = null;
            AutomationElement aeOverview = null;
            try
            {
                #region // open System overview
                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                if (root != null)
                {
                    root.SetFocus();
                    AUICommon.ClearDisplayedScreens(root, 2);
                    aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, SYSTEM_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Input.MoveToAndClick(aePanelLink);
                        Thread.Sleep(10000);
                        // wait extra one minute if project is TestProject.zip
                        if (sProjectFile.IndexOf("TestProject") >= 0)
                        {
                            sStartTime = DateTime.Now;
                            TimeSpan mTime = DateTime.Now - sStartTime;
                            while (mTime.Seconds < 20)
                            {
                                Thread.Sleep(2000);
                                mTime = DateTime.Now - sStartTime;
                                Console.WriteLine("wait extra time(sec) : " + mTime.Seconds);
                            }
                        }
                    }
                }
                else
                {
                    sErrorMessage = "aeWindow not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                #endregion

                #region// Find System Overview Window
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Condition cWindow = new AndCondition(new PropertyCondition(AutomationElement.NameProperty, SYSTEM_OVERVIEW_TITLE),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                    // Find the SystemOverview element.
                    aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cWindow);
                    if (aeOverview == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = SYSTEM_OVERVIEW + " Window not found";
                        Console.WriteLine("FindElementByID failed:" + SYSTEM_OVERVIEW);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                        string ms = SYSTEM_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        Console.WriteLine(ms);
                    }
                }
                #endregion

                bool status = true;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region control menuitem
                    Console.WriteLine("Overview name:" + aeOverview.Current.Name);
                    Point windowPt = new Point(aeOverview.Current.BoundingRectangle.TopLeft.X+20, aeOverview.Current.BoundingRectangle.TopLeft.Y+50);
                    Input.MoveTo(windowPt);
                    Thread.Sleep(2000);
                    Input.MoveToAndRightClick(windowPt);
                    Thread.Sleep(3000);
                    string[] SystemOverviewAllItems = new string[] { "Agv Traffic", "Show Locked Segments", "Show Locked Track", "Show Requested Track", "Show Leave Track", 
                       "Show Hull"
                    };

                    string[] VisibleItems = new string[] { "Agv Traffic", "Agv Tooltips" };  /// level 1 menuItem
                    //status = ProjBasicUI.ValidateSystemOverviewMenuItemActionElement("MainForm", VisibleItems, ref sErrorMessage, "Visible");
                    if (status == false)
                    {
                        sErrorMessage = "SystemOverviewVisibleActionMenuitems:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Level 2 MenuItems --> SystemOverviewAllItems: Level1MenuItems, Level2MenuItems1, Level2MenuItems2, Level2MenuItems3, 
                        status = ProjBasicUI.ValidateSystemOverviewMenuItemActionElement("MainForm", SystemOverviewAllItems, ref sErrorMessage, "Visible");
                        if (status == false)
                        {
                            sErrorMessage = "SystemOverviewAllItems:" + sErrorMessage;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    #endregion
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region control layer menuitem
                    // Find the layers button element.

                    AutomationElement aeLayersBtn = AUIUtilities.FindElementByID("cmbLayers", aeOverview);
                    if (aeLayersBtn == null)
                    {
                        sErrorMessage = SYSTEM_OVERVIEW + ": Layers button not found";
                        Console.WriteLine("FindElementByID failed:" + SYSTEM_OVERVIEW);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        string ms = SYSTEM_OVERVIEW + ":Layers button found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        Console.WriteLine(ms);
                        Epia3Common.WriteTestLogMsg(slogFilePath, ms, sOnlyUITest);

                        Point pt = aeLayersBtn.GetClickablePoint();
                        Input.MoveTo(pt);
                        Console.WriteLine("moved to X: " + pt.X);
                        Thread.Sleep(500);
                        TestTools.AUIUtilities.ClickElement(aeLayersBtn);
                        Thread.Sleep(500);

                        string[] LayersAllItems = new string[] { "Agvs", "Background (Layout LAYOUTFLV )", "Backward paths (Layout LAYOUTFLV )", "Backward paths (Layout LAYOUTTUG )", 
                       "Field functions","Forward paths (Layout LAYOUTFLV )", "Forward paths (Layout LAYOUTTUG )", "Loads", "Locations", "Navigation beacons", "Stations"
                        };
                        status = ProjBasicUI.ValidateSystemOverviewCheckboxMenuItems("MainForm", LayersAllItems, ref sErrorMessage, "Visible");
                        TestTools.AUIUtilities.ClickElement(aeLayersBtn);
                        Thread.Sleep(500);
                        if (status == false)
                        {
                            sErrorMessage = "SystemOverviewVisibleActionMenuitems:" + sErrorMessage;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    #endregion
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region control Tooltips menuitem
                    AutomationElement aeTooltipsBtn = AUIUtilities.FindElementByID("cmbTooltips", aeOverview);
                    if (aeTooltipsBtn == null)
                    {
                        sErrorMessage = SYSTEM_OVERVIEW + ": Tooltips button not found";
                        Console.WriteLine("FindElementByID failed:" + SYSTEM_OVERVIEW);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        string ms = SYSTEM_OVERVIEW + ":TooltipsBtn button found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        Console.WriteLine(ms);
                        Epia3Common.WriteTestLogMsg(slogFilePath, ms, sOnlyUITest);

                        Point pt = aeTooltipsBtn.GetClickablePoint();
                        Input.MoveTo(pt);
                        Console.WriteLine("moved to X: " + pt.X);
                        Thread.Sleep(500);
                        TestTools.AUIUtilities.ClickElement(aeTooltipsBtn);
                        Thread.Sleep(1000);

                        string[] TooltipsAllItems = new string[] { "Agvs", 
                            "Field functions", "Loads", "Locations", "Stations"};
                        status = ProjBasicUI.ValidateSystemOverviewCheckboxMenuItems("MainForm", TooltipsAllItems, ref sErrorMessage, "Visible");
                        TestTools.AUIUtilities.ClickElement(aeTooltipsBtn);
                        Thread.Sleep(1000);
                        if (status == false)
                        {
                            sErrorMessage = "SystemOverviewVisibleActionMenuitems:" + sErrorMessage;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        } 
                    }
                    #endregion
                }
                
                Console.WriteLine("   ----    Validate LegendeInfo UI");
                Thread.Sleep(1000);
                DateTime testCaseCreateDate = new DateTime(2011, 11, 17);
                if (sBuildNr.IndexOf("Hotfix") >= 0)
                {
                    if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 8, 29), ref sErrorMessage) == true)
                    {
                        sErrorMessage = "Release date of this hotfix is earlier then this test case created date, Not test Legende";
                        Epia3Common.WriteTestLogTitle(slogFilePath, sErrorMessage, Counter, sOnlyUITest);
                    }
                    else
                    {
                        #region // Validate LegendeInfo UI
                        AutomationElement aeLegendeButton = null;
                        AutomationElement aeLegendeTreeView = null;
                        if (TestCheck == ConstCommon.TEST_PASS)
                        {
                            // Map legende
                            string legendeBtnId = "LegendeButton";
                            string mapLegendTreeViewId = "MapLegend"; // check IsOffScreen 
                            aeLegendeButton = AUIUtilities.FindElementByID(legendeBtnId, aeOverview);
                            aeLegendeTreeView = AUIUtilities.FindElementByID(mapLegendTreeViewId, aeOverview);
                            if (aeLegendeButton == null || aeLegendeTreeView == null)
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = "LegendeButton or aeLegendeTreeView not found";
                                Console.WriteLine(sErrorMessage);
                            }
                            else
                            {
                                if (aeLegendeTreeView.Current.IsOffscreen)
                                {
                                    Console.WriteLine("   ----    aeLegendeTreeView.Current.IsOffscreen   --> Click(aeLegendeButton) ");
                                    Thread.Sleep(1000);
                                    Input.MoveToAndClick(aeLegendeButton);
                                    Thread.Sleep(2000);
                                }
                            }
                        }

                        // validate legende treeview
                        Console.WriteLine("   ----    Validate legende treeview");
                        if (TestCheck == ConstCommon.TEST_PASS)
                        {
                            aeWindow = EtriccUtilities.GetCategoryWindow("System overview", ref sErrorMessage);
                            aeLegendeTreeView = AUIUtilities.FindElementByID("MapLegend", aeWindow);
                            if (aeLegendeTreeView.Current.IsOffscreen == true)
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = "aeLegendeTreeView is stell OffScreen after legende button clicked";
                                Console.WriteLine(sErrorMessage);
                            }
                            else
                            {
                                Condition cTreeItem = new AndCondition(new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                                AutomationElementCollection elementNodes = aeLegendeTreeView.FindAll(TreeScope.Element | TreeScope.Descendants, cTreeItem);
                                //TreeWalker walker = TreeWalker.ControlViewWalker;
                                //AutomationElement elementNode = walker.GetFirstChild(aeLegendeTreeView);
                                for (int i = 0; i < elementNodes.Count; i++)
                                {
                                    Console.WriteLine("aeTreeView node name is: " + elementNodes[i].Current.Name);
                                    try
                                    {
                                        ExpandCollapsePattern ep = (ExpandCollapsePattern)elementNodes[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                        ep.Expand();
                                        Thread.Sleep(1000);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("aeLegendeTreeView can not expaned: " + aeLegendeTreeView.Current.Name);
                                        Console.WriteLine("aeLegendeTreeView can not expaned: " + ex.Message);
                                    }//elementNode = walker.GetNextSibling(elementNode);
                                }
                            }
                        }
                        #endregion

                        // check mini map                
                        #region // Validate MiniMap UI
                        AutomationElement aeMiniMapButton = null;
                        AutomationElement aeMiniMapScreen = null;
                        if (TestCheck == ConstCommon.TEST_PASS)
                        {
                            aeOverview = EtriccUtilities.GetCategoryWindow("System overview", ref sErrorMessage);
                            // Map legende
                            string miniMapBtnId = "MiniMapButton";
                            string miniMapScreenId = "MiniMap"; // if IsOffScreen true --> click button
                            aeMiniMapButton = AUIUtilities.FindElementByID(miniMapBtnId, aeOverview);
                            aeMiniMapScreen = AUIUtilities.FindElementByID(miniMapScreenId, aeOverview);
                            if (aeMiniMapButton == null || aeMiniMapScreen == null)
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = "aeMiniMapButton or aeMiniMapScreen not found";
                                Console.WriteLine(sErrorMessage);
                            }
                            else
                            {
                                if (aeMiniMapScreen.Current.IsOffscreen)
                                {
                                    Input.MoveToAndClick(aeMiniMapButton);
                                    Thread.Sleep(2000);
                                }
                            }
                        }

                        // validate Mini Map Screen
                        if (TestCheck == ConstCommon.TEST_PASS)
                        {
                            aeOverview = EtriccUtilities.GetCategoryWindow("System overview", ref sErrorMessage);
                            aeMiniMapScreen = AUIUtilities.FindElementByID("MiniMap", aeWindow);
                            if (aeMiniMapScreen.Current.IsOffscreen == true)
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = "aeMiniMapScreen is stell OffScreen after minimap button clicked";
                                Console.WriteLine(sErrorMessage);
                            }
                        }
                        #endregion
                    }
                }
                else
                {
                    #region // Validate LegendeInfo UI
                    AutomationElement aeLegendeButton = null;
                    AutomationElement aeLegendeTreeView = null;
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        // Map legende
                        string legendeBtnId = "LegendeButton";
                        string mapLegendTreeViewId = "MapLegend"; // check IsOffScreen 
                        aeLegendeButton = AUIUtilities.FindElementByID(legendeBtnId, aeOverview);
                        aeLegendeTreeView = AUIUtilities.FindElementByID(mapLegendTreeViewId, aeOverview);
                        if (aeLegendeButton == null || aeLegendeTreeView == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "LegendeButton or aeLegendeTreeView not found";
                            Console.WriteLine(sErrorMessage);
                        }
                        else
                        {
                            if (aeLegendeTreeView.Current.IsOffscreen)
                            {
                                Input.MoveToAndClick(aeLegendeButton);
                                Thread.Sleep(2000);
                            }
                        }
                    }

                    // validate legende treeview
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        aeWindow = EtriccUtilities.GetCategoryWindow("System overview", ref sErrorMessage);
                        aeLegendeTreeView = AUIUtilities.FindElementByID("MapLegend", aeWindow);
                        if (aeLegendeTreeView.Current.IsOffscreen == true)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "aeLegendeTreeView is stell OffScreen after legende button clicked";
                            Console.WriteLine(sErrorMessage);
                        }
                        else
                        {
                            Condition cTreeItem = new AndCondition(new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem));
                            AutomationElementCollection elementNodes = aeLegendeTreeView.FindAll(TreeScope.Element | TreeScope.Descendants, cTreeItem);
                            //TreeWalker walker = TreeWalker.ControlViewWalker;
                            //AutomationElement elementNode = walker.GetFirstChild(aeLegendeTreeView);
                            for (int i = 0; i < elementNodes.Count; i++)
                            {
                                Console.WriteLine("aeTreeView node name is: " + elementNodes[i].Current.Name);
                                try
                                {
                                    ExpandCollapsePattern ep = (ExpandCollapsePattern)elementNodes[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                    ep.Expand();
                                    Thread.Sleep(1000);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("aeLegendeTreeView can not expaned1: " + aeLegendeTreeView.Current.Name);
                                    Console.WriteLine("aeLegendeTreeView can not expaned2: " + ex.Message);
                                }//elementNode = walker.GetNextSibling(elementNode);
                                
                            }
                        }
                    }
                    #endregion
                    // check mini map                
                    #region // Validate MiniMap UI
                    AutomationElement aeMiniMapButton = null;
                    AutomationElement aeMiniMapScreen = null;
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        aeOverview = EtriccUtilities.GetCategoryWindow("System overview", ref sErrorMessage);
                        // Map legende
                        string miniMapBtnId = "MiniMapButton";
                        string miniMapScreenId = "MiniMap"; // if IsOffScreen true --> click button
                        aeMiniMapButton = AUIUtilities.FindElementByID(miniMapBtnId, aeOverview);
                        aeMiniMapScreen = AUIUtilities.FindElementByID(miniMapScreenId, aeOverview);
                        if (aeMiniMapButton == null || aeMiniMapScreen == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "aeMiniMapButton or aeMiniMapScreen not found";
                            Console.WriteLine(sErrorMessage);
                        }
                        else
                        {
                            if (aeMiniMapScreen.Current.IsOffscreen)
                            {
                                Input.MoveToAndClick(aeMiniMapButton);
                                Thread.Sleep(2000);
                            }
                        }
                    }

                    // validate Mini Map Screen
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        aeOverview = EtriccUtilities.GetCategoryWindow("System overview", ref sErrorMessage);
                        aeMiniMapScreen = AUIUtilities.FindElementByID("MiniMap", aeWindow);
                        if (aeMiniMapScreen.Current.IsOffscreen == true)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "aeMiniMapScreen is stell OffScreen after minimap button clicked";
                            Console.WriteLine(sErrorMessage);
                        }
                    }
                    #endregion
                }
                //
                Thread.Sleep(2000);
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }
        }
        #endregion SystemOverviewDisplay
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region AgvOverviewDisplay
        public static void AgvOverviewDisplay(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            AutomationElement aePanelLink = null;
            try
            {
                AUICommon.ClearDisplayedScreens(root, 2);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Input.MoveToAndClick(aePanelLink);

                Thread.Sleep(10000);

                // Find AGV Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, AGV_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the AGVOverview element.
                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = AGV_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + AGV_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_PASS;
                    string ms = AGV_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }


        }
        #endregion AgvOverviewDisplay
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LocationOverviewDisplay
        public static void LocationOverviewDisplay(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            AutomationElement aeOverview = null;
            AutomationElement aePanelLink = null;
            if (sOnlyUITest)
                root = EtriccUtilities.GetMainWindow("MainForm");

            try
            {
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, LOCATION_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    Input.MoveToAndClick(aePanelLink);

                AutomationElement aeGrid = null;
                int ky = 0;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(5000);
                    aeOverview = ProjBasicUI.GetSelectedOverviewWindow(LOCATION_OVERVIEW_TITLE, ref sErrorMessage);
                    while (aeOverview == null && ky < 5)
                    {
                        Console.WriteLine("wait until selected " + LOCATION_OVERVIEW_TITLE + " window open :" + ky++);
                        aeOverview = ProjBasicUI.GetSelectedOverviewWindow(LOCATION_OVERVIEW_TITLE, ref sErrorMessage);
                        Thread.Sleep(5000);
                    }

                    if (aeOverview == null)
                    {
                        Console.WriteLine("Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Find Location GridView
                        aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                        if (aeGrid == null)
                        {
                            sErrorMessage = "Find LocationDataGridView failed:" + "LocationDataGridView";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }
               
                AutomationElement aeCell = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(5000);
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row 0";
                    // Get the Element with the Row Col Coordinates
                    aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                    if (aeCell == null)
                    {
                        sErrorMessage = "Find LocationDataGridView aeCell failed:" + cellname;
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // // find cell value
                        Thread.Sleep(3000);
                        string TransportValue = string.Empty;
                        try
                        {
                            ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            TransportValue = vp.Current.Value;
                            Console.WriteLine("Get element.Current Value:" + TransportValue);
                        }
                        catch (System.NullReferenceException)
                        {
                            TransportValue = string.Empty;
                        }

                        if (TransportValue == null || TransportValue == string.Empty)
                        {
                            sErrorMessage = "LocationDataGridView aeCell Value not found:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else if (!TransportValue.Equals("FLV_L101"))
                        {

                            sErrorMessage = "LocationDataGridView aeCell Value not equal to FLV_L101, but:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                // open Select a node from this point as start node for the query screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                    Input.MoveToAndRightClick(point);
                    Thread.Sleep(3000);

                    // find Select a node from this point as start node for the query screen point
                    System.Windows.Automation.Condition cM = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Select a node from this point as start node for the query screen"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                   );

                    // Find the MenuItem Select a node from this point as start node for the query screen element
                    AutomationElement aeMenuItemSelect = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
                    if (aeMenuItemSelect != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemSelect.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemSelect.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemSelect.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemSelect.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemSelect);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Select a node from this point as start node for the query screen menu item not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    Thread.Sleep(2000);
                }

                string selectWindowId = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.SelectNode";
                string btnCancelId = "m_BtnClose";
                Point cancelBtnPoint = new Point();


                AutomationElement aeSelectWindow = null;
                AutomationElement aeListNodes = null;
                AutomationElement aeBtnCancel = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find the SelectWindow element
                    aeSelectWindow = AUIUtilities.FindElementByID(selectWindowId, root);
                    //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                    if (aeSelectWindow == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = " Selected node Window not found";
                        Console.WriteLine(sErrorMessage);
                    }
                    else // find list node
                    {
                        aeListNodes = AUIUtilities.FindElementByID("m_LstNodes", aeSelectWindow);

                        //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                        if (aeListNodes == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = " aeListNodes not found";
                            Console.WriteLine(sErrorMessage);
                        }
                        else
                        {
                            aeBtnCancel = AUIUtilities.FindElementByID(btnCancelId, aeSelectWindow);
                            cancelBtnPoint = AUIUtilities.GetElementCenterPoint(aeBtnCancel);
                        }
                    }
                }

                // select list item FLV_L1
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    AutomationElement item = AUIUtilities.FindElementByName("FLV_L1", aeListNodes);
                    if (item != null)
                    {
                        Console.WriteLine("FLV_L1" + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(2000);
                        SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemPattern.Select();
                        Thread.Sleep(2000);
                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));
                        Thread.Sleep(2000);
                    }
                    else
                    {
                        sErrorMessage = "Finding Command item " + "FLV_L1" + "  failed";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // check main form enable
                root = EtriccUtilities.GetMainWindow("MainForm");
                if (root.Current.IsEnabled)
                {
                    Console.WriteLine("MainForm is enabled test OK:");
                    Thread.Sleep(5000);
                }
                else
                {
                    EtriccUtilities.TryToGetErrorMessageAndCloseErrorScreen(ref sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    // cancel select window 
                    Input.MoveToAndClick(cancelBtnPoint);
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }
        }
        #endregion LocationOverviewDisplay
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LocationOverviewDisplayEndNode
        public static void LocationOverviewDisplayEndNode(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            AutomationElement aeOverview = null;
            AutomationElement aePanelLink = null;
            if (sOnlyUITest)
                root = EtriccUtilities.GetMainWindow("MainForm");

            try
            {
                AUICommon.ClearDisplayedScreens(root, 5);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, LOCATION_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    Input.MoveToAndClick(aePanelLink);

                Thread.Sleep(10000);
                // Find Location Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, LOCATION_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                AutomationElement aeGrid = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find the LocationOverview element.
                    aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeOverview == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = LOCATION_OVERVIEW + " Window not found";
                        Console.WriteLine("FindElementByID failed:" + "Locations");
                    }
                    else
                    {
                        // Find Location GridView
                        aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                        if (aeGrid == null)
                        {
                            sErrorMessage = "Find LocationDataGridView failed:" + "LocationDataGridView";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                AutomationElement aeCell = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row 1";
                    // Get the Element with the Row Col Coordinates
                    aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                    if (aeCell == null)
                    {
                        sErrorMessage = "Find LocationDataGridView aeCell failed:" + cellname;
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // // find cell value
                        Thread.Sleep(3000);
                        string TransportValue = string.Empty;
                        try
                        {
                            ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            TransportValue = vp.Current.Value;
                            Console.WriteLine("Get element.Current Value:" + TransportValue);
                        }
                        catch (System.NullReferenceException)
                        {
                            TransportValue = string.Empty;
                        }

                        if (TransportValue == null || TransportValue == string.Empty)
                        {
                            sErrorMessage = "LocationDataGridView aeCell Value not found:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else if (!TransportValue.Equals("FLV_L102"))
                        {

                            sErrorMessage = "LocationDataGridView aeCell Value not equal to FLV_L102, but:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                // open Select a node from this point as start node for the query screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                    Input.MoveToAndRightClick(point);
                    Thread.Sleep(3000);

                    // find Select a node from this point as start node for the query screen point
                    System.Windows.Automation.Condition cM = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Select a node from this point as end node for the query screen"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                   );

                    // Find the MenuItem Select a node from this point as start node for the query screen element
                    AutomationElement aeMenuItemSelect = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
                    if (aeMenuItemSelect != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemSelect.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemSelect.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemSelect.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemSelect.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemSelect);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Select a node from this point as end node for the query screen menu item not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    Thread.Sleep(2000);
                }

                string selectWindowId = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.SelectNode";
                string btnCancelId = "m_BtnClose";
                Point cancelBtnPoint = new Point();


                AutomationElement aeSelectWindow = null;
                AutomationElement aeListNodes = null;
                AutomationElement aeBtnCancel = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find the SelectWindow element
                    aeSelectWindow = AUIUtilities.FindElementByID(selectWindowId, root);
                    //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                    if (aeSelectWindow == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = " Selected node Window not found";
                        Console.WriteLine(sErrorMessage);
                    }
                    else // find list node
                    {
                        aeListNodes = AUIUtilities.FindElementByID("m_LstNodes", aeSelectWindow);

                        //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                        if (aeListNodes == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = " aeListNodes not found";
                            Console.WriteLine(sErrorMessage);
                        }
                        else
                        {
                            aeBtnCancel = AUIUtilities.FindElementByID(btnCancelId, aeSelectWindow);
                            cancelBtnPoint = AUIUtilities.GetElementCenterPoint(aeBtnCancel);
                        }
                    }
                }

                // select list item FLV_L1
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    AutomationElement item = AUIUtilities.FindElementByName("FLV_L2", aeListNodes);
                    if (item != null)
                    {
                        Console.WriteLine("FLV_L2" + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(2000);
                        SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemPattern.Select();
                        Thread.Sleep(2000);
                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));
                        Thread.Sleep(2000);
                    }
                    else
                    {
                        sErrorMessage = "Finding Command item " + "FLV_L2" + "  failed";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }



                // check main form enable
                root = EtriccUtilities.GetMainWindow("MainForm");
                if (root.Current.IsEnabled)
                {
                    Console.WriteLine("MainForm is enabled test OK:");
                    Thread.Sleep(5000);
                }
                else
                {
                    EtriccUtilities.TryToGetErrorMessageAndCloseErrorScreen(ref sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    // cancel select window 
                    Input.MoveToAndClick(cancelBtnPoint);

                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }


        }
        #endregion LocationOverviewDisplayEndNode
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region StationOverviewDisplay
        public static void StationOverviewDisplay(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            AutomationElement aeOverview = null;
            AutomationElement aePanelLink = null;
            if (sOnlyUITest)
                root = EtriccUtilities.GetMainWindow("MainForm");

            try
            {
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, STATION_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    Input.MoveToAndClick(aePanelLink);

                Thread.Sleep(10000);
                // Find Station Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, STATION_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                AutomationElement aeGrid = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find the StationOverview element.
                    aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeOverview == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = STATION_OVERVIEW + " Window not found";
                        Console.WriteLine("FindElementByID failed:" + "Stations");
                    }
                    else
                    {
                        // Find Location GridView
                        aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                        if (aeGrid == null)
                        {
                            sErrorMessage = "Find StationDataGridView failed:" + "StationDataGridView";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                AutomationElement aeCell = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row 0";
                    // Get the Element with the Row Col Coordinates
                    aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                    if (aeCell == null)
                    {
                        sErrorMessage = "Find StationDataGridView aeCell failed:" + cellname;
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // // find cell value
                        Thread.Sleep(3000);
                        string TransportValue = string.Empty;
                        try
                        {
                            ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            TransportValue = vp.Current.Value;
                            Console.WriteLine("Get element.Current Value:" + TransportValue);
                        }
                        catch (System.NullReferenceException)
                        {
                            TransportValue = string.Empty;
                        }

                        if (TransportValue == null || TransportValue == string.Empty)
                        {
                            sErrorMessage = "StationDataGridView aeCell Value not found:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else if (!TransportValue.Equals("X01"))
                        {

                            sErrorMessage = "StationDataGridView aeCell Value not equal to X01, but:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                // open Select a node from this point as start node for the query screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                    Input.MoveToAndRightClick(point);
                    Thread.Sleep(3000);

                    // find Select a node from this point as start node for the query screen point
                    System.Windows.Automation.Condition cM = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Select a node from this point as start node for the query screen"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                   );

                    // Find the MenuItem Select a node from this point as start node for the query screen element
                    AutomationElement aeMenuItemSelect = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
                    if (aeMenuItemSelect != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemSelect.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemSelect.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemSelect.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemSelect.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemSelect);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Select a node from this point as start node for the query screen menu item not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    Thread.Sleep(2000);
                }

                string selectWindowId = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.SelectNode";
                string btnCancelId = "m_BtnClose";
                Point cancelBtnPoint = new Point();


                AutomationElement aeSelectWindow = null;
                AutomationElement aeListNodes = null;
                AutomationElement aeBtnCancel = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find the SelectWindow element
                    aeSelectWindow = AUIUtilities.FindElementByID(selectWindowId, root);
                    //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                    if (aeSelectWindow == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = " Selected node Window not found";
                        Console.WriteLine(sErrorMessage);
                    }
                    else // find list node
                    {
                        aeListNodes = AUIUtilities.FindElementByID("m_LstNodes", aeSelectWindow);

                        //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                        if (aeListNodes == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = " aeListNodes not found";
                            Console.WriteLine(sErrorMessage);
                        }
                        else
                        {
                            aeBtnCancel = AUIUtilities.FindElementByID(btnCancelId, aeSelectWindow);
                            cancelBtnPoint = AUIUtilities.GetElementCenterPoint(aeBtnCancel);
                        }
                    }
                }

                // select list item PRK_FLV
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    AutomationElement item = AUIUtilities.FindElementByName("PRK_FLV", aeListNodes);
                    if (item != null)
                    {
                        Console.WriteLine("PRK_FLV" + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(2000);
                        SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemPattern.Select();
                        Thread.Sleep(2000);
                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));
                        Thread.Sleep(2000);
                    }
                    else
                    {
                        sErrorMessage = "Finding Command item " + "PRK_FLV" + "  failed";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }



                // check main form enable
                root = EtriccUtilities.GetMainWindow("MainForm");
                if (root.Current.IsEnabled)
                {
                    Console.WriteLine("MainForm is enabled test OK:");
                    Thread.Sleep(5000);
                }
                else
                {
                    EtriccUtilities.TryToGetErrorMessageAndCloseErrorScreen(ref sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    // cancel select window 
                    Input.MoveToAndClick(cancelBtnPoint);

                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }


        }
        #endregion StationOverviewDisplay
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region StationOverviewDisplayEndNode
        public static void StationOverviewDisplayEndNode(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            AutomationElement aeOverview = null;
            AutomationElement aePanelLink = null;
            if (sOnlyUITest)
                root = EtriccUtilities.GetMainWindow("MainForm");

            try
            {
                AUICommon.ClearDisplayedScreens(root);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, STATION_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    Input.MoveToAndClick(aePanelLink);

                Thread.Sleep(10000);
                // Find Location Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, STATION_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                AutomationElement aeGrid = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find the LocationOverview element.
                    aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeOverview == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = STATION_OVERVIEW + " Window not found";
                        Console.WriteLine("FindElementByID failed:" + "Stations");
                    }
                    else
                    {
                        // Find Location GridView
                        aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                        if (aeGrid == null)
                        {
                            sErrorMessage = "Find StationDataGridView failed:" + "StationDataGridView";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                AutomationElement aeCell = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row 1";
                    // Get the Element with the Row Col Coordinates
                    aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                    if (aeCell == null)
                    {
                        sErrorMessage = "Find StationDataGridView aeCell failed:" + cellname;
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // // find cell value
                        Thread.Sleep(3000);
                        string TransportValue = string.Empty;
                        try
                        {
                            ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            TransportValue = vp.Current.Value;
                            Console.WriteLine("Get element.Current Value:" + TransportValue);
                        }
                        catch (System.NullReferenceException)
                        {
                            TransportValue = string.Empty;
                        }

                        if (TransportValue == null || TransportValue == string.Empty)
                        {
                            sErrorMessage = "StationDataGridView aeCell Value not found:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else if (!TransportValue.Equals("X02"))
                        {

                            sErrorMessage = "StationDataGridView aeCell Value not equal to X02, but:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                // open Select a node from this point as start node for the query screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                    Input.MoveToAndRightClick(point);
                    Thread.Sleep(3000);

                    // find Select a node from this point as start node for the query screen point
                    System.Windows.Automation.Condition cM = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Select a node from this point as end node for the query screen"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                   );

                    // Find the MenuItem Select a node from this point as start node for the query screen element
                    AutomationElement aeMenuItemSelect = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
                    if (aeMenuItemSelect != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemSelect.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemSelect.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemSelect.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemSelect.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemSelect);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Select a node from this point as end node for the query screen menu item not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    Thread.Sleep(2000);
                }

                string selectWindowId = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.SelectNode";
                string btnCancelId = "m_BtnClose";
                Point cancelBtnPoint = new Point();


                AutomationElement aeSelectWindow = null;
                AutomationElement aeListNodes = null;
                AutomationElement aeBtnCancel = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find the SelectWindow element
                    aeSelectWindow = AUIUtilities.FindElementByID(selectWindowId, root);
                    //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                    if (aeSelectWindow == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = " Selected node Window not found";
                        Console.WriteLine(sErrorMessage);
                    }
                    else // find list node
                    {
                        aeListNodes = AUIUtilities.FindElementByID("m_LstNodes", aeSelectWindow);

                        //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                        if (aeListNodes == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = " aeListNodes not found";
                            Console.WriteLine(sErrorMessage);
                        }
                        else
                        {
                            aeBtnCancel = AUIUtilities.FindElementByID(btnCancelId, aeSelectWindow);
                            cancelBtnPoint = AUIUtilities.GetElementCenterPoint(aeBtnCancel);
                        }
                    }
                }

                // select list item X02
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    AutomationElement item = AUIUtilities.FindElementByName("X02", aeListNodes);
                    if (item != null)
                    {
                        Console.WriteLine("X02" + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(2000);
                        SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemPattern.Select();
                        Thread.Sleep(2000);
                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));
                        Thread.Sleep(2000);
                    }
                    else
                    {
                        sErrorMessage = "Finding Command item " + "X02" + "  failed";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // check main form enable
                root = EtriccUtilities.GetMainWindow("MainForm");
                if (root.Current.IsEnabled)
                {
                    Console.WriteLine("MainForm is enabled test OK:");
                    Thread.Sleep(5000);
                }
                else
                {
                    EtriccUtilities.TryToGetErrorMessageAndCloseErrorScreen(ref sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    // cancel select window 
                    Input.MoveToAndClick(cancelBtnPoint);

                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }


        }
        #endregion StationOverviewDisplayEndNode
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region StationActionMenuitemsValidation
        public static void StationActionMenuitemsValidation(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            AutomationElement aeSelectedWindow = null;
            string StateValue = string.Empty;
            Point Pnt = new System.Windows.Point();
            try
            {
                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                if (root != null)
                {
                    root.SetFocus();
                    AUICommon.ClearDisplayedScreens(root, 2);
                    aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, STATION_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Pnt = AUIUtilities.GetElementCenterPoint(aePanelLink);
                        Input.MoveToAndClick(Pnt);
                        Thread.Sleep(5000);
                    }
                }
                else
                {
                    sErrorMessage = "aeWindow not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                AutomationElement aeGrid = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(STATION_OVERVIEW_TITLE, ref sErrorMessage);
                    if (aeSelectedWindow == null)
                    {
                        Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Find AGV GridView
                        aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeSelectedWindow);
                        if (aeGrid == null)
                        {
                            sErrorMessage = "Find StationDataGrid failed:" + AGV_GRIDDATA_ID;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row 0";
                    // Get the Element with the Row Col Coordinates
                    AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                    if (aeCell == null)
                    {
                        sErrorMessage = "Find StationDataGridView aeCell failed:" + cellname;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // find cell value
                        string StationValue = string.Empty;
                        try
                        {
                            ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            StationValue = vp.Current.Value;
                            Console.WriteLine("Get element.Current Value:" + StationValue);
                        }
                        catch (System.NullReferenceException)
                        {
                            StationValue = string.Empty;
                        }

                        if (StationValue == null || StationValue == string.Empty)
                        {
                            sErrorMessage = "StationDataGridView aeCell Value not found:" + cellname;
                            Console.WriteLine("StationDataGridView aeCell Value not found:" + cellname);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "StationDataGridView cell value not found:" + cellname, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Thread.Sleep(3000);
                            System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                            //
                            // get State Value 
                            //
                            string cellState = "Mode Row 0";
                            // Get the Element with the Row Col Coordinates
                            AutomationElement aeCellState = AUIUtilities.FindElementByName(cellState, aeGrid);
                            if (aeCellState == null)
                            {
                                sErrorMessage = "Find StationDataGridView aeCellState failed:" + cellState;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                Console.WriteLine("cellState StationDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                // find cellState value
                                try
                                {
                                    ValuePattern vp1 = aeCellState.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                    StateValue = vp1.Current.Value;
                                    Console.WriteLine("Get element.Current Value:" + StateValue);
                                }
                                catch (System.NullReferenceException)
                                {
                                    StateValue = string.Empty;
                                }

                                if (StateValue == null || StateValue == string.Empty)
                                {
                                    sErrorMessage = "StationDataGridView aeCell Value not found:" + cellState;
                                    Console.WriteLine("StationDataGridView aeCell Value not found:" + cellState);
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "StationDataGridView cell value not found:" + cellState, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                                else
                                {
                                    System.Windows.Point point2 = AUIUtilities.GetElementCenterPoint(aeCellState);
                                    Thread.Sleep(3000);
                                    Input.MoveToAndRightClick(point);
                                    Thread.Sleep(3000);
                                }
                            }
                        }
                    }
                }

                bool status = true;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    string[] StationAllItems = new string[] { "Station detail", "Select a node from this point as start node for the query screen",
                         "Select a node from this point as end node for the query screen", "Mode", "Disabled", 
                        "Normal",  "Copy"
                    };

                    /*string[] VisibleItems = new string[] { "Restart Agv", "Retire Agv", "Agv detail", "New Job", "New Week Plan", 
                        "Stop Agv", "Suspend Agv", "Cancel current Batch", "Jobs", "Battery charge plan", "Mode", "Engineering"
                    };*
                    status = ProjBasicUI.ValidateWindowMenuItemActionElement("MainForm", VisibleItems, ref sErrorMessage, "Visible");
                    if (status == false)
                    {
                        sErrorMessage = "AgvVisibleActionMenuitems:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {*/
                        status = ProjBasicUI.ValidateWindowMenuItemActionElement("MainForm", StationAllItems, ref sErrorMessage, "All");
                        if (status == false)
                        {
                            sErrorMessage = "StationAllActionMenuitems:" + sErrorMessage;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    //}
                }

                Thread.Sleep(1000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(1000);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine("\nTest scenario" + testname + " : Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                sErrorMessage = sErrorMessage.Trim();
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region StationModeControl
        public static void StationModeControl(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            try
            {
                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                if (root != null)
                {
                    root.SetFocus();
                    AUICommon.ClearDisplayedScreens(root, 2);
                    aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, STATION_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Point Pnt = AUIUtilities.GetElementCenterPoint(aePanelLink);
                        Input.MoveToAndClick(Pnt);
                        Thread.Sleep(5000);
                    }
                }
                else
                {
                    sErrorMessage = "aeWindow not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                AutomationElement aeOverview = null;
                // Find STATION Overview Window
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("000  Test mode :" + STATION_OVERVIEW);
                    System.Windows.Automation.Condition c = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, STATION_OVERVIEW_TITLE),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                   );

                    // Find the STATION Overview element.
                    aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeOverview == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = STATION_OVERVIEW + " Window not found";
                        Console.WriteLine("FindElementByID failed:" + STATION_OVERVIEW);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    }
                    else
                        Console.WriteLine("FindElementByID OK:" + STATION_OVERVIEW);
                }

                string testMode = "Disabled";
                Console.WriteLine("Test mode :" + STATION_OVERVIEW);
                for (int i = 0; i < 2 && (TestCheck == ConstCommon.TEST_PASS); i++)
                {
                    if (i == 1) testMode = "Normal";
                    #region  test change Mode
                    AutomationElement aeGrid = AUIUtilities.FindElementByID(DATAGRIDVIEW_ID, aeOverview);
                    if (aeGrid == null)
                    {
                        Console.WriteLine("Find LocationDataGridView failed:" + DATAGRIDVIEW_ID);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView failed:" + DATAGRIDVIEW_ID, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("StationDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        if (sProjectFile.IndexOf("TestProject") >= 0)
                        {
                            sErrorMessage = "testproject location too large, skip change location mode: ";
                            Console.WriteLine(sErrorMessage);
                            Thread.Sleep(2000);
                            result = ConstCommon.TEST_PASS;
                            Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
                            return;
                        }
                    }

                    Thread.Sleep(3000);
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row 2";
                    // Get the Element with the Row Col Coordinates
                    AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                    if (aeCell == null)
                    {
                        Console.WriteLine("Find StationDataGridView aeCell failed:" + cellname);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Find StationDataGridView cell failed:" + cellname, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                    else
                    {
                        Console.WriteLine("cell StationDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    }

                    Console.WriteLine("// find cell value :" + STATION_OVERVIEW);
                    // find cell value
                    string LocValue = string.Empty;
                    try
                    {
                        ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        LocValue = vp.Current.Value;
                        Console.WriteLine("Get element.Current Value:" + LocValue);
                    }
                    catch (System.NullReferenceException)
                    {
                        LocValue = string.Empty;
                    }

                    if (LocValue == null || LocValue == string.Empty)
                    {
                        sErrorMessage = "StationDataGridView aeCell Value not found:" + cellname;
                        Console.WriteLine("StationDataGridView aeCell Value not found:" + cellname);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "StationDataGridView cell value not found:" + cellname, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }

                    Thread.Sleep(3000);
                    //System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                    System.Windows.Point point = new Point();
                    try
                    {
                        point = aeCell.GetClickablePoint();
                        //point = AUIUtilities.GetElementCenterPoint(aeCell);
                    }
                    catch (Exception ex)
                    {
                        result = ConstCommon.TEST_FAIL;
                        sErrorMessage = "aeCell.GetClickablePoint()" + ex.Message;
                        Console.WriteLine("aeCell.GetClickablePoint()" + ex.Message);
                        Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "---" + ex.StackTrace, sOnlyUITest);
                        return;
                    }

                    string ModeValue = AUICommon.GetDataGridViewCellValueAt(2, "Mode", aeGrid);
                    if (ModeValue == null || ModeValue == string.Empty)
                    {
                        sErrorMessage = "StationDataGridView aeCell Mode Value not found:" + "Mode Row 2";
                        Console.WriteLine("StationDataGridView aeCell Mode Value not found:" + "Mode Row 2");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "StationDataGridView cell Mode value not found:" + "Mode Row 2", sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }

                    Input.MoveToAndRightClick(point);    // Mode menuitem will open
                    Thread.Sleep(2000);

                    // find Mode point
                    System.Windows.Automation.Condition cMode = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, Constants.AGV_MENUITEM_Mode),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                   );

                    // Find the MenuItemMode element
                    AutomationElement aeMenuItemMode = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cMode);
                    if (aeMenuItemMode != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemMode.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemMode.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemMode.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemMode.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemMode);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "mode not found ------------:";
                        Console.WriteLine(sErrorMessage);
                        return;
                    }

                    Thread.Sleep(2000);
                    // find Manual point
                    System.Windows.Automation.Condition cManual = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, testMode),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                   );

                    // Find the MenuItemMode element
                    AutomationElement aeMenuItemManual = aeMenuItemMode.FindFirst(TreeScope.Element | TreeScope.Descendants, cManual);
                    if (aeMenuItemManual != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemManual.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemManual.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemManual.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemManual.Current.ControlType.ProgrammaticName);
                        Console.WriteLine("new element x: " + TestTools.AUIUtilities.GetElementCenterPoint(aeMenuItemManual).X);
                        Console.WriteLine("new element Y: " + TestTools.AUIUtilities.GetElementCenterPoint(aeMenuItemManual).Y);

                        Input.MoveTo(new Point(AUIUtilities.GetElementCenterPoint(aeMenuItemManual).X, AUIUtilities.GetElementCenterPoint(aeMenuItemManual).Y));
                        Thread.Sleep(2000);
                        Input.ClickAtPoint(new Point(AUIUtilities.GetElementCenterPoint(aeMenuItemManual).X, AUIUtilities.GetElementCenterPoint(aeMenuItemManual).Y));

                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = testMode + " not found ------------:";
                        Console.WriteLine(sErrorMessage);
                        return;
                    }

                    Thread.Sleep(2000);

                    // Find  Confirm Loc state change Dialog Window
                    //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                    System.Windows.Automation.Condition c2 = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Put Stations in mode "+testMode+"?"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                   );

                    // Find the Manual Location Dialog element
                    AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                    if (aeDialog == null)
                    {
                        result = ConstCommon.TEST_FAIL;
                        sErrorMessage = " Dialog Window not found";
                        Console.WriteLine("FindElementByID failed: Dialog");
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        return;
                    }

                    // Find Yes Button
                    System.Windows.Automation.Condition c3 = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                   );

                    // Find Yes Button element
                    AutomationElement aeYes = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c3);
                    if (aeYes == null)
                    {
                        result = ConstCommon.TEST_FAIL;
                        sErrorMessage = " Yes Button not found";
                        Console.WriteLine("FindElementByID failed: Yes button");
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        return;
                    }

                    Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeYes));
                    Thread.Sleep(2000);
                    string StateValue = AUICommon.GetDataGridViewCellValueAt(2, "Mode", aeGrid);

                    sStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - sStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalSeconds);
                    while (!StateValue.Equals("Manual") && mTime.TotalSeconds < 30)
                    {
                        Thread.Sleep(2000);
                        mTime = DateTime.Now - sStartTime;
                        StateValue = AUICommon.GetDataGridViewCellValueAt(2, "Mode", aeGrid);
                        Console.WriteLine("time is (sec) : " + mTime.Seconds + " and mode is " + StateValue);
                    }

                    if (StateValue.Equals(testMode))
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                        Console.WriteLine(testname + " ---pass --- " + StateValue);
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Console.WriteLine(testname + " ---fail --- " + StateValue);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    }
                    #endregion
                }



                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine("\nTest scenario" + testname + " : Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region TransportOverviewDisplay
        public static void TransportOverviewDisplay(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            AutomationElement aePanelLink = null;
            try
            {
                AUICommon.ClearDisplayedScreens(root);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    //Input.MoveToAndClick(aePanelLink);
                    // expand transport node
                    //AUIUtilities.TreeViewNodeExpandCollapseState(aePanelLink, ExpandCollapseState.Expanded);
                    Thread.Sleep(3000);
                }

                /*System.Windows.Forms.TreeNode treeNode = new TreeNode();
                AutomationElement aeNode = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                if (aeNode == null)
                {
                    Input.MoveToAndDoubleClick(aePanelLink.GetClickablePoint());
                    Thread.Sleep(9000);
                    aeNode = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                    //Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    //result = ConstCommon.TEST_FAIL;
                    //return;
                }
                //else
                //    Input.MoveToAndClick(aeNode);

                if (aeNode == null)
                {
                    Console.WriteLine("Node not exist:" + TRANSPORT_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Input.MoveToAndClick(aeNode);
                */
                Input.MoveToAndClick(aePanelLink);
                Thread.Sleep(10000);

                // Find Transport Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, TRANSPORT_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the TransportOverview element.
                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + TRANSPORT_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_PASS;
                    string ms = TRANSPORT_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }


        }
        #endregion TransportOverviewDisplay
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region MultiLanguageCheck
        public static void MultiLanguageCheck(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 8, 29), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                string epiaDataResourceFolder = OSVersionInfoClass.ProgramFilesx86() + "\\Dematic\\Epia Server\\Data\\Resources";
                //string resourceFileName = "Epia.Modules.RnD_cn.resources";
                string[] resourceFileNames = { "Etricc.Global_cn.resources","Etricc.Global_fr.resources",
                                                 "Etricc.Global_nl.resources","Etricc.Global_de.resources",
                                                 "Etricc.Global_es.resources","Etricc.Global_en.resources"};


                for (int i = 0; i < resourceFileNames.Length; i++)
                {
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        if (EtriccUtilities.SwitchLanguageAndFindText(epiaDataResourceFolder, resourceFileNames[i], ref sErrorMessage))
                            Epia3Common.WriteTestLogMsg(slogFilePath, resourceFileNames[i] + " OK", sOnlyUITest);
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            Console.WriteLine(sErrorMessage);
                        }
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
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
        #region AgvOverviewOpenDetail
        public static void AgvOverviewOpenDetail(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
           
            AutomationElement aePanelLink = null;
            AutomationElement aeSelectedWindow = null;
            try
            {
                if (sEtriccServerStartupOK == false)
                {
                    sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                    return;
                }

                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                if (root != null)
                {
                    root.SetFocus();
                    AUICommon.ClearDisplayedScreens(root, 2);
                    aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Point Pnt = AUIUtilities.GetElementCenterPoint(aePanelLink);
                        Input.MoveToAndClick(Pnt);
                        Thread.Sleep(5000);
                    }
                }
                else
                {
                    sErrorMessage = "aeWindow not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                AutomationElement aeGrid = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(AGV_OVERVIEW_TITLE, ref sErrorMessage);
                    if (aeSelectedWindow == null)
                    {
                        Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Find AGV GridView
                        aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeSelectedWindow);
                        if (aeGrid == null)
                        {
                            sErrorMessage = "Find AgvDataGrid failed:" + AGV_GRIDDATA_ID;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }


                string AgvValue = string.Empty;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row 0";
                    // Get the Element with the Row Col Coordinates
                    AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                    if (aeCell == null)
                    {
                        sErrorMessage = "Find AgvDataGridView aeCell failed:" + cellname;
                        Console.WriteLine("Find AgvDataGridView aeCell failed:" + cellname);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("cell AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        // find cell value
                        try
                        {
                            ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            AgvValue = vp.Current.Value;
                            Console.WriteLine("Get element.Current Value:" + AgvValue);
                        }
                        catch (System.NullReferenceException)
                        {
                            AgvValue = string.Empty;
                        }

                        if (AgvValue == null || AgvValue == string.Empty)
                        {
                            sErrorMessage = "AgvDataGrid aeCell Value not found:" + cellname;
                            Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellname);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                            Thread.Sleep(3000);
                            Input.MoveToAndDoubleClick(point);
                            Thread.Sleep(3000);
                        }
                    }
                }

                // Check AGV Detail Screen Opened
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    string detailsWindowName = "Agv detail - " + AgvValue;
                    AutomationElement aeSelectedDetailsWindow = ProjBasicUI.GetSelectedOverviewWindow(detailsWindowName, ref sErrorMessage);
                    int k = 0;
                    while (aeSelectedDetailsWindow == null && k < 5)
                    {
                        Console.WriteLine("wait until selected detail window open :" + k++);
                        aeSelectedDetailsWindow = ProjBasicUI.GetSelectedOverviewWindow(detailsWindowName, ref sErrorMessage);
                        Thread.Sleep(5000);
                    }

                    if (aeSelectedDetailsWindow == null)
                    {
                        Console.WriteLine("Selected detail Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Check AGV text value
                        string textID = "m_IdValueLabel";
                        AutomationElement aeAgvText = AUIUtilities.FindElementByID(textID, aeSelectedDetailsWindow);
                        if (aeAgvText == null)
                        {
                            sErrorMessage = "Find AgvTextElement failed:" + textID;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            string agvTextValue = aeAgvText.Current.Name;
                            if (agvTextValue.Equals(AgvValue))
                            {
                                TestCheck = ConstCommon.TEST_PASS;
                            }
                            else
                            {
                                sErrorMessage = detailsWindowName + "Agv Value should be " + AgvValue + ", but  " + agvTextValue;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine("\nTest scenario " + testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LocationOverviewOpenDetail
        public static void LocationOverviewOpenDetail(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            try
            {
                AUICommon.ClearDisplayedScreens(root);

                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, LOCATION_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Input.MoveToAndClick(aePanelLink);

                Thread.Sleep(10000);

                // Find LOCATION Overview Window Element
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, LOCATION_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );

                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = LOCATION_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + LOCATION_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }
                // Find Location GridView
                AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                if (aeGrid == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine("Find LocationDataGridView failed:" + "LocationDataGridView");
                    Epia3Common.WriteTestLogFail(slogFilePath, "Find LocationDataGridView failed:" + "LocationDataGridView", sOnlyUITest);
                    return;
                }
                else
                {
                    Console.WriteLine("Button LOCATIONDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    if (sProjectFile.IndexOf("TestProject") >= 0)
                    {
                        sErrorMessage = "testproject too large, skip open location detailed window: ";
                        Console.WriteLine(sErrorMessage);
                        Thread.Sleep(2000);
                        result = ConstCommon.TEST_PASS;
                        Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
                        return;
                    }
                }

                Thread.Sleep(3000);
                // Construct the Grid Cell Element Name
                string cellname = "Id Row 0";
                // Get the Element with the Row Col Coordinates
                AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                if (aeCell == null)
                {
                    Console.WriteLine("Find LocationDataGridView aeCell failed:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find LocationDataGridView cell failed:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("cell LocationDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                Thread.Sleep(3000);
                //System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);

                System.Windows.Point point = new Point();
                try
                {
                    point = aeCell.GetClickablePoint();
                    //point = AUIUtilities.GetElementCenterPoint(aeCell);
                }
                catch (Exception ex)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "aeCell.GetClickablePoint()" + ex.Message;
                    Console.WriteLine("aeCell.GetClickablePoint()" + ex.Message);
                    Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "---" + ex.StackTrace, sOnlyUITest);
                    return;
                }

                // Open Detail screen
                Input.MoveToAndDoubleClick(point);
                Thread.Sleep(2000);

                // wait extra one minute is project is TestProject.zip
                if (sProjectFile.IndexOf("TestProject") >= 0)
                {
                    sStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - sStartTime;
                    while (mTime.Seconds < 50)
                    {
                        Thread.Sleep(2000);
                        mTime = DateTime.Now - sStartTime;
                        Console.WriteLine("wait extra time(sec) : " + mTime.Seconds);
                    }
                }

                // find location detail screen
                if (sProjectFile.IndexOf("TestProject") >= 0)
                {
                    Console.WriteLine("project too large, skip find detailed window: ");
                }
                else
                {   // get cell value
                    string cellValue = AUICommon.GetDataGridViewCellValueAt(0, "Id", aeGrid);
                    Console.WriteLine("Location id: " + cellValue);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "End id: " + cellValue, sOnlyUITest);

                    string locScreenName = "Location detail - " + cellValue;
                    Console.WriteLine("locScreenName: " + locScreenName);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "locScreenName: " + locScreenName, sOnlyUITest);

                    // Find Detail Screen
                    System.Windows.Automation.Condition c2 = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, locScreenName),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                    );

                    AutomationElement aeDetailScreen = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                    if (aeDetailScreen == null)
                    {
                        result = ConstCommon.TEST_FAIL;
                        Console.WriteLine("Find LocationDetailView failed:" + "LocationDetailView");
                        Epia3Common.WriteTestLogFail(slogFilePath, "Find LocationDetailView failed:" + "LocationDetailView", sOnlyUITest);
                        return;
                    }

                    // Check Location Value
                    string textID = "m_IdValueLabel";
                    AutomationElement aeLocText = AUIUtilities.FindElementByID(textID, aeDetailScreen);
                    if (aeLocText == null)
                    {
                        sErrorMessage = "Find locTextElement failed:" + textID;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }

                    string locTextValue = aeLocText.Current.Name;
                    if (locTextValue.Equals(cellValue))
                    {
                        result = ConstCommon.TEST_PASS;
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    }
                    else
                    {
                        result = ConstCommon.TEST_FAIL;
                        sErrorMessage = locScreenName + "Loc Value should be " + cellValue + ", but  " + locTextValue;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    }
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion

        private static AutomationElement ElementFromCursor()
        {
            // Convert mouse position from System.Drawing.Point to System.Windows.Point.
            System.Windows.Point point = new System.Windows.Point(Cursor.Position.X, Cursor.Position.Y);
            AutomationElement element = AutomationElement.FromPoint(point);
            return element;
        }
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LocationModeManual
        public static void LocationModeManual(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            try
            {
                AUICommon.ClearDisplayedScreens(root);

                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, LOCATION_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Input.MoveToAndClick(aePanelLink);

                Thread.Sleep(10000);

                // Find LOATION Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, LOCATION_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the LOCATIONOverview element.
                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = LOCATION_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + LOCATION_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }


                AutomationElement aeGrid = AUIUtilities.FindElementByID(DATAGRIDVIEW_ID, aeOverview);
                if (aeGrid == null)
                {
                    Console.WriteLine("Find LocationDataGridView failed:" + DATAGRIDVIEW_ID);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView failed:" + DATAGRIDVIEW_ID, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("LocationDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    if (sProjectFile.IndexOf("TestProject") >= 0)
                    {
                        sErrorMessage = "testproject location too large, skip change location mode: ";
                        Console.WriteLine(sErrorMessage);
                        Thread.Sleep(2000);
                        result = ConstCommon.TEST_PASS;
                        Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
                        return;
                    }
                }

                Thread.Sleep(3000);
                // Construct the Grid Cell Element Name
                string cellname = "Id Row 2";
                // Get the Element with the Row Col Coordinates
                AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                if (aeCell == null)
                {
                    Console.WriteLine("Find LocationDataGridView aeCell failed:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find LocationDataGridView cell failed:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("cell LocationDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                // find cell value
                string LocValue = string.Empty;
                try
                {
                    ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    LocValue = vp.Current.Value;
                    Console.WriteLine("Get element.Current Value:" + LocValue);
                }
                catch (System.NullReferenceException)
                {
                    LocValue = string.Empty;
                }

                if (LocValue == null || LocValue == string.Empty)
                {
                    sErrorMessage = "LocDataGridView aeCell Value not found:" + cellname;
                    Console.WriteLine("LocDataGridView aeCell Value not found:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "LocDataGridView cell value not found:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Thread.Sleep(3000);
                //System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                System.Windows.Point point = new Point();
                try
                {
                    point = aeCell.GetClickablePoint();
                    //point = AUIUtilities.GetElementCenterPoint(aeCell);
                }
                catch (Exception ex)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "aeCell.GetClickablePoint()" + ex.Message;
                    Console.WriteLine("aeCell.GetClickablePoint()" + ex.Message);
                    Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "---" + ex.StackTrace, sOnlyUITest);
                    return;
                }

                string ModeValue = AUICommon.GetDataGridViewCellValueAt(2, "Mode", aeGrid);
                if (ModeValue == null || ModeValue == string.Empty)
                {
                    sErrorMessage = "LocDataGridView aeCell Mode Value not found:" + "Mode Row 2";
                    Console.WriteLine("LocDataGridView aeCell Mode Value not found:" + "Mode Row 2");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "LocDataGridView cell Mode value not found:" + "Mode Row 2", sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Input.MoveToAndRightClick(point);    // Mode menuitem will open
                Thread.Sleep(2000);

                // find Mode point
                System.Windows.Automation.Condition cMode = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, Constants.AGV_MENUITEM_Mode),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
               );

                // Find the MenuItemMode element
                AutomationElement aeMenuItemMode = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cMode);
                if (aeMenuItemMode != null)
                {
                    Console.WriteLine("new element found: " + aeMenuItemMode.Current.Name);
                    Console.WriteLine("err.Current.AutomationId: " + aeMenuItemMode.Current.AutomationId);
                    Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemMode.Current.ControlType.ToString());
                    Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemMode.Current.ControlType.ProgrammaticName);
                    Input.MoveToAndClick(aeMenuItemMode);
                }
                else
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = "mode not found ------------:";
                    Console.WriteLine(sErrorMessage);
                    return;
                }

                Thread.Sleep(2000);
                // find Manual point
                System.Windows.Automation.Condition cManual = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, Constants.LOCATION_MENUITEM_Manual),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
               );

                // Find the MenuItemMode element
                AutomationElement aeMenuItemManual = aeMenuItemMode.FindFirst(TreeScope.Element | TreeScope.Descendants, cManual);
                if (aeMenuItemManual != null)
                {
                    Console.WriteLine("new element found: " + aeMenuItemManual.Current.Name);
                    Console.WriteLine("err.Current.AutomationId: " + aeMenuItemManual.Current.AutomationId);
                    Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemManual.Current.ControlType.ToString());
                    Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemManual.Current.ControlType.ProgrammaticName);
                    Console.WriteLine("new element x: " + TestTools.AUIUtilities.GetElementCenterPoint(aeMenuItemManual).X);
                    Console.WriteLine("new element Y: " + TestTools.AUIUtilities.GetElementCenterPoint(aeMenuItemManual).Y);

                    Input.MoveTo(new Point(AUIUtilities.GetElementCenterPoint(aeMenuItemManual).X, AUIUtilities.GetElementCenterPoint(aeMenuItemManual).Y));
                    Thread.Sleep(2000);
                    Input.ClickAtPoint(new Point(AUIUtilities.GetElementCenterPoint(aeMenuItemManual).X, AUIUtilities.GetElementCenterPoint(aeMenuItemManual).Y));

                }
                else
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = "manual not found ------------:";
                    Console.WriteLine(sErrorMessage);
                    return;
                }

                Thread.Sleep(2000);

                // Find  Confirm Loc state change Dialog Window
                //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                System.Windows.Automation.Condition c2 = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, "Put Locations in mode Manual?"),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the Manual Location Dialog element
                AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                if (aeDialog == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = " Dialog Window not found";
                    Console.WriteLine("FindElementByID failed: Dialog");
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                // Find Yes Button
                System.Windows.Automation.Condition c3 = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
               );

                // Find Yes Button element
                AutomationElement aeYes = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c3);
                if (aeYes == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = " Yes Button not found";
                    Console.WriteLine("FindElementByID failed: Yes button");
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeYes));
                Thread.Sleep(2000);
                string StateValue = AUICommon.GetDataGridViewCellValueAt(2, "Mode", aeGrid);

                sStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - sStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);
                while (!StateValue.Equals("Manual") && mTime.TotalSeconds < 30)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - sStartTime;
                    StateValue = AUICommon.GetDataGridViewCellValueAt(2, "Mode", aeGrid);
                    Console.WriteLine("time is (sec) : " + mTime.Seconds + " and mode is " + StateValue);
                }

                if (StateValue.Equals("Manual"))
                {
                    result = ConstCommon.TEST_PASS;
                    Console.WriteLine(testname + " ---pass --- " + StateValue);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine(testname + " ---fail --- " + StateValue);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region AgvActionMenuitemsValidation
        public static void AgvActionMenuitemsValidation(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            AutomationElement aeSelectedWindow = null;
            string StateValue = string.Empty;
            Point Pnt = new System.Windows.Point();
            try
            {
                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                if (root != null)
                {
                    root.SetFocus();
                    AUICommon.ClearDisplayedScreens(root,2);
                    aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Pnt = AUIUtilities.GetElementCenterPoint(aePanelLink);
                        Input.MoveToAndClick(Pnt);
                        Thread.Sleep(5000);
                    }
                }
                else
                {
                    sErrorMessage = "aeWindow not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                AutomationElement aeGrid = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(AGV_OVERVIEW_TITLE, ref sErrorMessage);
                    if (aeSelectedWindow == null)
                    {
                        Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Find AGV GridView
                        aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeSelectedWindow);
                        if (aeGrid == null)
                        {
                            sErrorMessage = "Find AgvDataGrid failed:" + AGV_GRIDDATA_ID;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row 0";
                    // Get the Element with the Row Col Coordinates
                    AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                    if (aeCell == null)
                    {
                        sErrorMessage = "Find AgvDataGridView aeCell failed:" + cellname;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // find cell value
                        string AgvValue = string.Empty;
                        try
                        {
                            ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            AgvValue = vp.Current.Value;
                            Console.WriteLine("Get element.Current Value:" + AgvValue);
                        }
                        catch (System.NullReferenceException)
                        {
                            AgvValue = string.Empty;
                        }

                        if (AgvValue == null || AgvValue == string.Empty)
                        {
                            sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellname;
                            Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellname);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Thread.Sleep(3000);
                            System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                            //
                            // get State Value 
                            //
                            string cellState = "State Row 0";
                            // Get the Element with the Row Col Coordinates
                            AutomationElement aeCellState = AUIUtilities.FindElementByName(cellState, aeGrid);
                            if (aeCellState == null)
                            {
                                sErrorMessage = "Find AgvDataGridView aeCellState failed:" + cellState;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                Console.WriteLine("cellState AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                // find cellState value
                                try
                                {
                                    ValuePattern vp1 = aeCellState.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                    StateValue = vp1.Current.Value;
                                    Console.WriteLine("Get element.Current Value:" + StateValue);
                                }
                                catch (System.NullReferenceException)
                                {
                                    StateValue = string.Empty;
                                }

                                if (StateValue == null || StateValue == string.Empty)
                                {
                                    sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellState;
                                    Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellState);
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellState, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                                else
                                {
                                    System.Windows.Point point2 = AUIUtilities.GetElementCenterPoint(aeCellState);
                                    Thread.Sleep(3000);
                                    Input.MoveToAndRightClick(point);
                                    Console.WriteLine("MoveToAndRightClick(point) --------------------------------------:" );
                                    Thread.Sleep(3000);
                                }
                            }
                        }
                    }
                }

                bool status =true;
                bool sinfo = false;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    /*string[] AgvAllItems = new string[] { "Restart Agv", "Retire Agv", "Agv detail", "New Job", "New Week Plan", 
                        "Stop Agv", "Suspend Agv", "Cancel current Batch", "Jobs", "Battery charge plan", 
                        "Mode", "Automatic", "Disabled", "Semi-Automatic", "Removed",
                        "Engineering", "Download", "Simulation", "Save Areas", "Logging",
                            "Standard",
                            "Vehicle", "Info", "Warning", "Error", "Debug",
                            "Protocol", "Info", "Warning", "Error", "Debug",
                            "Communication", "Info", "Warning", "Error", "Debug",
                            "Test Sequence", "Start", "Stop", "Pauze",
                
                    };*/

                    // new version
                    string[] AgvAllItems = new string[] { "Restart Agv", "Retire Agv", "Agv detail", "New Job", "New Week Plan",
                         "Shutdown Agv", "Stop Agv", "Suspend Agv", "Cancel current Batch", "Jobs", "Battery charge plan",
                        "Mode", "Automatic", "Disabled", "Semi-Automatic", "Removed",
                        "Engineering", "Download", "Simulation", "Save Areas", "Logging",
                            "Standard",
                            "Vehicle", "Info", "Warning", "Error", "Debug",
                            "Protocol", "Info", "Warning", "Error", "Debug",
                            "Communication", "Info", "Warning", "Error", "Debug",
                            "Test Sequence", "Start", "Stop", "Pauze",
                        "Copy"
                    };

                    string[] VisibleItems = new string[] { "Restart Agv", "Retire Agv", "Agv detail", "New Job",  "New Week Plan",
                        "Shutdown Agv", "Stop Agv", "Suspend Agv", "Cancel current Batch", "Jobs", "Battery charge plan", "Mode", "Engineering","Copy"
                    };

                    if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2013, 10, 15), ref sErrorMessage) == true)
                    {
                        sErrorMessage = "Date before 2013oct15, TestSequence menu item not tested";
                        Console.WriteLine("Visible1 --------" + sErrorMessage);
                        sinfo = true;
                    }
                    else
                    {
                        status = ProjBasicUI.ValidateWindowMenuItemActionElement("MainForm", VisibleItems, ref sErrorMessage, "Visible");
                        if (status == false)
                        {
                            sErrorMessage = "AgvVisibleActionMenuitems:" + sErrorMessage;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2013, 10, 15), ref sErrorMessage) == true)
                            {
                                sErrorMessage = "Date before 2013oct15, TestSequence menu item not tested";
                                
                                Console.WriteLine("Visible2 --------" + sErrorMessage);
                                sinfo = true;
                            }
                            else
                            {
                                Console.WriteLine("Visible3 ValidateWindowMenuItemActionElement All");
                                status = ProjBasicUI.ValidateWindowMenuItemActionElement("MainForm", AgvAllItems, ref sErrorMessage, "All");
                                if (status == false)
                                {
                                    sErrorMessage = "AgvAllActionMenuitems:" + sErrorMessage;
                                    Console.WriteLine(sErrorMessage);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                            }
                        }
                    }
                }

                Thread.Sleep(1000);
                Input.MoveToAndClick(Pnt);
                Thread.Sleep(1000);

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    if (sinfo == false )
                        Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine("\nTest scenario" + testname + " : Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                sErrorMessage = sErrorMessage.Trim();
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region RestartAgv
        public static void RestartAgv(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            AutomationElement aeSelectedWindow = null;
            string StateValue = string.Empty;
            try
            {
                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                if (root != null)
                {
                    root.SetFocus();
                    AUICommon.ClearDisplayedScreens(root,2);
                    aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Point Pnt = AUIUtilities.GetElementCenterPoint(aePanelLink);
                        Input.MoveToAndClick(Pnt);
                        Thread.Sleep(5000);
                    }
                }
                else
                {
                    sErrorMessage = "aeWindow not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                AutomationElement aeGrid = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(AGV_OVERVIEW_TITLE, ref sErrorMessage);
                    if (aeSelectedWindow == null)
                    {
                        Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Find AGV GridView
                        aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeSelectedWindow);
                        if (aeGrid == null)
                        {
                            sErrorMessage = "Find AgvDataGrid failed:" + AGV_GRIDDATA_ID;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row 0";
                    // Get the Element with the Row Col Coordinates
                    AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                    if (aeCell == null)
                    {
                        sErrorMessage = "Find AgvDataGridView aeCell failed:" + cellname;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // find cell value
                        string AgvValue = string.Empty;
                        try
                        {
                            ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            AgvValue = vp.Current.Value;
                            Console.WriteLine("Get element.Current Value:" + AgvValue);
                        }
                        catch (System.NullReferenceException)
                        {
                            AgvValue = string.Empty;
                        }

                        if (AgvValue == null || AgvValue == string.Empty)
                        {
                            sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellname;
                            Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellname);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Thread.Sleep(3000);
                            System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                            //
                            // get State Value 
                            //
                            string cellState = "State Row 0";
                            // Get the Element with the Row Col Coordinates
                            AutomationElement aeCellState = AUIUtilities.FindElementByName(cellState, aeGrid);
                            if (aeCellState == null)
                            {
                                sErrorMessage = "Find AgvDataGridView aeCellState failed:" + cellState;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                Console.WriteLine("cellState AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                // find cellState value
                                try
                                {
                                    ValuePattern vp1 = aeCellState.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                    StateValue = vp1.Current.Value;
                                    Console.WriteLine("Get element.Current Value:" + StateValue);
                                }
                                catch (System.NullReferenceException)
                                {
                                    StateValue = string.Empty;
                                }

                                if (StateValue == null || StateValue == string.Empty)
                                {
                                    sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellState;
                                    Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellState);
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellState, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                                else
                                {
                                    System.Windows.Point point2 = AUIUtilities.GetElementCenterPoint(aeCellState);
                                    Thread.Sleep(3000);
                                    Input.MoveToAndRightClick(point);
                                    Thread.Sleep(3000);
                                }
                            }
                        }
                    }
                }

                AutomationElement aeMenuItemRestartAgv =null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeMenuItemRestartAgv = ProjBasicUI.GetWindowMenuItemActionElement("MainForm", Constants.AGV_MENUITEM_Restart_Agv, ref sErrorMessage);
                    if (aeMenuItemRestartAgv == null)
                    {
                        sErrorMessage = "Restart Agv menu iteme not found ------------:";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                        Input.MoveToAndClick(aeMenuItemRestartAgv);
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(2000);
                    // Find the ARestart Agv Dialog element
                    AutomationElement aeDialog = ProjBasicUI.GetPopupDialogFromMainWindow("MainForm", "Restart Agvs?", "name", ref sErrorMessage);
                    if (aeDialog == null)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        if (ProjBasicUI.ClickButtonInThisElement("Yes", "name", aeDialog, ref sErrorMessage) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }  
                    }
                }

                // validate state value
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    StateValue = AUICommon.GetDataGridViewCellValueAt(0, "State", aeGrid);
                    sStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - sStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalSeconds);
                    while (!StateValue.Equals("Ready") && mTime.TotalSeconds < 30)
                    {
                        Thread.Sleep(2000);
                        mTime = DateTime.Now - sStartTime;
                        StateValue = AUICommon.GetDataGridViewCellValueAt(0, "State", aeGrid);
                        Console.WriteLine("time is (sec) : " + mTime.Seconds + " and state is " + StateValue);
                    }

                    if (!StateValue.Equals("Ready"))
                    {
                        sErrorMessage = "Agv State value is not equal to Ready";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine("\nTest scenario" + testname + " : Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                sErrorMessage = sErrorMessage.Trim();
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region AgvJobOverview
        public static void AgvJobOverview(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            try
            {
                AUICommon.ClearDisplayedScreens(root);

                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Input.MoveToAndClick(aePanelLink);
                    Thread.Sleep(2000);
                    Input.MoveToAndClick(aePanelLink);
                }

                Thread.Sleep(10000);

                // Find AGV Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, AGV_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the AGVOverview element.
                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = AGV_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + AGV_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
                if (aeGrid == null)
                {
                    sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Console.WriteLine("AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

                Thread.Sleep(3000);
                // Construct the Grid Cell Element Name
                string cellname = "Id Row 1";
                // Get the Element with the Row Col Coordinates
                AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                if (aeCell == null)
                {
                    Console.WriteLine("Find AgvDataGridView aeCell failed:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("cell AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                Thread.Sleep(3000);
                // find cell value
                string AgvValue = string.Empty;
                try
                {
                    ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    AgvValue = vp.Current.Value;
                    Console.WriteLine("Get element.Current Value:" + AgvValue);
                }
                catch (System.NullReferenceException)
                {
                    AgvValue = string.Empty;
                }

                if (AgvValue == null || AgvValue == string.Empty)
                {
                    sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellname;
                    Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Thread.Sleep(3000);
                System.Windows.Point point = new Point();
                try
                {
                    point = aeCell.GetClickablePoint();
                }
                catch (Exception ex)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "aeCell.GetClickablePoint()" + ex.Message;
                    Console.WriteLine("aeCell.GetClickablePoint()" + ex.Message);
                    return;
                }

                //string AgvValue = AUICommon.GetDataGridViewCellValueAt(0, "Id", aeGrid);
                Input.MoveToAndRightClick(point);
                Thread.Sleep(3000);

                // find Jobs point
                System.Windows.Automation.Condition cJobs = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, Constants.AGV_MENUITEM_Jobs),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
               );

                // Find the MenuItem Job element
                AutomationElement aeMenuItemJobs = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cJobs);
                if (aeMenuItemJobs != null)
                {
                    Console.WriteLine("new element found: " + aeMenuItemJobs.Current.Name);
                    Console.WriteLine("err.Current.AutomationId: " + aeMenuItemJobs.Current.AutomationId);
                    Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemJobs.Current.ControlType.ToString());
                    Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemJobs.Current.ControlType.ProgrammaticName);
                    Input.MoveToAndClick(aeMenuItemJobs);
                }
                else
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = "mode not found ------------:";
                    Console.WriteLine(sErrorMessage);
                    return;
                }

                Thread.Sleep(2000);

                // Find Agv Job Overview Window
                string JobsWindowID = "Jobs - " + AgvValue;
                // Find AGV Overview Window
                System.Windows.Automation.Condition c2 = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, JobsWindowID),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the JobsOverview element.
                AutomationElement aeJobsOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                if (aeJobsOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = JobsWindowID + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + JobsWindowID);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }
                else
                {
                    result = ConstCommon.TEST_PASS;
                    Console.WriteLine(testname + " ---pass --- ");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region AgvJobOverviewOpenDetail
        public static void AgvJobOverviewOpenDetail(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            try
            {
                AUICommon.ClearDisplayedScreens(root);

                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Input.MoveToAndClick(aePanelLink);
                    Thread.Sleep(2000);
                    Input.MoveToAndClick(aePanelLink);
                }

                Thread.Sleep(10000);

                // Find AGV Overview Window Element
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, AGV_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );

                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = AGV_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + AGV_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
                if (aeGrid == null)
                {
                    sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Console.WriteLine("AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

                Thread.Sleep(3000);

                //sProjectFile = "xxeurobaltic";
                int row = 0;  // for demo or testProject 
                if (sProjectFile.ToLower().IndexOf("eurobaltic") >= 0)
                    row = 4;

                string cellname = "Id Row " + row;
                // Get the Element with the Row Col Coordinates
                AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                if (aeCell == null)
                {
                    Console.WriteLine("Find AgvDataGridView aeCell failed:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("Agv cell  found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                Thread.Sleep(3000);
                // find cell value
                string AgvValue = string.Empty;
                try
                {
                    ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    AgvValue = vp.Current.Value;
                    Console.WriteLine("Get element.Current Value:" + AgvValue);
                }
                catch (System.NullReferenceException)
                {
                    AgvValue = string.Empty;
                }

                if (AgvValue == null || AgvValue == string.Empty)
                {
                    sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellname;
                    Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Thread.Sleep(3000);
                System.Windows.Point point = new Point();
                try
                {
                    point = aeCell.GetClickablePoint();
                }
                catch (Exception ex)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "aeCell.GetClickablePoint()" + ex.Message;
                    Console.WriteLine("aeCell.GetClickablePoint()" + ex.Message);
                    return;
                }

                //System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                //System.Windows.Point point = AUICommon.GetDataGridViewCellPointAt(0, "Id", aeGrid);


                //string AgvValue = AUICommon.GetDataGridViewCellValueAt(0, "Id", aeGrid);
                Input.MoveToAndRightClick(point);
                Thread.Sleep(3000);

                // find Jobs point
                System.Windows.Automation.Condition cJobs = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, Constants.AGV_MENUITEM_Jobs),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
               );

                // Find the MenuItem Job element
                AutomationElement aeMenuItemJobs = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cJobs);
                if (aeMenuItemJobs != null)
                {
                    Console.WriteLine("new element found: " + aeMenuItemJobs.Current.Name);
                    Console.WriteLine("err.Current.AutomationId: " + aeMenuItemJobs.Current.AutomationId);
                    Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemJobs.Current.ControlType.ToString());
                    Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemJobs.Current.ControlType.ProgrammaticName);
                    Input.MoveToAndClick(aeMenuItemJobs);
                }
                else
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = "mode not found ------------:";
                    Console.WriteLine(sErrorMessage);
                    return;
                }

                Thread.Sleep(2000);

                // Find Agv Job Overview Window
                string JobsWindowID = "Jobs - " + AgvValue;
                // Find AGV Overview Window
                System.Windows.Automation.Condition c2 = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, JobsWindowID),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the JobsOverview element.
                AutomationElement aeJobsOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                if (aeJobsOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = JobsWindowID + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + JobsWindowID);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                // Find Job Detail screen        
                AutomationElement aeJobGrid = AUIUtilities.FindElementByID(DATAGRIDVIEW_ID, aeJobsOverview);
                if (aeJobGrid == null)
                {
                    Console.WriteLine("Find JobDataGridView failed:" + DATAGRIDVIEW_ID);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find JobDataGridView failed:" + "JobDataGridView", sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Console.WriteLine("JobDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

                string JobCellname = "Id Row " + row;
                AutomationElement aeJobCell = AUIUtilities.FindElementByName(JobCellname, aeJobGrid);
                if (aeJobCell == null)
                {
                    sErrorMessage = "Find JobDataGridView aeCell failed:" + JobCellname;
                    Console.WriteLine("Find JobDataGridView aeCell failed:" + JobCellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find JobDataGridView cell failed:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("Job cell  found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                Point Point = AUIUtilities.GetElementCenterPoint(aeJobCell);
                Input.MoveTo(Point);
                Thread.Sleep(1000);

                string JobValue = null;
                try
                {
                    ValuePattern vp = aeJobCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    JobValue = vp.Current.Value;
                    Console.WriteLine("Get element.Current Value:" + vp.Current.Value);
                }
                catch (System.NullReferenceException)
                {
                    JobValue = null;
                    sErrorMessage = "Find aeJobCell value failed:" + JobCellname;
                    Console.WriteLine("Find aeJobCell value failed:" + JobCellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find aeJobCell value failed:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }


                Console.WriteLine("Job Id value: " + JobValue);
                Epia3Common.WriteTestLogMsg(slogFilePath, "Job Id cell value: " + JobValue, sOnlyUITest);

                Thread.Sleep(3000);
                System.Windows.Point JobPoint = AUIUtilities.GetElementCenterPoint(aeJobCell);
                Input.MoveToAndDoubleClick(JobPoint);

                Thread.Sleep(3000);

                //string JobValue = "JOB1";
                string JobScreenName = "Job detail - " + AgvValue + " - " + JobValue;
                Console.WriteLine("JobScreenName: " + JobScreenName);
                Epia3Common.WriteTestLogMsg(slogFilePath, "JobScreenName: " + JobScreenName, sOnlyUITest);

                // Find All Windows
                System.Windows.Automation.Condition c3 = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.Window);

                // Find All Windows.
                AutomationElementCollection aeAllWindows = root.FindAll(TreeScope.Element | TreeScope.Descendants, c3);
                Thread.Sleep(3000);

                AutomationElement aeDetailScreen = null;
                bool isOpen = false;
                string windowName = "No Name";
                for (int i = 0; i < aeAllWindows.Count; i++)
                {
                    windowName = aeAllWindows[i].Current.Name;
                    Console.WriteLine("Window name: " + windowName);
                    Console.WriteLine("JobScreenName: " + JobScreenName);
                    if (windowName.StartsWith(JobScreenName))
                    {
                        isOpen = true;
                        aeDetailScreen = aeAllWindows[i];
                        break;
                    }
                }

                if (!isOpen)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Find JobDetailView failed:" + JobScreenName;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    return;
                }

                Thread.Sleep(3000);
                // Check Job ID Value
                string textID = "m_IdValueLabel";
                AutomationElement aeJobText = AUIUtilities.FindElementByID(textID, aeDetailScreen);
                if (aeJobText == null)
                {
                    sErrorMessage = "Find JobTextElement failed:" + textID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                string JobTextValue = aeJobText.Current.Name;
                if (JobTextValue.StartsWith(JobValue))
                {
                    Console.WriteLine("--pass--");
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = JobScreenName + "Job Value should be " + JobValue + ", but  " + JobTextValue;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region AgvModeSemiAutomatic
        public static void AgvModeSemiAutomatic(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            AutomationElement aeSelectedWindow = null;
            try
            {
                if (sOnlyUITest)
                    root = EtriccUtilities.GetMainWindow("MainForm");

                AUICommon.ClearDisplayedScreens(root);

                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {
                    Input.MoveToAndClick(aePanelLink);
                }

                // Find AGV Overview Window
                AutomationElement aeGrid = null;
                AutomationElement aeCell = null;
                string AgvValue = string.Empty;
                string cellname = "Id Row 0";

                int ky = 0;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(5000);
                    aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(AGV_OVERVIEW_TITLE, ref sErrorMessage);
                    while (aeSelectedWindow == null && ky < 5)
                    {
                        Console.WriteLine("wait until selected " + AGV_OVERVIEW_TITLE + " window open :" + ky++);
                        aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(AGV_OVERVIEW, ref sErrorMessage);
                        Thread.Sleep(5000);
                    }

                    if (aeSelectedWindow == null)
                    {
                        Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

               
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeSelectedWindow);
                    if (aeGrid == null)
                    {
                        sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Thread.Sleep(3000);
                        // Construct the Grid Cell Element Name
                        // Get the Element with the Row Col Coordinates
                        aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                        if (aeCell == null)
                        {
                            sErrorMessage = "Find AgvDataGridView aeCell failed:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            try
                            {
                                ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                AgvValue = vp.Current.Value;
                                Console.WriteLine("Get element.Current Value:" + AgvValue);
                            }
                            catch (System.NullReferenceException)
                            {
                                AgvValue = string.Empty;
                            }

                            if (AgvValue == null || AgvValue == string.Empty)
                            {
                                sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellname;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                    }
                }

                System.Windows.Point AgvPoint = new Point();
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    //System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                    try
                    {
                        AgvPoint = aeCell.GetClickablePoint();
                    }
                    catch (Exception ex)
                    {
                        sErrorMessage = "aeCell.GetClickablePoint()" + ex.Message;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                AutomationElement aeMenuItemMode = null;
                string[] Modes = new string[] { Constants.AGV_MENUITEM_Semi_Automatic, "Disabled", "Removed", "Automatic" };
                int k = 0;
                while (TestCheck == ConstCommon.TEST_PASS && k < 4)
                {
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        Input.MoveToAndRightClick(AgvPoint);
                        Thread.Sleep(4000);
                        // find Mode point
                        System.Windows.Automation.Condition cModeMenuItem = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, Constants.AGV_MENUITEM_Mode),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                        );

                        // Find the MenuItemMode element
                        aeMenuItemMode = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cModeMenuItem);
                        if (aeMenuItemMode != null)
                        {
                            Console.WriteLine("new element found: " + aeMenuItemMode.Current.Name);
                            Console.WriteLine("err.Current.AutomationId: " + aeMenuItemMode.Current.AutomationId);
                            Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemMode.Current.ControlType.ToString());
                            Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemMode.Current.ControlType.ProgrammaticName);
                            Input.MoveToAndClick(aeMenuItemMode);
                        }
                        else
                        {
                            sErrorMessage = "aeMenuItemMode not found ------------:";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }

                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        Thread.Sleep(2000);
                        // find click point
                        System.Windows.Automation.Condition cSemiAuto = new AndCondition(
                           new PropertyCondition(AutomationElement.NameProperty, Modes[k]),
                           new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                       );

                        // Find the MenuItemMode element
                        AutomationElement aeSemiAutomatic = aeMenuItemMode.FindFirst(TreeScope.Element | TreeScope.Descendants, cSemiAuto);
                        if (aeSemiAutomatic != null)
                        {
                            Console.WriteLine("new element found: " + aeSemiAutomatic.Current.Name);
                            Console.WriteLine("err.Current.AutomationId: " + aeSemiAutomatic.Current.AutomationId);
                            Console.WriteLine(" err.Current.ControlType.ToString(): " + aeSemiAutomatic.Current.ControlType.ToString());
                            Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeSemiAutomatic.Current.ControlType.ProgrammaticName);
                            Console.WriteLine("new element x: " + TestTools.AUIUtilities.GetElementCenterPoint(aeSemiAutomatic).X);
                            Console.WriteLine("new element Y: " + TestTools.AUIUtilities.GetElementCenterPoint(aeSemiAutomatic).Y);

                            Input.MoveTo(new Point(AUIUtilities.GetElementCenterPoint(aeSemiAutomatic).X, AUIUtilities.GetElementCenterPoint(aeSemiAutomatic).Y));
                            Thread.Sleep(2000);
                            Input.ClickAtPoint(new Point(AUIUtilities.GetElementCenterPoint(aeSemiAutomatic).X, AUIUtilities.GetElementCenterPoint(aeSemiAutomatic).Y));

                        }
                        else
                        {
                            sErrorMessage = Modes[k] + " not found ------------:";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }

                    }
                    AutomationElement aeDialog = null;
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        Thread.Sleep(2000);
                        // Find Restart Agv Dialog Window
                        System.Windows.Automation.Condition c2 = new AndCondition(
                           new PropertyCondition(AutomationElement.NameProperty, "Put Agvs in mode " + Modes[k] + "?"),
                           new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                        );

                        // Find the ARestart Agv Dialog element
                        aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);

                        if (aeDialog == null)
                        {
                            sErrorMessage = " Dialog Window not found";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            // Find Yes Button
                            System.Windows.Automation.Condition c3 = new AndCondition(
                               new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                           );

                            // Find Yes Button element
                            AutomationElement aeYes = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c3);
                            if (aeYes == null)
                            {
                                sErrorMessage = " Yes Button not found";
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeYes));
                                Thread.Sleep(2000);
                            }
                        }
                    }

                    // find cellelement FLV
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        bool validate = EtriccUtilities.ValidateGridData(aeGrid, "Id", AgvValue, "Mode", Modes[k], 2, ref sErrorMessage);
                        if (validate == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                            k++;
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    sBeforeAfterActivateScriptOK = false;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                Thread.Sleep(3000);


            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region AgvEngeeringSimulation
        public static void AgvEngeeringSimulation(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            AutomationElement aePanelLink = null;
            AutomationElement aeAgvOverview = null;
            try
            {
                 if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 3, 15), ref sErrorMessage) == true)
                 {
                     result = ConstCommon.TEST_UNDEFINED;
                     return;
                 }

                if (sOnlyUITest)
                    root = EtriccUtilities.GetMainWindow("MainForm");

                TransformPattern tranform = root.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                if (tranform != null)
                {
                    //tranform.Resize(System.Windows.Forms.SystemInformation.VirtualScreen.Width - System.Windows.Forms.SystemInformation.VirtualScreen.Width * 0.1,
                    tranform.Resize(System.Windows.Forms.SystemInformation.VirtualScreen.Width/2,
                        System.Windows.Forms.SystemInformation.VirtualScreen.Height - System.Windows.Forms.SystemInformation.VirtualScreen.Height*0.2);
                    Thread.Sleep(1000);
                    tranform.Move(0, 0);
                }

                AUICommon.ClearDisplayedScreens(root, 3);
            
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {
                    Input.MoveToAndClick(aePanelLink);
                    Thread.Sleep(10000);
                    // Find AGV Overview Window
                    System.Windows.Automation.Condition c = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, AGV_OVERVIEW_TITLE),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                   );

                    // Find the AGVOverview element.
                    aeAgvOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeAgvOverview == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = AGV_OVERVIEW + " Window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    }
                }

                AutomationElement aeGrid = null;
                AutomationElement aeCell = null;
                string AgvValue = string.Empty;
                string cellname = "Id Row 0";
               
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeAgvOverview);
                    if (aeGrid == null)
                    {
                        sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Thread.Sleep(3000);
                        // Construct the Grid Cell Element Name
                        // Get the Element with the Row Col Coordinates
                        aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                        if (aeCell == null)
                        {
                            sErrorMessage = "Find AgvDataGridView aeCell failed:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            try
                            {
                                ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                AgvValue = vp.Current.Value;
                                Console.WriteLine("Get element.Current Value:" + AgvValue);
                            }
                            catch (System.NullReferenceException)
                            {
                                AgvValue = string.Empty;
                            }

                            if (AgvValue == null || AgvValue == string.Empty)
                            {
                                sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellname;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                    }
                }

                System.Windows.Point AgvPoint = new Point();
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    //System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                    try
                    {
                        AgvPoint = aeCell.GetClickablePoint();
                    }
                    catch (Exception ex)
                    {
                        sErrorMessage = "aeCell.GetClickablePoint()" + ex.Message;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                AutomationElement aeMenuItemEngineering = null;
                #region open simulation window
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Input.MoveToAndRightClick(AgvPoint);
                    Thread.Sleep(2000);
                    // find Mode point
                    System.Windows.Automation.Condition cEngineeringMenuItem = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, Constants.AGV_MENUITEM_Engineering),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                    );

                    // Find the aeMenuItemEngineering element
                    aeMenuItemEngineering = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cEngineeringMenuItem);
                    if (aeMenuItemEngineering != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemEngineering.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemEngineering.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemEngineering.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemEngineering.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemEngineering);
                    }
                    else
                    {
                        sErrorMessage = "aeMenuItemMode not found ------------:";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(2000);
                    // find click point
                    System.Windows.Automation.Condition cSimulation = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, Constants.AGV_MENUITEM_Simulation),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                    );

                    // Find the MenuItemMode element
                    AutomationElement aeSimulation = aeMenuItemEngineering.FindFirst(TreeScope.Element | TreeScope.Descendants, cSimulation);
                    if (aeSimulation != null)
                    {
                        Console.WriteLine("new element found: " + aeSimulation.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeSimulation.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeSimulation.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeSimulation.Current.ControlType.ProgrammaticName);
                        Console.WriteLine("new element x: " + TestTools.AUIUtilities.GetElementCenterPoint(aeSimulation).X);
                        Console.WriteLine("new element Y: " + TestTools.AUIUtilities.GetElementCenterPoint(aeSimulation).Y);

                        Input.MoveTo(new Point(AUIUtilities.GetElementCenterPoint(aeSimulation).X, AUIUtilities.GetElementCenterPoint(aeSimulation).Y));
                        Thread.Sleep(2000);
                        Input.ClickAtPoint(new Point(AUIUtilities.GetElementCenterPoint(aeSimulation).X, AUIUtilities.GetElementCenterPoint(aeSimulation).Y));

                    }
                    else
                    {
                        sErrorMessage = Constants.AGV_MENUITEM_Simulation + " not found ------------:";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }

                }
                #endregion open simulation window

                //AutomationElement aeMenuItemAgvDetail = null;
                #region open and vertical shift Agv details window
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Input.MoveToAndDoubleClick(AgvPoint);
                }
                
                #region Find Agv Detail screen
                AutomationElement aeAgvDetailWindowTab = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(2000);
                    root = EtriccUtilities.GetMainWindow("MainForm");

                    Condition cAgvDetailWindow = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
                    // Find the simulation element.
                    DateTime mStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - mStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalSeconds + "    testcheck = " + TestCheck);
                    while (aeAgvDetailWindowTab == null && mTime.TotalSeconds < 60)
                    {
                        AutomationElementCollection aeAllWindowTabs = root.FindAll(TreeScope.Element | TreeScope.Descendants, cAgvDetailWindow);
                        for (int i = 0; i < aeAllWindowTabs.Count; i++)
                        {
                            Console.WriteLine("All sub window: " + aeAllWindowTabs[i].Current.Name);
                            if (aeAllWindowTabs[i].Current.Name.StartsWith("Agv detail"))
                            {
                                aeAgvDetailWindowTab = aeAllWindowTabs[i];
                                break;
                            }
                        }

                        Thread.Sleep(2000);
                        mTime = DateTime.Now - mStartTime;
                        Console.WriteLine(" time2 is :" + mTime.TotalSeconds);
                    }

                    if (aeAgvDetailWindowTab == null)
                    {
                        sErrorMessage = "aeAgvDetailWindowTab not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        string ms = " aeAgvDetailWindowTab found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                        Console.WriteLine(ms);
                        Epia3Common.WriteTestLogMsg(slogFilePath, ms, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                }
                #endregion
                #endregion open Agv detail window

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    root = EtriccUtilities.GetMainWindow("MainForm");
                    int retCode = EtriccUtilities.ValidateAgvSimulation(root, "Battery Low", ref sErrorMessage);
                    if (retCode == 1)
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                    else if (retCode == 0)
                    {
                        sErrorMessage = "Battery Low" + " Not in Active List ------------:";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else if (retCode == -1)
                    {
                        sErrorMessage = "Battery Low" + " simulation error :" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    root = EtriccUtilities.GetMainWindow("MainForm");
                    int retCode = EtriccUtilities.ValidateAgvSimulation(root, "Battery Low", ref sErrorMessage);
                    if (retCode == 1)
                    {
                        sErrorMessage = "Battery Low" + " still in Active List ------------:";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else if (retCode == 0)
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                        
                    }
                    else if (retCode == -1)
                    {
                        sErrorMessage = "Battery Low" + " simulation error :" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    sBeforeAfterActivateScriptOK = false;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region AgvsAllModeRemoved
        public static void AgvsAllModeRemoved(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            try
            {
                AUICommon.ClearDisplayedScreens(root, 2);

                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Input.MoveToAndClick(aePanelLink);

                Thread.Sleep(10000);

                // Find AGV Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, AGV_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the AGVOverview element.
                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = AGV_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + AGV_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
                if (aeGrid == null)
                {
                    sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Console.WriteLine("AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

                Thread.Sleep(3000);
                // Construct the Grid Cell Element Name
                string cellname = "Id Row 0";
                // Get the Element with the Row Col Coordinates
                AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                if (aeCell == null)
                {
                    Console.WriteLine("Find AgvDataGridView aeCell failed:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("cell AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                // find cell value
                string AgvValue = string.Empty;
                try
                {
                    ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    AgvValue = vp.Current.Value;
                    Console.WriteLine("Get element.Current Value:" + AgvValue);
                }
                catch (System.NullReferenceException)
                {
                    AgvValue = string.Empty;
                }

                if (AgvValue == null || AgvValue == string.Empty)
                {
                    sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellname;
                    Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Thread.Sleep(3000);
                //System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                System.Windows.Point point = new Point();
                try
                {
                    point = aeCell.GetClickablePoint();
                }
                catch (Exception ex)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "aeCell.GetClickablePoint()" + ex.Message;
                    Console.WriteLine("aeCell.GetClickablePoint()" + ex.Message);
                    Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "---" + ex.StackTrace, sOnlyUITest);
                    return;
                }

                string StateValue = AUICommon.GetDataGridViewCellValueAt(0, "Mode", aeGrid);
                if (StateValue == null || StateValue == string.Empty)
                {
                    sErrorMessage = "AgvDataGridView aeCell Value not found:" + "Mode Row 0";
                    Console.WriteLine("AgvDataGridView aeCell Value not found:" + "Mode Row 0");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + "Mode Row 0", sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Thread.Sleep(2000);
                Input.MoveToAndClick(point);

                // Select All Agvs
                //Thread.Sleep(2000);
                //Input.SendKeyboardInput(System.Windows.Input.Key.Down, true);
                //Thread.Sleep(2000);
                //Input.SendKeyboardInput(System.Windows.Input.Key.Down, false);

                Thread.Sleep(2000);
                Input.SendKeyboardInput(System.Windows.Input.Key.RightCtrl, true);
                Thread.Sleep(2000);
                Input.SendKeyboardInput(System.Windows.Input.Key.A, true);

                Thread.Sleep(2000);
                Input.SendKeyboardInput(System.Windows.Input.Key.RightCtrl, false);
                Thread.Sleep(2000);
                Input.SendKeyboardInput(System.Windows.Input.Key.A, false);

                Thread.Sleep(2000);
                Input.MoveToAndRightClick(point);

                Thread.Sleep(3000);
                // find Mode point
                System.Windows.Automation.Condition cMode = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, Constants.AGV_MENUITEM_Mode),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
               );

                // Find the MenuItemMode element
                AutomationElement aeMenuItemMode = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cMode);
                if (aeMenuItemMode != null)
                {
                    Console.WriteLine("new element found: " + aeMenuItemMode.Current.Name);
                    Console.WriteLine("err.Current.AutomationId: " + aeMenuItemMode.Current.AutomationId);
                    Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemMode.Current.ControlType.ToString());
                    Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemMode.Current.ControlType.ProgrammaticName);
                    Input.MoveToAndClick(aeMenuItemMode);
                }
                else
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = "mode not found ------------:";
                    Console.WriteLine(sErrorMessage);
                    return;
                }

                Thread.Sleep(2000);
                // find Manual point
                System.Windows.Automation.Condition cManual = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, Constants.AGV_MENUITEM_Removed),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
               );

                // Find the MenuItemMode element
                AutomationElement aeMenuItemManual = aeMenuItemMode.FindFirst(TreeScope.Element | TreeScope.Descendants, cManual);
                if (aeMenuItemManual != null)
                {
                    Console.WriteLine("new element found: " + aeMenuItemManual.Current.Name);
                    Console.WriteLine("err.Current.AutomationId: " + aeMenuItemManual.Current.AutomationId);
                    Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemManual.Current.ControlType.ToString());
                    Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemManual.Current.ControlType.ProgrammaticName);
                    Console.WriteLine("new element x: " + TestTools.AUIUtilities.GetElementCenterPoint(aeMenuItemManual).X);
                    Console.WriteLine("new element Y: " + TestTools.AUIUtilities.GetElementCenterPoint(aeMenuItemManual).Y);

                    Input.MoveTo(new Point(AUIUtilities.GetElementCenterPoint(aeMenuItemManual).X, AUIUtilities.GetElementCenterPoint(aeMenuItemManual).Y));
                    Thread.Sleep(2000);
                    Input.ClickAtPoint(new Point(AUIUtilities.GetElementCenterPoint(aeMenuItemManual).X, AUIUtilities.GetElementCenterPoint(aeMenuItemManual).Y));

                }
                else
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Removed not found ------------:";
                    Console.WriteLine(sErrorMessage);
                    return;
                }

                Thread.Sleep(2000);

                // Find Restart Agv Dialog Window
                System.Windows.Automation.Condition c2 = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, "Put Agvs in mode Removed?"),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );

                // Find the ARestart Agv Dialog element
                AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                if (aeDialog == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = " Dialog Window not found";
                    Console.WriteLine("FindElementByID failed: Dialog");
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                // Find Yes Button
                System.Windows.Automation.Condition c3 = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
               );

                // Find Yes Button element
                AutomationElement aeYes = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c3);
                if (aeYes == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = " Yes Button not found";
                    Console.WriteLine("FindElementByID failed: Yes button");
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeYes));
                Thread.Sleep(2000);
                StateValue = AUICommon.GetDataGridViewCellValueAt(0, "Mode", aeGrid);

                sStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - sStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);
                while (!StateValue.Equals("Removed") && mTime.TotalSeconds < 30)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - sStartTime;
                    StateValue = AUICommon.GetDataGridViewCellValueAt(1, "Mode", aeGrid);
                    Console.WriteLine("time is (sec) : " + mTime.Seconds + " and mode is " + StateValue);
                }

                if (StateValue.Equals("Removed"))
                {
                    result = ConstCommon.TEST_PASS;
                    Console.WriteLine(testname + "  ---pass --- " + StateValue);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine(testname + " ---fail --- " + StateValue);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    return;
                }

                // Check All mode status Removed
                for (int i = 1; i < sNumAgvs; i++)
                {
                    try
                    {
                        StateValue = AUICommon.GetDataGridViewCellValueAt(i, "Mode", aeGrid);
                    }
                    catch (System.NullReferenceException)
                    {
                        break;
                    }

                    sStartTime = DateTime.Now;
                    mTime = DateTime.Now - sStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalSeconds);
                    while (!StateValue.Equals("Removed") && mTime.TotalSeconds < 30)
                    {
                        Thread.Sleep(2000);
                        mTime = DateTime.Now - sStartTime;
                        StateValue = AUICommon.GetDataGridViewCellValueAt(1, "Mode", aeGrid);
                        Console.WriteLine("time is (sec) : " + mTime.Seconds + " and mode is " + StateValue);
                    }

                    if (StateValue.Equals("Removed"))
                    {
                        result = ConstCommon.TEST_PASS;
                        Console.WriteLine(testname + " " + i + " de ---pass --- " + StateValue);
                        //Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    }
                    else
                    {
                        result = ConstCommon.TEST_FAIL;
                        Console.WriteLine(testname + " ---fail --- " + StateValue);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        return;
                    }
                }
                Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region AgvsIdSorting
        public static void AgvsIdSorting(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            try
            {
                AUICommon.ClearDisplayedScreens(root, 2);

                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Input.MoveToAndClick(aePanelLink);

                Thread.Sleep(10000);

                // Find AGV Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, AGV_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the AGVOverview element.
                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = AGV_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + AGV_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
                if (aeGrid == null)
                {
                    sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Console.WriteLine("AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

                Thread.Sleep(3000);
                string[] AgvsIdCells = new string[sNumAgvs];
                for (int i = 0; i < sNumAgvs; i++)
                {
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row " + i;
                    // Get the Element with the Row Col Coordinates
                    AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                    if (aeCell == null)
                    {
                        Console.WriteLine("Find AgvDataGridView aeCell failed:" + cellname);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                    else
                    {
                        Console.WriteLine("cell AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    }

                    // find cell value
                    string AgvValue = string.Empty;
                    try
                    {
                        ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        AgvValue = vp.Current.Value;
                        Console.WriteLine("Get element.Current Value:" + AgvValue);
                    }
                    catch (System.NullReferenceException)
                    {
                        AgvValue = string.Empty;
                    }

                    if (AgvValue == null || AgvValue == string.Empty)
                    {
                        sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellname;
                        Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellname);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }

                    AgvsIdCells[i] = AgvValue;
                }

                // check sorting order
                bool PreSortAscending = true;
                if (AgvsIdCells[0].CompareTo(AgvsIdCells[1]) < 0)
                    PreSortAscending = true;
                else
                    PreSortAscending = false;


                Console.WriteLine("Sort Ascending = " + PreSortAscending);
                Thread.Sleep(3000);
                // Click ID Header Cell
                //double x = aeGrid.Current.BoundingRectangle.Left;
                //double y = aeGrid.Current.BoundingRectangle.Top;
                //Point headpoint = new Point(x + 5, y + 5);
                // Find AGV Overview Window
                System.Windows.Automation.Condition cHeaderId = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, "Id"),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Header)
               );

                AutomationElement aeHeaderId = aeGrid.FindFirst(TreeScope.Element | TreeScope.Descendants, cHeaderId); ;
                Point headpoint = AUIUtilities.GetElementCenterPoint(aeHeaderId);

                Thread.Sleep(1000);

                Input.MoveToAndClick(headpoint);
                Thread.Sleep(2000);
                Input.ClickAtPoint(headpoint);
                Thread.Sleep(3000);

                for (int i = 0; i < sNumAgvs; i++)
                {
                    // Construct the Grid Cell Element Name
                    string cellname = "Id Row " + i;
                    // Get the Element with the Row Col Coordinates
                    AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                    if (aeCell == null)
                    {
                        Console.WriteLine("Find AgvDataGridView aeCell failed:" + cellname);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                    else
                    {
                        Console.WriteLine("cell AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    }

                    // find cell value
                    string AgvValue = string.Empty;
                    try
                    {
                        ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        AgvValue = vp.Current.Value;
                        Console.WriteLine("Get element.Current Value:" + AgvValue);
                    }
                    catch (System.NullReferenceException)
                    {
                        AgvValue = string.Empty;
                    }

                    if (AgvValue == null || AgvValue == string.Empty)
                    {
                        sErrorMessage = "AgvDataGridView aeCell Value not found:" + cellname;
                        Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellname);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }

                    AgvsIdCells[i] = AgvValue;
                }

                Console.WriteLine(" ---Total agvs --- " + sNumAgvs);
                Epia3Common.WriteTestLogMsg(slogFilePath, " ---Total agvs --- " + sNumAgvs, sOnlyUITest);

                bool sortResult = true;
                // Check result
                for (int i = 0; i < sNumAgvs - 1; i++)
                {
                    if (PreSortAscending)
                    {
                        //Console.WriteLine(AgvsIdCells[i] + " ---compare --- " + AgvsIdCells[i + 1]);
                        Epia3Common.WriteTestLogMsg(slogFilePath, AgvsIdCells[i] + " ---compare --- " + AgvsIdCells[i + 1], sOnlyUITest);

                        if (AgvsIdCells[i].CompareTo(AgvsIdCells[i + 1]) < 0)
                        {
                            sortResult = false;
                            break;
                        }
                    }
                    else
                    {
                        if (AgvsIdCells[i].CompareTo(AgvsIdCells[i + 1]) > 0)
                        {
                            sortResult = false;
                            break;
                        }
                    }
                }

                if (sortResult)
                {
                    result = ConstCommon.TEST_PASS;
                    Console.WriteLine(testname + " ---pass --- ");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine(testname + " ---fail --- ");
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    return;
                }

                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region CreateNewTransport
        public static void CreateNewTransport(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            AutomationElement aePanelLink = null;
            try
            {
                if (sEtriccServerStartupOK == false)
                {
                    sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
				if (root != null)
				{
                    root.SetFocus();
					AUICommon.ClearDisplayedScreens(root, 2);
                    Console.WriteLine("\nFind FindTreeViewNodeLevel1:  " + INFRASTRUCTURE + " ===");
                    AutomationElement aeTreeViewNode = AUICommon.FindTreeViewNodeLevel1(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                    if (aeTreeViewNode == null)
					{
                        Console.WriteLine("\nFind aeTreeViewNode == null:  " + INFRASTRUCTURE + " ===");
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
                        Console.WriteLine("\n=== Find aeTreeViewNode and click:  " + INFRASTRUCTURE + " ===");
                        Point Pnt = AUIUtilities.GetElementCenterPoint(aeTreeViewNode);
						Input.MoveToAndClick(Pnt);
						Thread.Sleep(5000);
					}
				}
				else
				{
					sErrorMessage = "aeWindow not found";
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					TestCheck = ConstCommon.TEST_FAIL;
				}

                AutomationElement aeSelectedWindow = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(TRANSPORT_OVERVIEW, ref sErrorMessage);
                    if (aeSelectedWindow == null)
                    {
                        Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeSelectedWindow);
                        if (aeGrid == null)
                        {
                            Console.WriteLine("Find TransportDataGridView failed:" + "TransportDataGridView");
                            Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView failed:" + "TransportDataGridView", sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine("TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            // find NewTransport button
                            /*System.Windows.Automation.Condition cButtonAutosize = new AndCondition(
                                new PropertyCondition(AutomationElement.NameProperty, "m_TxtFilter"),
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                            );*/

                            AutomationElement aeToolBar = AUIUtilities.FindElementByType(ControlType.ToolBar, aeSelectedWindow);
                            if (aeToolBar == null)
                            {
                                Console.WriteLine("Find TOOLBAR failed:" + "TransportDataGridView");
                                Epia3Common.WriteTestLogMsg(slogFilePath, "Find TOOLBAR failed:" + "TransportDataGridView", sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                //AutomationElement aeButtonNewTransport = aeSelectedWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cButtonAutosize);
                                AutomationElement aeButtonNewTransport = AUIUtilities.FindElementByType(ControlType.Edit, aeToolBar);
                                if (aeButtonNewTransport == null)
                                {
                                    sErrorMessage = "aeButtonNewTransport not find :" + aeButtonNewTransport.Current.Name;
                                    Console.WriteLine(sErrorMessage);
                                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                                else
                                {
                                    double x = aeButtonNewTransport.Current.BoundingRectangle.Right + 40.0;
                                    double y = (aeButtonNewTransport.Current.BoundingRectangle.Bottom + aeButtonNewTransport.Current.BoundingRectangle.Top) / 2.0;
                                    Point pt = new Point(x, y);
                                    Input.MoveTo(pt);
                                    Input.ClickAtPoint(pt);
                                    Thread.Sleep(3000);
                                }
                            }
                        }
                    }
                }

                //find new transport screen
                // first resize root (there is a bug create button not displayed)
                AutomationElement aeNewT = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                    if (root != null)
                    {
                        TransformPattern tranform = root.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                        if (tranform != null)
                        {
                            tranform.Resize(root.Current.BoundingRectangle.Width,
                                0.85 * System.Windows.Forms.SystemInformation.VirtualScreen.Height);
                            Thread.Sleep(1000);
                            tranform.Move(0, 0);
                        }
                    }

                    System.Windows.Automation.Condition c2 = new AndCondition(
                      new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                    );

                    // Find the NewTransport Screen element.
                    aeNewT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                    if (aeNewT == null)
                    {
                        sErrorMessage = "New Transport Window not found";
                        Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                //MessageBox.Show("Check New Transport Window Open?", "Click OK to continue ....", MessageBoxButtons.OK);
                Thread.Sleep(2000);
                //find command list box
                // Find the NewTransport Screen element.
                String CommandID = "Drop";
                String MoverID = "FLV";
                String SourceID = "";
                String DestID = "";
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(2000);
                    AutomationElement aeCommandList = AUIUtilities.FindElementByID("commandListBox", aeNewT);
                    if (aeCommandList == null)
                    {
                        sErrorMessage = "Command listBox not found";
                        Console.WriteLine("FindElementByID failed:" + "commandListBox");
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Find listitem Command.
                        SelectionPattern selectPattern =
                                  aeCommandList.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                        //sProjectFile = "TestProject.zip";
                        if (sProjectFile.ToLower().IndexOf("demo") >= 0)
                        {
                            CommandID = "Drop";
                            MoverID = "FLV";
                            SourceID = "FLV_L101";
                            DestID = "FLV_L203";
                        }
                        else if (sProjectFile.ToLower().IndexOf("eurobaltic") >= 0)
                        {
                            CommandID = "Move";
                            MoverID = "AGV1";
                            SourceID = "0040-01-01";
                            DestID = "0030-01-01";
                        }
                        else
                        {
                            CommandID = "Move";
                            MoverID = "AGV1";
                            SourceID = "M_01_01_01_01";
                            DestID = "ABF_1_1_T";
                        }

                        AutomationElement item
                            = AUIUtilities.FindElementByName(CommandID, aeCommandList);
                        if (item != null)
                        {
                            Console.WriteLine(CommandID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(2000);

                            SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            itemPattern.Select();

                            Thread.Sleep(2000);

                            Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));

                            Thread.Sleep(2000);
                        }
                        else
                        {
                            sErrorMessage = "Finding Command item " + CommandID + "  failed";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }


                    // find source location button
                    string sourceLocId = "rdSourceLocations";
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        AutomationElement aeSourceRadio = AUIUtilities.FindElementByID(sourceLocId, aeNewT);
                        if (aeSourceRadio == null)
                        {
                            sErrorMessage = "New Transport aeSourceRadio not found";
                            Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            if (CommandID.StartsWith("Pick") || CommandID.StartsWith("Move"))
                            {
                                SelectionItemPattern itemRadioPattern = aeSourceRadio.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemRadioPattern.Select();
                                Thread.Sleep(3000);

                                string sourceIDListBoxId = "sourceIdListBox";
                                AutomationElement aeSrcListBox = AUIUtilities.FindElementByID(sourceIDListBoxId, aeNewT);

                                if (aeSrcListBox == null)
                                {
                                    sErrorMessage = "New Transport aeSourceListBox not found";
                                    Console.WriteLine("FindElementByID failed:" + sourceIDListBoxId);
                                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                                else
                                {
                                    AutomationElement SrcItem
                                       = AUIUtilities.FindElementByName(SourceID, aeSrcListBox);

                                    if (SrcItem != null)
                                    {
                                        Console.WriteLine(SourceID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                        Thread.Sleep(2000);

                                        SelectionItemPattern itemPattern = SrcItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                        itemPattern.Select();

                                        Thread.Sleep(2000);
                                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(SrcItem));
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        sErrorMessage = "Finding Command item " + SourceID + "  failed";
                                        Console.WriteLine(sErrorMessage);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                    }
                                }
                            }
                        }
                    }

                    // find destination location button
                    string destLocId = "rdDestinationLocations";
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        AutomationElement aeDestRadio = AUIUtilities.FindElementByID(destLocId, aeNewT);
                        if (aeDestRadio == null)
                        {
                            sErrorMessage = "New Transport aeDestRadio not found";
                            Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            if (CommandID.StartsWith("Drop") || CommandID.StartsWith("Move") || CommandID.StartsWith("Wait"))
                            {
                                SelectionItemPattern itemPattern = aeDestRadio.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemPattern.Select();
                                Thread.Sleep(3000);
                                string destIDListBoxId = "destinationIdListBox";
                                AutomationElement aeDestListBox = AUIUtilities.FindElementByID(destIDListBoxId, aeNewT);
                                if (aeDestListBox == null)
                                {
                                    sErrorMessage = "New Transport aeDestListBox not found";
                                    Console.WriteLine("FindElementByID failed:" + destIDListBoxId);
                                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                                else
                                {
                                    AutomationElement DestItem
                                       = AUIUtilities.FindElementByName(DestID, aeDestListBox);

                                    if (DestItem != null)
                                    {
                                        Console.WriteLine(DestID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                        Thread.Sleep(2000);

                                        SelectionItemPattern itemDestPattern = DestItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                        itemDestPattern.Select();

                                        Thread.Sleep(2000);
                                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(DestItem));
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        sErrorMessage = "Finding Command item " + DestID + "  failed";
                                        Console.WriteLine(sErrorMessage);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                    }
                                }
                            }
                        }
                    }



                    // Find MOVER element.
                    string MoverId = "moverIDComboBox";
                    /*if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        AutomationElement aeMoverList = AUIUtilities.FindElementByID(MoverId, aeNewT);
                        if (aeMoverList == null)
                        {
                            sErrorMessage = "Mover aeMover not found";
                            Console.WriteLine("FindElementByID failed:" + MoverId);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            // Find listitem Mover.
                            SelectionPattern selectMoverPattern =
                                      aeMoverList.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                            AutomationElement MoverItem
                                = AUIUtilities.FindElementByName(MoverID, aeMoverList);
                            if (MoverItem != null)
                            {
                                Console.WriteLine(MoverID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                Thread.Sleep(2000);

                                SelectionItemPattern itemPattern = MoverItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemPattern.Select();

                                Thread.Sleep(2000);

                                //Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(MoverItem));
                                Thread.Sleep(2000);
                            }
                            else
                            {
                                sErrorMessage = "Finding Command item " + MoverID + "  failed";
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                    }
                    */
                    // Find Create element. 
                    string id = "m_btnSave";
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        AutomationElement aeCreate = AUIUtilities.FindElementByID(id, aeNewT);
                        if (aeCreate == null)
                        {
                            sErrorMessage = "New Transport aeCreate not found";
                            Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine("aeCreate Found:");
                            Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCreate));
                            Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCreate));
                        }

                        Thread.Sleep(4000);
                        /* Cancel button removed since 22 sept 2011
                        if ( sBuildNr.IndexOf("Dev02") >= 0)
                        {
                            Console.WriteLine("Do nothing  No Cancel button :" + NEW_TRANSPORT);
                        }
                        else     // close pop up transport screen  // old screen
                        {
                            // Dispose new Transport screen, Find Cancel element. 
                            string CancelId = "m_btnCancel";
                            AutomationElement aeCancel = AUIUtilities.FindElementByID(CancelId, aeNewT);
                            if (aeCancel == null)
                            {
                                result = ConstCommon.TEST_FAIL;
                                sErrorMessage = "New Transport aeCancel not found";
                                Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                                Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                                return;
                            }
                            else
                            {
                                Console.WriteLine("Cancel Found:");
                                Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancel));
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCancel));
                            }
                        }
                        */
                        Thread.Sleep(4000);

                        if (sProjectFile.ToLower().IndexOf("testproject") >= 0)
                            Thread.Sleep(10000);
                    }
                }

               
                // Check transport created
                //AUICommon.ClearDisplayedScreens(root);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(2000);
                    aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Input.MoveToAndClick(aePanelLink);
                        Thread.Sleep(2000);
                        aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(TRANSPORT_OVERVIEW, ref sErrorMessage);
                        if (aeSelectedWindow == null)
                        {
                            Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }
                
                // Find the TransportOverview element.
                // Find Transport Overview Window
                /*System.Windows.Automation.Condition c22222 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, TRANSPORT_OVERVIEW_TITLE),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );
                AutomationElement aeOverview22222 = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c22222);

                if (aeOverview22222 == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + TRANSPORT_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_PASS;
                    string ms = TRANSPORT_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }*/

                Thread.Sleep(2000);
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine("\nTest scenario " + testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }


        }

        public static void CreateNewTransport1(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            AutomationElement aePanelLink = null;
            try
            {
                if (sEtriccServerStartupOK == false)
                {
                    sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                if (root != null)
                {
                    root.SetFocus();
                    AUICommon.ClearDisplayedScreens(root, 2);
                    aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                    else
                        Input.MoveToAndClick(aePanelLink);

                    Thread.Sleep(10000);
                    /*
                    // Find Transport Overview Window Element
                    System.Windows.Automation.Condition c = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, TRANSPORT_OVERVIEW_TITLE),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                    );

                    AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeOverview == null)
                    {
                        result = ConstCommon.TEST_FAIL;
                        sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
                        Console.WriteLine("FindElementByID failed:" + TRANSPORT_OVERVIEW);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        return;
                    }
                    */
                }
                else
                {
                    sErrorMessage = "aeWindow not found";
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                AutomationElement aeSelectedWindow = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(TRANSPORT_OVERVIEW, ref sErrorMessage);
                    if (aeSelectedWindow == null)
                    {
                        Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeSelectedWindow);
                        if (aeGrid == null)
                        {
                            Console.WriteLine("Find TransportDataGridView failed:" + "TransportDataGridView");
                            Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView failed:" + "TransportDataGridView", sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine("TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            // find NewTransport button
                            /*System.Windows.Automation.Condition cButtonAutosize = new AndCondition(
                                new PropertyCondition(AutomationElement.NameProperty, "m_TxtFilter"),
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                            );*/

                            AutomationElement aeToolBar = AUIUtilities.FindElementByType(ControlType.ToolBar, aeSelectedWindow);
                            if (aeToolBar == null)
                            {
                                Console.WriteLine("Find TOOLBAR failed:" + "TransportDataGridView");
                                Epia3Common.WriteTestLogMsg(slogFilePath, "Find TOOLBAR failed:" + "TransportDataGridView", sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                //AutomationElement aeButtonNewTransport = aeSelectedWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cButtonAutosize);
                                AutomationElement aeButtonNewTransport = AUIUtilities.FindElementByType(ControlType.Edit, aeToolBar);
                                if (aeButtonNewTransport == null)
                                {
                                    sErrorMessage = "aeButtonNewTransport not find :" + aeButtonNewTransport.Current.Name;
                                    Console.WriteLine(sErrorMessage);
                                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                                else
                                {
                                    Console.WriteLine("aeButtonNewTransport find ...");
                                    double x = aeButtonNewTransport.Current.BoundingRectangle.Right + 40.0;
                                    double y = (aeButtonNewTransport.Current.BoundingRectangle.Bottom + aeButtonNewTransport.Current.BoundingRectangle.Top) / 2.0;
                                    Point pt = new Point(x, y);
                                    Input.MoveTo(pt);
                                    Input.ClickAtPoint(pt);
                                    Thread.Sleep(3000);
                                }
                            }
                        }
                    }
                }

                //find new transport screen
                // first resize root (there is a bug create button not displayed)
                AutomationElement aeNewT = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                    if (root != null)
                    {
                        TransformPattern tranform = root.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                        if (tranform != null)
                        {
                            tranform.Resize(root.Current.BoundingRectangle.Width,
                                0.85 * System.Windows.Forms.SystemInformation.VirtualScreen.Height);
                            Thread.Sleep(1000);
                            tranform.Move(0, 0);
                        }
                    }

                    System.Windows.Automation.Condition c2 = new AndCondition(
                      new PropertyCondition(AutomationElement.NameProperty, "Add new transport"),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                    );

                    // Find the NewTransport Screen element.
                    aeNewT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                    if (aeNewT == null)
                    {
                        sErrorMessage = "New Transport Window not found";
                        Console.WriteLine("FindElementByID failed:" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                //MessageBox.Show("Check New Transport Window Open?", "Click OK to continue ....", MessageBoxButtons.OK);
                Thread.Sleep(2000);
                //find command list box
                // Find the NewTransport Screen element.
                String CommandID = "Drop";
                String MoverID = "FLV";
                String SourceID = "";
                String DestID = "";
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(2000);
                    AutomationElement aeCommandList = AUIUtilities.FindElementByID("commandListBox", aeNewT);
                    if (aeCommandList == null)
                    {
                        sErrorMessage = "Command listBox not found";
                        Console.WriteLine("FindElementByID failed:" + "commandListBox");
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Find listitem Command.
                        SelectionPattern selectPattern =
                                  aeCommandList.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                        //sProjectFile = "TestProject.zip";
                        if (sProjectFile.ToLower().IndexOf("demo") >= 0)
                        {
                            CommandID = "Move";
                            MoverID = "FLV";
                            SourceID = "FLV_L101";
                            DestID = "FLV_L203";
                        }
                        else if (sProjectFile.ToLower().IndexOf("eurobaltic") >= 0)
                        {
                            CommandID = "Move";
                            MoverID = "AGV1";
                            SourceID = "0040-01-01";
                            DestID = "0030-01-01";
                        }
                        else
                        {
                            CommandID = "Move";
                            MoverID = "AGV1";
                            SourceID = "M_01_01_01_01";
                            DestID = "ABF_1_1_T";
                        }

                        AutomationElement item
                            = AUIUtilities.FindElementByName(CommandID, aeCommandList);
                        if (item != null)
                        {
                            Console.WriteLine(CommandID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(2000);

                            SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            itemPattern.Select();

                            Thread.Sleep(2000);

                            Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));

                            Thread.Sleep(2000);
                        }
                        else
                        {
                            sErrorMessage = "Finding Command item " + CommandID + "  failed";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }


                    // find source location button
                    string sourceLocId = "rdSourceLocations";
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        AutomationElement aeSourceRadio = AUIUtilities.FindElementByID(sourceLocId, aeNewT);
                        if (aeSourceRadio == null)
                        {
                            sErrorMessage = "New Transport aeSourceRadio not found";
                            Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            if (CommandID.StartsWith("Pick") || CommandID.StartsWith("Move"))
                            {
                                SelectionItemPattern itemRadioPattern = aeSourceRadio.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemRadioPattern.Select();
                                Thread.Sleep(3000);

                                string sourceIDListBoxId = "sourceIdListBox";
                                AutomationElement aeSrcListBox = AUIUtilities.FindElementByID(sourceIDListBoxId, aeNewT);

                                if (aeSrcListBox == null)
                                {
                                    sErrorMessage = "New Transport aeSourceListBox not found";
                                    Console.WriteLine("FindElementByID failed:" + sourceIDListBoxId);
                                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                                else
                                {
                                    AutomationElement SrcItem
                                       = AUIUtilities.FindElementByName(SourceID, aeSrcListBox);

                                    if (SrcItem != null)
                                    {
                                        Console.WriteLine(SourceID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                        Thread.Sleep(2000);

                                        SelectionItemPattern itemPattern = SrcItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                        itemPattern.Select();

                                        Thread.Sleep(2000);
                                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(SrcItem));
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        sErrorMessage = "Finding Command item " + SourceID + "  failed";
                                        Console.WriteLine(sErrorMessage);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                    }
                                }
                            }
                        }
                    }

                    // find destination location button
                    string destLocId = "rdDestinationLocations";
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        AutomationElement aeDestRadio = AUIUtilities.FindElementByID(destLocId, aeNewT);
                        if (aeDestRadio == null)
                        {
                            sErrorMessage = "New Transport aeDestRadio not found";
                            Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            if (CommandID.StartsWith("Drop") || CommandID.StartsWith("Move") || CommandID.StartsWith("Wait"))
                            {
                                SelectionItemPattern itemPattern = aeDestRadio.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemPattern.Select();
                                Thread.Sleep(3000);
                                string destIDListBoxId = "destinationIdListBox";
                                AutomationElement aeDestListBox = AUIUtilities.FindElementByID(destIDListBoxId, aeNewT);
                                if (aeDestListBox == null)
                                {
                                    sErrorMessage = "New Transport aeDestListBox not found";
                                    Console.WriteLine("FindElementByID failed:" + destIDListBoxId);
                                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                                else
                                {
                                    AutomationElement DestItem
                                       = AUIUtilities.FindElementByName(DestID, aeDestListBox);

                                    if (DestItem != null)
                                    {
                                        Console.WriteLine(DestID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                        Thread.Sleep(2000);

                                        SelectionItemPattern itemDestPattern = DestItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                        itemDestPattern.Select();

                                        Thread.Sleep(2000);
                                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(DestItem));
                                        Thread.Sleep(2000);
                                    }
                                    else
                                    {
                                        sErrorMessage = "Finding Command item " + DestID + "  failed";
                                        Console.WriteLine(sErrorMessage);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                    }
                                }
                            }
                        }
                    }



                    // Find MOVER element.
                    string MoverId = "moverIDComboBox";
                    /*if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        AutomationElement aeMoverList = AUIUtilities.FindElementByID(MoverId, aeNewT);
                        if (aeMoverList == null)
                        {
                            sErrorMessage = "Mover aeMover not found";
                            Console.WriteLine("FindElementByID failed:" + MoverId);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            // Find listitem Mover.
                            SelectionPattern selectMoverPattern =
                                      aeMoverList.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                            AutomationElement MoverItem
                                = AUIUtilities.FindElementByName(MoverID, aeMoverList);
                            if (MoverItem != null)
                            {
                                Console.WriteLine(MoverID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                Thread.Sleep(2000);

                                SelectionItemPattern itemPattern = MoverItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemPattern.Select();

                                Thread.Sleep(2000);

                                //Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(MoverItem));
                                Thread.Sleep(2000);
                            }
                            else
                            {
                                sErrorMessage = "Finding Command item " + MoverID + "  failed";
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                    }*/

                    // Find Create element. 
                    string id = "m_btnSave";
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        AutomationElement aeCreate = AUIUtilities.FindElementByID(id, aeNewT);
                        if (aeCreate == null)
                        {
                            sErrorMessage = "New Transport aeCreate not found";
                            Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine("aeCreate Found:");
                            Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCreate));
                            Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCreate));
                        }

                        Thread.Sleep(4000);
                        /* Cancel button removed since 22 sept 2011
                        if ( sBuildNr.IndexOf("Dev02") >= 0)
                        {
                            Console.WriteLine("Do nothing  No Cancel button :" + NEW_TRANSPORT);
                        }
                        else     // close pop up transport screen  // old screen
                        {
                            // Dispose new Transport screen, Find Cancel element. 
                            string CancelId = "m_btnCancel";
                            AutomationElement aeCancel = AUIUtilities.FindElementByID(CancelId, aeNewT);
                            if (aeCancel == null)
                            {
                                result = ConstCommon.TEST_FAIL;
                                sErrorMessage = "New Transport aeCancel not found";
                                Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                                Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                                return;
                            }
                            else
                            {
                                Console.WriteLine("Cancel Found:");
                                Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancel));
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCancel));
                            }
                        }
                        */
                        Thread.Sleep(4000);

                        if (sProjectFile.ToLower().IndexOf("testproject") >= 0)
                            Thread.Sleep(10000);
                    }
                }


                // Check transport created
                //AUICommon.ClearDisplayedScreens(root);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(2000);
                    aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Input.MoveToAndClick(aePanelLink);
                        Thread.Sleep(2000);
                        aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(TRANSPORT_OVERVIEW, ref sErrorMessage);
                        if (aeSelectedWindow == null)
                        {
                            Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                // Find the TransportOverview element.
                // Find Transport Overview Window
                /*System.Windows.Automation.Condition c22222 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, TRANSPORT_OVERVIEW_TITLE),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );
                AutomationElement aeOverview22222 = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c22222);

                if (aeOverview22222 == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + TRANSPORT_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_PASS;
                    string ms = TRANSPORT_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }*/

                Thread.Sleep(2000);
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine("\nTest scenario " + testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }


        }
        #endregion CreateNewTransport2
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region CreateNewTransport2
        public static void CreateNewTransport2(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_FAIL;

            AutomationElement aePanelLink = null;
            try
            {
                if (sEtriccServerStartupOK == false)
                {
                    sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                    return;
                }

                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                AUICommon.ClearDisplayedScreens(root, 2);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Console.WriteLine("\n=== Find new Transport ===");
                System.Windows.Forms.TreeNode treeNode = new TreeNode();
                AutomationElement aeNode = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, NEW_TRANSPORT, ref sErrorMessage);
                if (aeNode == null)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine("\n=== New Transport NOT Exist ===");
                    Input.MoveToAndDoubleClick(aePanelLink.GetClickablePoint());   // double click Transports menu item aePanelLink
                    Thread.Sleep(9000);
                    aeNode = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, NEW_TRANSPORT, ref sErrorMessage);
                    //Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    //result = ConstCommon.TEST_FAIL;
                    //return;
                }
                else
                {
                    Console.WriteLine("\n=== New Transport Exist ===");
                    //Input.MoveToAndClick(aeNode);
                }

                if (aeNode == null)
                {
                    Console.WriteLine("Node not exist:" + TRANSPORT_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Input.MoveToAndClick(aeNode);    // click New Transport menu item

                Thread.Sleep(2000);

                //find new transport screen
                // first resize root (there is a bug create button not displayed)
                root = ProjBasicUI.GetMainWindowWithinTime("MainForm",10);
                if (root != null)
                {
                    TransformPattern tranform = root.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                    if (tranform != null)
                    {
                        tranform.Resize(root.Current.BoundingRectangle.Width,
                            0.85* System.Windows.Forms.SystemInformation.VirtualScreen.Height);
                        Thread.Sleep(1000);
                        tranform.Move(0, 0);
                    }
                }

                System.Windows.Automation.Condition c2 = new AndCondition(
                  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );

                // Find the NewTransport Screen element.
                AutomationElement aeNewT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                if (aeNewT == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "New Transport Window not found";
                    Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                Thread.Sleep(2000);

                //find command list box
                // Find the NewTransport Screen element.
                AutomationElement aeCommandList = AUIUtilities.FindElementByID("commandListBox", aeNewT);
                if (aeCommandList == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Command listBox not found";
                    Console.WriteLine("FindElementByID failed:" + "commandListBox");
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                // Find listitem Command.
                SelectionPattern selectPattern =
                          aeCommandList.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                //sProjectFile = "TestProject.zip";
                String CommandID = "Drop";
                String MoverID = "FLV";
                String SourceID = "";
                String DestID = "";
                if (sProjectFile.ToLower().IndexOf("demo") >= 0)
                {
                    CommandID = "Drop";
                    MoverID = "FLV";
                    SourceID = "FLV_L101";
                    DestID = "FLV_L203";
                }
                else if (sProjectFile.ToLower().IndexOf("eurobaltic") >= 0)
                {
                    CommandID = "Move";
                    MoverID = "AGV1";
                    SourceID = "0040-01-01";
                    DestID = "0030-01-01";
                }
                else
                {
                    CommandID = "Move";
                    MoverID = "AGV1";
                    SourceID = "M_01_01_01_01";
                    DestID = "ABF_1_1_T";
                }

                AutomationElement item
                    = AUIUtilities.FindElementByName(CommandID, aeCommandList);
                if (item != null)
                {
                    Console.WriteLine(CommandID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    Thread.Sleep(2000);

                    SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                    itemPattern.Select();

                    Thread.Sleep(2000);

                    Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));

                    Thread.Sleep(2000);
                }
                else
                {
                    sErrorMessage = "Finding Command item " + CommandID + "  failed";
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // find source location button
                string sourceLocId = "rdSourceLocations";
                AutomationElement aeSourceRadio = AUIUtilities.FindElementByID(sourceLocId, aeNewT);

                if (aeSourceRadio == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "New Transport aeSourceRadio not found";
                    Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }
                else
                {
                    if (CommandID.StartsWith("Pick") || CommandID.StartsWith("Move"))
                    {
                        SelectionItemPattern itemRadioPattern = aeSourceRadio.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemRadioPattern.Select();
                        Thread.Sleep(3000);

                        string sourceIDListBoxId = "sourceIdListBox";
                        AutomationElement aeSrcListBox = AUIUtilities.FindElementByID(sourceIDListBoxId, aeNewT);

                        if (aeSrcListBox == null)
                        {
                            result = ConstCommon.TEST_FAIL;
                            sErrorMessage = "New Transport aeSourceListBox not found";
                            Console.WriteLine("FindElementByID failed:" + sourceIDListBoxId);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            return;
                        }
                        else
                        {
                            AutomationElement SrcItem
                               = AUIUtilities.FindElementByName(SourceID, aeSrcListBox);

                            if (SrcItem != null)
                            {
                                Console.WriteLine(SourceID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                Thread.Sleep(2000);

                                SelectionItemPattern itemPattern = SrcItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemPattern.Select();

                                Thread.Sleep(2000);
                                Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(SrcItem));
                                Thread.Sleep(2000);
                            }
                            else
                            {
                                sErrorMessage = "Finding Command item " + SourceID + "  failed";
                                Console.WriteLine(sErrorMessage);
                                result = ConstCommon.TEST_FAIL;
                                return;
                            }
                        }
                    }
                }

                // find destination location button
                string destLocId = "rdDestinationLocations";
                AutomationElement aeDestRadio = AUIUtilities.FindElementByID(destLocId, aeNewT);

                if (aeDestRadio == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "New Transport aeDestRadio not found";
                    Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }
                else
                {
                    if (CommandID.StartsWith("Drop") || CommandID.StartsWith("Move") || CommandID.StartsWith("Wait"))
                    {
                        SelectionItemPattern itemPattern = aeDestRadio.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemPattern.Select();
                        Thread.Sleep(3000);

                        string destIDListBoxId = "destinationIdListBox";
                        AutomationElement aeDestListBox = AUIUtilities.FindElementByID(destIDListBoxId, aeNewT);

                        if (aeDestListBox == null)
                        {
                            result = ConstCommon.TEST_FAIL;
                            sErrorMessage = "New Transport aeDestListBox not found";
                            Console.WriteLine("FindElementByID failed:" + destIDListBoxId);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                            return;
                        }
                        else
                        {
                            AutomationElement DestItem
                               = AUIUtilities.FindElementByName(DestID, aeDestListBox);

                            if (DestItem != null)
                            {
                                Console.WriteLine(DestID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                Thread.Sleep(2000);

                                SelectionItemPattern itemDestPattern = DestItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemDestPattern.Select();

                                Thread.Sleep(2000);
                                Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(DestItem));
                                Thread.Sleep(2000);
                            }
                            else
                            {
                                sErrorMessage = "Finding Command item " + DestID + "  failed";
                                Console.WriteLine(sErrorMessage);
                                result = ConstCommon.TEST_FAIL;
                                return;
                            }
                        }
                    }
                }

                // Find MOVER element.
                string MoverId = "moverIDComboBox";
                AutomationElement aeMoverList = AUIUtilities.FindElementByID(MoverId, aeNewT);
                if (aeMoverList == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Mover aeMover not found";
                    Console.WriteLine("FindElementByID failed:" + MoverId);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }
                else
                {
                    // Find listitem Mover.
                    SelectionPattern selectMoverPattern =
                              aeMoverList.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                    AutomationElement MoverItem
                        = AUIUtilities.FindElementByName(MoverID, aeMoverList);
                    if (MoverItem != null)
                    {
                        Console.WriteLine(MoverID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(2000);

                        SelectionItemPattern itemPattern = MoverItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemPattern.Select();

                        Thread.Sleep(2000);

                        //Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(MoverItem));
                        Thread.Sleep(2000);
                    }
                    else
                    {
                        sErrorMessage = "Finding Command item " + MoverID + "  failed";
                        Console.WriteLine(sErrorMessage);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                }

                // Find Create element. 
                string id = "m_btnSave";
                AutomationElement aeCreate = AUIUtilities.FindElementByID(id, aeNewT);
                if (aeCreate == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "New Transport aeCreate not found";
                    Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }
                else
                {
                    Console.WriteLine("aeCreate Found:");
                    Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCreate));
                    Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCreate));
                }

                Thread.Sleep(4000);
                /* Cancel button removed since 22 sept 2011
                if ( sBuildNr.IndexOf("Dev02") >= 0)
                {
                    Console.WriteLine("Do nothing  No Cancel button :" + NEW_TRANSPORT);
                }
                else     // close pop up transport screen  // old screen
                {
                    // Dispose new Transport screen, Find Cancel element. 
                    string CancelId = "m_btnCancel";
                    AutomationElement aeCancel = AUIUtilities.FindElementByID(CancelId, aeNewT);
                    if (aeCancel == null)
                    {
                        result = ConstCommon.TEST_FAIL;
                        sErrorMessage = "New Transport aeCancel not found";
                        Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        return;
                    }
                    else
                    {
                        Console.WriteLine("Cancel Found:");
                        Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancel));
                        Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCancel));
                    }
                }
                */
                Thread.Sleep(4000);

                if (sProjectFile.ToLower().IndexOf("testproject") >= 0)
                    Thread.Sleep(10000);


                Thread.Sleep(2000);
                // Check transport created
                //AUICommon.ClearDisplayedScreens(root);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Input.MoveToAndClick(aePanelLink);
                Thread.Sleep(2000);
                // Find the TransportOverview element.
                // Find Transport Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, TRANSPORT_OVERVIEW_TITLE),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );
                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);

                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + TRANSPORT_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_PASS;
                    string ms = TRANSPORT_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }


        }
        #endregion CreateNewTransport2
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region EditTransport
        public static void EditTransport(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            try
            {
                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                AUICommon.ClearDisplayedScreens(root, 2);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Thread.Sleep(3000);
                }

                Input.MoveToAndClick(aePanelLink);
                Thread.Sleep(10000);

                // Find Transport Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, TRANSPORT_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                // Find the TransportOverview element.
                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + TRANSPORT_OVERVIEW_TITLE);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                // Find Transport GridView
                AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                if (aeGrid == null)
                {
                    Console.WriteLine("Find TransportDataGridView failed:" + "TransportDataGridView");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView failed:" + "TransportDataGridView", sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    // find Autosize button
                    /*System.Windows.Automation.Condition cButtonAutosize = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, "m_TxtFilter"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                    );*/

                    AutomationElement aeToolBar = AUIUtilities.FindElementByType(ControlType.ToolBar, aeOverview);
                    if (aeToolBar == null)
                    {
                        Console.WriteLine("Find TOOLBAR failed:" + "TransportDataGridView");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Find TOOLBAR failed:" + "TransportDataGridView", sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        //AutomationElement aeButtonNewTransport = aeSelectedWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cButtonAutosize);
                        AutomationElement aeButtonAutosize = AUIUtilities.FindElementByType(ControlType.Edit, aeToolBar);
                        if (aeButtonAutosize == null)
                        {
                            Console.WriteLine("aeButtonAutosize not find :" + aeButtonAutosize.Current.Name);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            double x = aeButtonAutosize.Current.BoundingRectangle.Right + 88.0;
                            double y = (aeButtonAutosize.Current.BoundingRectangle.Bottom + aeButtonAutosize.Current.BoundingRectangle.Top) / 2.0;
                            Point pt = new Point(x, y);
                            for (int irole = 1; irole < 4; irole++)
                            {
                                Console.WriteLine("click:" + irole);
                                Input.MoveTo(pt);
                                Input.ClickAtPoint(pt);
                                Thread.Sleep(3000);
                                pt = new Point(x + 2, y);
                            }
                        }
                    }
                }

                Thread.Sleep(3000);
                // Construct the Grid Cell Element Name
                //string cellname = "Id Row 0";
                // Get the Element with the Row Col Coordinates
                //AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                string rowname = "Row 0";
                // Get the Element with the Row Coordinates
                AutomationElement aeRow0 = AUIUtilities.FindElementByName(rowname, aeGrid);

                if (aeRow0 == null)
                {
                    Console.WriteLine("Find TransportDataGridView aeRow0 failed:" + rowname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView row failed:" + rowname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("Row0 TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                Thread.Sleep(3000);
                System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeRow0);
                Input.MoveToAndRightClick(point);
                Thread.Sleep(2000);

                // find Edit Transport point
                System.Windows.Automation.Condition cM = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, Constants.TRANSPORT_MENUITEM_Edit_Transport),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
               );

                // Find the MenuItem Edit Transport element
                AutomationElement aeMenuItemEditTransport = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
                if (aeMenuItemEditTransport != null)
                {
                    Console.WriteLine("new element found: " + aeMenuItemEditTransport.Current.Name);
                    Console.WriteLine("err.Current.AutomationId: " + aeMenuItemEditTransport.Current.AutomationId);
                    Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemEditTransport.Current.ControlType.ToString());
                    Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemEditTransport.Current.ControlType.ProgrammaticName);
                    Input.MoveToAndClick(aeMenuItemEditTransport);
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Edit Transport menu item not found ------------:";
                    Console.WriteLine(sErrorMessage);
                    return;
                }

                Thread.Sleep(2000);
                //MessageBox.Show("Check Edit Transport 1?", "Click OK to continue ....", MessageBoxButtons.OK);
                //find edit transport screen
                System.Windows.Automation.Condition cOld = new AndCondition(
                 new PropertyCondition(AutomationElement.AutomationIdProperty, "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.AddEditTransport"),
                 new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
               );

                System.Windows.Automation.Condition cNew = new AndCondition(
                  new PropertyCondition(AutomationElement.AutomationIdProperty, "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.AddEditTransport.AddEditTransport"),
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );

                // Find the NewTransport Screen element.
                root = EtriccUtilities.GetMainWindow("MainForm");
                AutomationElement aeNewT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cNew);
                if (aeNewT == null)
                {
                    Console.WriteLine("cNew not found try cOld :" + EDIT_TRANSPORT);
                    aeNewT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cOld);
                }
                    

                if (aeNewT == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Edit Transport Window not found";
                    Console.WriteLine("FindElementByID failed:" + EDIT_TRANSPORT);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                Thread.Sleep(2000);

                String DestID = "";
                if (sProjectFile.ToLower().IndexOf("demo") >= 0)
                {
                    DestID = "FLV_L201";
                }
                else if (sProjectFile.ToLower().IndexOf("eurobaltic") >= 0)
                {
                    DestID = "0030-03-01";
                }
                else
                {
                    DestID = "ABF_1_3_T";
                }

                // find destination location button   

                string destLocId = "rdDestinationLocations";
                AutomationElement aeDestRadio = AUIUtilities.FindElementByID(destLocId, aeNewT);

                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                while (aeDestRadio == null && mTime.TotalSeconds < 120)
                {
                    aeDestRadio = AUIUtilities.FindElementByID(destLocId, aeNewT);
                    Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime.TotalSeconds);
                    mTime = DateTime.Now - mStartTime;
                }

                if (aeDestRadio == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "After 2 min, Edit Transport aeDestRadio still not found";
                    Console.WriteLine("FindElementByID failed:" + EDIT_TRANSPORT);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                SelectionItemPattern itemPattern = aeDestRadio.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                itemPattern.Select();
                Thread.Sleep(3000);

                string destIDListBoxId = "destinationIdListBox";
                AutomationElement aeDestListBox = AUIUtilities.FindElementByID(destIDListBoxId, aeNewT);

                if (aeDestListBox == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "New Transport aeDestListBox not found";
                    Console.WriteLine("FindElementByID failed:" + destIDListBoxId);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }
                else
                {
                    AutomationElement DestItem
                       = AUIUtilities.FindElementByName(DestID, aeDestListBox);

                    if (DestItem != null)
                    {
                        Console.WriteLine(DestID + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(2000);

                        SelectionItemPattern itemDestPattern = DestItem.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemDestPattern.Select();

                        Thread.Sleep(2000);
                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(DestItem));
                        Thread.Sleep(2000);
                    }
                    else
                    {
                        sErrorMessage = "Finding Command item " + DestID + "  failed";
                        Console.WriteLine(sErrorMessage);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }
                }
                //MessageBox.Show("Check Edit Transport 2?", "Click OK to continue ....", MessageBoxButtons.OK);
                // Find Save element.
                string id = "m_btnSave";
                AutomationElement aeBtnSave = AUIUtilities.FindElementByID(id, aeNewT);

                if (aeBtnSave == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = "New Transport aeBtnSave not found";
                    Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }
                else
                {
                    Console.WriteLine("aeBtnSave Found:");
                    Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeBtnSave));
                    Thread.Sleep(2000);
                    Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeBtnSave));
                }

                Thread.Sleep(4000);
                if (sProjectFile.ToLower().IndexOf("testproject") >= 0)
                    Thread.Sleep(10000);


                // Check Destination Value
                Thread.Sleep(3000);
                // Construct the Grid Cell Element Name
                string cellname_state = "Destination Row 0";
                // Get the Element with the Row Col Coordinates
                AutomationElement aeStateCell = AUIUtilities.FindElementByName(cellname_state, aeGrid);

                if (aeStateCell == null)
                {
                    Console.WriteLine("Find TransportDataGridView aeStateCell failed:" + cellname_state);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView state cell failed:" + cellname_state, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("state cell TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                Thread.Sleep(3000);
                //System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeStateCell);
                //System.Windows.Point point = AUICommon.GetDataGridViewCellPointAt(0, "Id", aeGrid);
                // find cell value
                string StateValue = string.Empty;
                try
                {
                    ValuePattern vp = aeStateCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    StateValue = vp.Current.Value;
                    Console.WriteLine("Get element.Current Value:" + StateValue);
                }
                catch (System.NullReferenceException)
                {
                    StateValue = string.Empty;
                }

                // Check state value
                if (StateValue.Equals(DestID))
                {
                    result = ConstCommon.TEST_PASS;
                    Console.WriteLine(testname + " ---pass --- ");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);

                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = " Transport destination is not changed to " + DestID + " , but:" + StateValue;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                }

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }

            //MessageBox.Show("Check Edit Transport 3?", "Click OK to continue ....", MessageBoxButtons.OK);
        }
        #endregion EditTransport
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region CancelTransport
        public static void CancelTransport(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            AutomationElement aeOverview = null;
            AutomationElement aeGrid = null;
            try
            {
                if (sOnlyUITest)
                    root = EtriccUtilities.GetMainWindow("MainForm");

                AUICommon.ClearDisplayedScreens(root, 2);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {
                    Input.MoveToAndClick(aePanelLink);
                    Thread.Sleep(10000);

                    // Find Transport Overview Window
                    System.Windows.Automation.Condition c = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, TRANSPORT_OVERVIEW_TITLE),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                   );

                    // Find the TransportOverview element.
                    aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeOverview == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    }
                }


                AutomationElement aeRow0 = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find Transport GridView
                    aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                    if (aeGrid == null)
                    {
                        sErrorMessage = "Find TransportDataGridView failed:" + "TransportDataGridView";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Thread.Sleep(3000);
                        // Construct the Grid Cell Element Name
                        string rowname = "Row 0";
                        // Get the Element with the Row Col Coordinates
                        aeRow0 = AUIUtilities.FindElementByName(rowname, aeGrid);
                        if (aeRow0 == null)
                        {
                            sErrorMessage = "Find TransportDataGridView aeRow0 failed:" + rowname;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Thread.Sleep(3000);
                            System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeRow0);
                            //System.Windows.Point point = AUICommon.GetDataGridViewCellPointAt(0, "Id", aeGrid);
                            // find cell value
                            string TransportValue = string.Empty;
                            try
                            {
                                ValuePattern vp = aeRow0.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                TransportValue = vp.Current.Value;
                                Console.WriteLine("Get element.Current Value:" + TransportValue);
                            }
                            catch (System.NullReferenceException)
                            {
                                TransportValue = string.Empty;
                            }

                            if (TransportValue == null || TransportValue == string.Empty)
                            {
                                sErrorMessage = "TransportDataGridView aeRow0 Value not found:" + aeRow0;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                //string TransportValue = AUICommon.GetDataGridViewCellValueAt(0, "Id", aeGrid);
                                Input.MoveToAndRightClick(point);
                                Thread.Sleep(3000);
                            }
                        }
                    }
                }

                AutomationElement aeMenuItemTransport = null; ;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // find Cancel Transport point
                   // System.Windows.Automation.Condition cM = new AndCondition(
                       //new PropertyCondition(AutomationElement.NameProperty, Constants.TRANSPORT_MENUITEM_Cancel_Transport),
                  //     new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                  // );


                    AutomationElementCollection aeAllItems = root.FindAll(TreeScope.Element | TreeScope.Descendants, 
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));
                    for (int i = 0; i < aeAllItems.Count; i++)
                    {
                        Console.WriteLine("MenuItem name: " + aeAllItems[i].Current.Name);
                        if (aeAllItems[i].Current.Name.StartsWith("Cancel "))
                        {

                            aeMenuItemTransport = aeAllItems[i];
                            break;
                        }
                    }

                    // Find the MenuItem Cancel Transport element
                    //AutomationElement aeMenuItemTransport = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
                    if (aeMenuItemTransport != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemTransport.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemTransport.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemTransport.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemTransport.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemTransport);
                        Thread.Sleep(2000);
                    }
                    else
                    {
                        sErrorMessage = "Cancel Transport menu item not found ------------:";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // Find Cancel Transport Dialog Window
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    System.Windows.Automation.Condition c2 = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Cancel Transports?"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                    );

                    // Find the Cancel Dialog element
                    AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                    //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                    if (aeDialog == null)
                    {
                        sErrorMessage = " Dialog Window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Find Yes Button
                        System.Windows.Automation.Condition c3 = new AndCondition(
                           new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                           new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                       );

                        // Find Yes Button element
                        AutomationElement aeYes = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c3);
                        if (aeYes == null)
                        {
                            sErrorMessage = " Yes Button not found";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeYes));
                            Thread.Sleep(5000);
                        }
                    }
                }

                // Check State Value
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    // Construct the Grid Cell Element Name
                    string cellname_state = "State Row 0";
                    // Get the Element with the Row Col Coordinates
                    AutomationElement aeStateCell = AUIUtilities.FindElementByName(cellname_state, aeGrid);
                    if (aeStateCell == null)
                    {
                        sErrorMessage = "Find TransportDataGridView aeStateCell failed:" + cellname_state;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("state cell TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(3000);
                        //System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeStateCell);
                        //System.Windows.Point point = AUICommon.GetDataGridViewCellPointAt(0, "Id", aeGrid);
                        // find cell value
                        string StateValue = string.Empty;
                        try
                        {
                            ValuePattern vp = aeStateCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            StateValue = vp.Current.Value;
                            Console.WriteLine("Get element.Current Value:" + StateValue);
                        }
                        catch (System.NullReferenceException)
                        {
                            StateValue = string.Empty;
                        }

                        // Check state value
                        if (StateValue.Equals("Finished"))
                        {
                            TestCheck = ConstCommon.TEST_PASS;
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = " Transport state is not Finished, but:" + StateValue;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                        }



                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    sErrorMessage = string.Empty;
                    Console.WriteLine(testname);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                }

                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }
        }
        #endregion CancelTransport
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region SuspendTransport
        public static void SuspendTransport(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            TestCheck = ConstCommon.TEST_PASS;
            result = ConstCommon.TEST_UNDEFINED;
            sTransportSuspendOK = true;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            AutomationElement aeOverview = null;
            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 3, 1), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                #region // open transport overview
                AUICommon.ClearDisplayedScreens(root, 2);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {
                    Thread.Sleep(3000);
                    Input.MoveToAndClick(aePanelLink);
                    Thread.Sleep(10000);

                    // Find Transport Overview Window
                    Condition c = new AndCondition(new PropertyCondition(AutomationElement.NameProperty, TRANSPORT_OVERVIEW_TITLE),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                    // Find the TransportOverview element.
                    aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeOverview == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
                        Console.WriteLine(sErrorMessage);
                    }

                }
                #endregion

                AutomationElement aeGrid = null;
                AutomationElement aeRow0 = null;
                Point row0Point = new Point();
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // Find Transport GridView
                    aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                    if (aeGrid == null)
                    {
                        sErrorMessage = "TransportDataGridView not found";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(3000);
                        // Construct the Grid Cell Element Name
                        string rowname = "Row 0";
                        // Get the Element with the Row Col Coordinates
                        aeRow0 = AUIUtilities.FindElementByName(rowname, aeGrid);
                        if (aeRow0 == null)
                        {
                            sErrorMessage = "Find TransportDataGridView aeCell failed:" + rowname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine("cell TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(3000);
                            row0Point = AUIUtilities.GetElementCenterPoint(aeRow0);
                            // find cell value
                            string TransportValue = string.Empty;
                            try
                            {
                                ValuePattern vp = aeRow0.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                TransportValue = vp.Current.Value;
                                Console.WriteLine("Get element.Current Value:" + TransportValue);
                            }
                            catch (System.NullReferenceException)
                            {
                                TransportValue = string.Empty;
                            }

                            if (TransportValue == null || TransportValue == string.Empty)
                            {
                                sErrorMessage = "TransportDataGridView aeCell Value not found:" + rowname;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, "TransportDataGridView cell value not found:" + rowname, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                    }
                    #endregion
                }

                AutomationElement aeMenuItemTransport = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // suspend transport menu action
                    Input.MoveToAndRightClick(row0Point);
                    Thread.Sleep(3000);

                    // find Suspend Transport point
                   // Condition cM = new AndCondition(
                   //    //new PropertyCondition(AutomationElement.NameProperty, Constants.TRANSPORT_MENUITEM_Suspend_Transport),
                   //    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));

                    AutomationElementCollection aeAllItems = root.FindAll(TreeScope.Element | TreeScope.Descendants, 
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));
                    for (int i = 0; i < aeAllItems.Count; i++)
                    {
                        Console.WriteLine("Window name: " + aeAllItems[i].Current.Name);
                        if (aeAllItems[i].Current.Name.StartsWith("Suspend "))
                        {

                            aeMenuItemTransport = aeAllItems[i];
                            break;
                        }
                    }

                    // Find the MenuItem Suspend  Transport element
                    //aeMenuItemTransport = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
                    if (aeMenuItemTransport != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemTransport.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemTransport.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemTransport.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemTransport.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemTransport);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Suspend Transport menu item not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    #endregion
                }

                Thread.Sleep(2000);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find Suspend Transport Dialog Window
                    Condition c2 = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Suspend Transport"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                    // Find the Susepnd Dialog element
                    AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                    //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                    if (aeDialog == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Suspend Transport Dialog Window not found";
                        Console.WriteLine(sErrorMessage);
                        EtriccUtilities.TryToGetErrorMessageAndCloseErrorScreen(ref sErrorMessage);
                    }
                    else
                    {
                        // Find Yes Button
                        System.Windows.Automation.Condition c3 = new AndCondition(
                           new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                           new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                       );

                        // Find Yes Button element
                        AutomationElement aeYes = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c3);
                        if (aeYes == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = " Yes Button not found";
                            Console.WriteLine("FindElementByID failed: Yes button");
                        }
                        else
                        {
                            Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeYes));
                            Thread.Sleep(5000);

                            Input.MoveToAndRightClick(row0Point);
                            Thread.Sleep(3000);
                        }
                    }
                }

                // validation Suspending
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // suspend transport menu action
                    Input.MoveToAndRightClick(row0Point);
                    Thread.Sleep(3000);

                    // find Suspend Transport point
                    Condition cM = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, Constants.TRANSPORT_MENUITEM_Edit_Transport),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));

                    // Find the MenuItem Cancel Transport element
                    aeMenuItemTransport = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
                    if (aeMenuItemTransport != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemTransport.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemTransport.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemTransport.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemTransport.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemTransport);
                    }
                    else
                    {
                        sErrorMessage = "Edit Transport menu item not found ------------:";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    #endregion
                }

                // Check suspend check box
                AutomationElement aeEditTransportWindow = null;
                AutomationElement aeSuspendCheckBox = null;
                // validation Suspending in Edit transport Screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                    if (aeWindow == null)
                    {
                        sErrorMessage = "Main Window not found";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        //string editId = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.AddEditTransport";
                        string editIdNew = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.AddEditTransport.AddEditTransport";
                        string editIdOld = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.AddEditTransport";
                        aeEditTransportWindow = AUIUtilities.FindElementByID(editIdNew, aeWindow);
                        if (aeEditTransportWindow == null)
                        {
                            aeEditTransportWindow = AUIUtilities.FindElementByID(editIdOld, aeWindow);
                        }
                        
                        if (aeEditTransportWindow == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = " aeEditTransportWindow not found";
                            Console.WriteLine(sErrorMessage);
                        }
                        else
                        {
                            aeSuspendCheckBox = AUIUtilities.FindElementByID("checkBoxSuspended", aeWindow);
                            if (aeSuspendCheckBox == null)
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = " aeSuspendCheckBox not found";
                                Console.WriteLine(sErrorMessage);
                            }
                            else
                            {
                                // check checkbox state
                                TogglePattern tg = aeSuspendCheckBox.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                Thread.Sleep(1500);
                                ToggleState tgState = tg.Current.ToggleState;
                                Console.WriteLine("ToggleState : " + tgState.ToString());

                                if (tgState == ToggleState.On)
                                {
                                    TestCheck = ConstCommon.TEST_PASS;
                                }
                                else
                                {
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    sErrorMessage = " aeSuspendCheckBox not checked";
                                    Console.WriteLine(sErrorMessage);
                                }

                                AutomationElement aeClose = AUIUtilities.FindElementByID("m_btnCancel", aeEditTransportWindow);
                                Input.MoveToAndClick(aeClose);
                                Thread.Sleep(3000);
                            }
                        }

                    }

                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    sTransportSuspendOK = false;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }

            //MessageBox.Show("Check Suspend Transport?", "Click OK to continue ....", MessageBoxButtons.OK);
        }
        #endregion SuspendTransport
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ReleaseTransport
        public static void ReleaseTransport(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            TestCheck = ConstCommon.TEST_PASS;
            result = ConstCommon.TEST_UNDEFINED;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            if (sTransportSuspendOK == false)
            {
                sErrorMessage = "Transport suspend test failed, Transport release cannot be tested";
                return;
            }

            AutomationElement aePanelLink = null;
            AutomationElement aeOverview = null;
            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 3, 1), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                #region // open transport overview
                AUICommon.ClearDisplayedScreens(root, 2);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {
                    Thread.Sleep(3000);
                    Input.MoveToAndClick(aePanelLink);
                    Thread.Sleep(10000);

                    // Find Transport Overview Window
                    Condition c = new AndCondition(new PropertyCondition(AutomationElement.NameProperty, TRANSPORT_OVERVIEW_TITLE),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                    // Find the TransportOverview element.
                    aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeOverview == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
                        Console.WriteLine(sErrorMessage);
                    }
                }
                #endregion

                AutomationElement aeGrid = null;
                AutomationElement aeRow0 = null;
                Point cellPoint = new Point();
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // Find Transport GridView
                    aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                    if (aeGrid == null)
                    {
                        sErrorMessage = "TransportDataGridView not found";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(3000);
                        // Construct the Grid Cell Element Name
                        string rowname = "Row 0";
                        // Get the Element with the Row Col Coordinates
                        aeRow0 = AUIUtilities.FindElementByName(rowname, aeGrid);
                        if (aeRow0 == null)
                        {
                            sErrorMessage = "Find TransportDataGridView aeRow0 failed:" + rowname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine("aeRow0 TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(3000);
                            cellPoint = AUIUtilities.GetElementCenterPoint(aeRow0);
                            // find cell value
                            string TransportValue = string.Empty;
                            try
                            {
                                ValuePattern vp = aeRow0.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                TransportValue = vp.Current.Value;
                                Console.WriteLine("Get element.Current Value:" + TransportValue);
                            }
                            catch (System.NullReferenceException)
                            {
                                TransportValue = string.Empty;
                            }

                            if (TransportValue == null || TransportValue == string.Empty)
                            {
                                sErrorMessage = "TransportDataGridView aeRow0 Value not found:" + rowname;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, "TransportDataGridView aeRow0 value not found:" + rowname, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                    }
                    #endregion
                }

                AutomationElement aeMenuItemTransport = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // release transport menu action
                    Input.MoveToAndRightClick(cellPoint);
                    Thread.Sleep(3000);

                    // find Release Transport point
                    //Condition cM = new AndCondition(
                       //new PropertyCondition(AutomationElement.NameProperty, Constants.TRANSPORT_MENUITEM_Release_Transport),
                     //  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));

                    AutomationElementCollection aeAllItems = root.FindAll(TreeScope.Element | TreeScope.Descendants, 
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));
                    for (int i = 0; i < aeAllItems.Count; i++)
                    {
                        Console.WriteLine("MenuItem name: " + aeAllItems[i].Current.Name);
                        if (aeAllItems[i].Current.Name.StartsWith("Release "))
                        {

                            aeMenuItemTransport = aeAllItems[i];
                            break;
                        }
                    }

                    // Find the MenuItem Cancel Transport element
                    //aeMenuItemTransport = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
                    if (aeMenuItemTransport != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemTransport.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemTransport.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemTransport.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemTransport.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemTransport);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Release Transport menu item not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    #endregion
                }

                Thread.Sleep(2000);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Find Release Transport Dialog Window
                    Condition c2 = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "Release Transports?"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                    // Find the Susepnd Dialog element
                    AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                    //AutomationElement aeDialog = AUIUtilities.FindElementByID(MESSAGESCREEN_ID, root);
                    if (aeDialog == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = " Dialog Window not found";
                        Console.WriteLine("FindElementByID failed: Dialog");
                        EtriccUtilities.TryToGetErrorMessageAndCloseErrorScreen(ref sErrorMessage);
                    }
                    else
                    {
                        // Find Yes Button
                        System.Windows.Automation.Condition c3 = new AndCondition(
                           new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                           new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                       );

                        // Find Yes Button element
                        AutomationElement aeYes = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c3);
                        if (aeYes == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = " Yes Button not found";
                            Console.WriteLine("FindElementByID failed: Yes button");
                        }
                        else
                        {
                            Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeYes));
                            Thread.Sleep(5000);

                            Input.MoveToAndRightClick(cellPoint);
                            Thread.Sleep(2000);
                        }
                    }
                }

                // validation Release
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // edit transport menu action
                    Input.MoveToAndRightClick(cellPoint);
                    Thread.Sleep(3000);

                    // find Suspend Transport point
                    Condition cM = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, Constants.TRANSPORT_MENUITEM_Edit_Transport),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));

                    // Find the MenuItem Cancel Transport element
                    aeMenuItemTransport = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
                    if (aeMenuItemTransport != null)
                    {
                        Console.WriteLine("new element found: " + aeMenuItemTransport.Current.Name);
                        Console.WriteLine("err.Current.AutomationId: " + aeMenuItemTransport.Current.AutomationId);
                        Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemTransport.Current.ControlType.ToString());
                        Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemTransport.Current.ControlType.ProgrammaticName);
                        Input.MoveToAndClick(aeMenuItemTransport);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Edit Transport menu item not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    #endregion
                }

                // Check suspend check check box
                AutomationElement aeEditTransportWindow = null;
                AutomationElement aeSuspendCheckBox = null;
                // validation Suspending in Edit transport Screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    AutomationElement aeWindow = EtriccUtilities.GetMainWindow("MainForm");
                    if (aeWindow == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Main Window not found";
                        Console.WriteLine(sErrorMessage);
                    }
                    else
                    {
                        string editIdNew = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.AddEditTransport.AddEditTransport";
                        string editIdOld = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.AddEditTransport";
                        aeEditTransportWindow = AUIUtilities.FindElementByID(editIdNew, aeWindow);
                        if (aeEditTransportWindow == null)
                        {
                            aeEditTransportWindow = AUIUtilities.FindElementByID(editIdOld, aeWindow);
                        }


                        if (aeEditTransportWindow == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = " aeEditTransportWindow not found";
                            Console.WriteLine(sErrorMessage);
                        }
                        else
                        {
                            aeSuspendCheckBox = AUIUtilities.FindElementByID("checkBoxSuspended", aeWindow);
                            if (aeSuspendCheckBox == null)
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = " aeSuspendCheckBox not found";
                                Console.WriteLine(sErrorMessage);
                            }
                            else
                            {
                                // check checkbox state
                                TogglePattern tg = aeSuspendCheckBox.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                Thread.Sleep(1500);
                                ToggleState tgState = tg.Current.ToggleState;
                                Console.WriteLine("ToggleState : " + tgState.ToString());

                                if (tgState == ToggleState.Off)
                                {
                                    TestCheck = ConstCommon.TEST_PASS;
                                }
                                else
                                {
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    sErrorMessage = " aeSuspendCheckBox not checked";
                                    Console.WriteLine(sErrorMessage);
                                }

                                AutomationElement aeClose = AUIUtilities.FindElementByID("m_btnCancel", aeEditTransportWindow);
                                Input.MoveToAndClick(aeClose);
                                Thread.Sleep(3000);
                            }
                        }

                    }

                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }

            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
            }
            //MessageBox.Show("Check Release Transport?", "Click OK to continue ....", MessageBoxButtons.OK);

        }
        #endregion ReleaseTransport
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region TransportOverviewOpenDetail
        public static void TransportOverviewOpenDetail(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            AutomationElement aePanelLink = null;
            try
            {
                AUICommon.ClearDisplayedScreens(root, 2);

                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Input.MoveToAndClick(aePanelLink);

                Thread.Sleep(10000);

                // Find Transport Overview Window Element
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, TRANSPORT_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );

                AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeOverview == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
                    Console.WriteLine("FindElementByID failed:" + TRANSPORT_OVERVIEW);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }

                AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                if (aeGrid == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine("Find TransportDataGridView failed:" + "TransportDataGridView");
                    Epia3Common.WriteTestLogFail(slogFilePath, "Find TransportDataGridView failed:" + "TransportDataGridView", sOnlyUITest);

                }
                else
                    Console.WriteLine("Button TransportataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                ////--------------------------
                Thread.Sleep(3000);
                // Construct the Grid Cell Element Name
                string rowname = "Row 0";
                // Get the Element with the Row Col Coordinates
                AutomationElement aeRow0 = AUIUtilities.FindElementByName(rowname, aeGrid);

                if (aeRow0 == null)
                {
                    Console.WriteLine("Find TransportDataGridView aeCell failed:" + rowname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView cell failed:" + rowname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("cell TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                Thread.Sleep(3000);
                System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeRow0);

                // Open Detail Screen
                Input.MoveToAndDoubleClick(point);
                Thread.Sleep(2000);

                // find transport detail screen                
                string cellValue = AUICommon.GetDataGridViewCellValueAt(0, "Id", aeGrid);
                Console.WriteLine("Transport Id value: " + cellValue);
                Epia3Common.WriteTestLogMsg(slogFilePath, "Id cell value: " + cellValue, sOnlyUITest);

                string TrnScreenName = "Transport detail - " + cellValue;
                Console.WriteLine("TrnScreenName: " + TrnScreenName);
                Epia3Common.WriteTestLogMsg(slogFilePath, "TrnScreenName: " + TrnScreenName, sOnlyUITest);
                // Find All Windows
                System.Windows.Automation.Condition c2 = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.Window);

                // Find All Windows.
                AutomationElementCollection aeAllWindows = root.FindAll(TreeScope.Element | TreeScope.Descendants, c2);
                Thread.Sleep(3000);

                AutomationElement aeDetailScreen = null;
                bool isOpen = false;
                string windowName = "No Name";
                for (int i = 0; i < aeAllWindows.Count; i++)
                {
                    windowName = aeAllWindows[i].Current.Name;
                    Console.WriteLine("Window name: " + windowName);
                    if (windowName.StartsWith(TrnScreenName))
                    {
                        isOpen = true;
                        aeDetailScreen = aeAllWindows[i];
                        break;
                    }
                }

                if (!isOpen)
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine("Find TransportDetailView failed:" + "TransportDetailView");
                    Epia3Common.WriteTestLogFail(slogFilePath, "Find TransportDetailView failed:" + "TransportDetailView", sOnlyUITest);
                    return;
                }

                // Check Transport ID Value
                string textID = "m_IdValueLabel";
                AutomationElement aeTrnText = AUIUtilities.FindElementByID(textID, aeDetailScreen);
                if (aeTrnText == null)
                {
                    sErrorMessage = "Find TrnTextElement failed:" + textID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                string TrnTextValue = aeTrnText.Current.Name;
                if (TrnTextValue.Equals(cellValue))
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = TrnScreenName + "Trn Value should be " + cellValue + ", but  " + TrnTextValue;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ScriptBeforeAfterActivate
        public static void ScriptBeforeAfterActivate(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 5, 11), ref sErrorMessage) == true)
            {
                Console.WriteLine(sErrorMessage);
                result = ConstCommon.TEST_UNDEFINED;
                return;
            }


            sBeforeAfterActivateScriptOK = true;
            string EtriccServerPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server";
            try
            {
                string EtriccServerDataPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server\Data\Etricc";
                string EtriccCurrentDataPath = @"C:\EtriccTests\EtriccUI\BeforeAfterActivateScripts";
                if (!EtriccUtilities.CopyFilesWithWildcards(EtriccCurrentDataPath, EtriccServerDataPath, "Test*.cs"))
                {
                    sErrorMessage = "Copy scripts failed";
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;

                }

                Thread.Sleep(5000);

                #region stop etricc service
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Stop ETRICC SERVER as Service : ");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Stop ETRICC SERVER as Service : "
                        + EtriccServerPath + " - " + ConstCommon.EGEMIN_ETRICC_SERVER_EXE, sOnlyUITest);
                    TestTools.ProcessUtilities.StartProcessWaitForExit(EtriccServerPath,
                        ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /stop");
                    Thread.Sleep(2000);

                    ServiceController svcEtricc = new ServiceController("Egemin Etricc Server");
                    Console.WriteLine(svcEtricc.ServiceName + " has status " + svcEtricc.Status.ToString());
                    Thread.Sleep(2000);
                    //svcEtricc.WaitForStatus(ServiceControllerStatus.Stopped);
                    // wait until etricc service status is Stoped
                    sStartTime = DateTime.Now;
                    sTime = DateTime.Now - sStartTime;
                    int wait = 0;
                    string etriccServiceStatus = svcEtricc.Status.ToString().ToLower();
                    Console.WriteLine(" time is :" + sTime.TotalSeconds);
                    while (!etriccServiceStatus.StartsWith("stopped") && wait < 150)
                    {
                        Thread.Sleep(2000);
                        wait = wait + 2;
                        etriccServiceStatus = svcEtricc.Status.ToString().ToLower();
                        Console.WriteLine("wait etricc service status stopped:  time is (sec) : " + wait + "  and status is:" + etriccServiceStatus);
                    }

                    if (svcEtricc.Status != ServiceControllerStatus.Stopped)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Etricc Service stop up failed: " + etriccServiceStatus, sOnlyUITest);
                        throw new Exception("Etricc service stop up failed:"); //   get message from log file sErrorMessage//
                    }

                    Console.WriteLine("ETRICC SERVER Service Stopped : ");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "ETRICC SERVER Service Stopped:", sOnlyUITest);
                    Thread.Sleep(2000);
                }
                #endregion

                string configFile = Path.Combine(EtriccServerPath, "Egemin.Etricc.Server.exe.config");
                string configFileBackup = Path.Combine(Directory.GetCurrentDirectory(), "Egemin.Etricc.Server.exe.config");
                Console.WriteLine("configFileBackup:" + configFileBackup);
                System.IO.File.Copy(configFile, configFileBackup, true);


                #region update etricc config file
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    try
                    {
                        var xDoc = new XmlDocument();
                        //xDoc.Load("C:\\Etricc\\Server\\Egemin.Etricc.Server.exe.config");
                        string EtriccPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server";
                        xDoc.Load(Path.Combine(EtriccPath, "Egemin.Etricc.Server.exe.config"));

                        var xPathNav = xDoc.CreateNavigator();
                        xPathNav.MoveToFirstChild();
                        xPathNav.MoveToFirstChild();
                        while (!xPathNav.LocalName.StartsWith("epia.componentconfiguration"))
                        {
                            xPathNav.MoveToNext();
                        }

                        xPathNav.MoveToFirstChild();

                        //while (!xPathNav.LocalName.Equals("parameter"))
                        while (!xPathNav.LocalName.Equals("parameters"))
                        {
                            xPathNav.MoveToFirstChild();
                        }

                        xPathNav.AppendChild("<parameter name=\"BeforeActivateScript\" value=\".\\Data\\Etricc\\TestBeforeActivateScript.cs\" />");
                        xPathNav.AppendChild("<parameter name=\"AfterActivateScript\" value=\".\\Data\\Etricc\\TestAfterActivateScript.cs\" />");
                        xPathNav.AppendChild("<parameter name=\"BeforeDeactivateScript\" value=\".\\Data\\Etricc\\TestBeforeDeactivateScript.cs\" />");
                        xPathNav.AppendChild("<parameter name=\"AfterDeactivateScript\" value=\".\\Data\\Etricc\\TestAfterDeactivateScript.cs\" />");

                        xDoc.Save(Path.Combine(EtriccPath, "Egemin.Etricc.Server.exe.config"));
                    }
                    catch (Exception ex)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = testname + " = update config file == " + ex.Message + "------" + ex.StackTrace;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    }
                    finally
                    {
                        Thread.Sleep(3000);
                    }
                }
                #endregion

                #region start etricc service
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Start ETRICC SERVER as Service : ");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Star ETRICC SERVER as Service : "
                        + EtriccServerPath + " - " + ConstCommon.EGEMIN_ETRICC_SERVER_EXE, sOnlyUITest);
                    TestTools.ProcessUtilities.StartProcessWaitForExit(EtriccServerPath,
                        ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /start");
                    Thread.Sleep(2000);

                    ServiceController svcEtricc = new ServiceController("Egemin Etricc Server");
                    Console.WriteLine(svcEtricc.ServiceName + " has status " + svcEtricc.Status.ToString());
                    Thread.Sleep(2000);
                    //svcEtricc.WaitForStatus(ServiceControllerStatus.Running);
                    // wait until etricc service status is Running
                    sStartTime = DateTime.Now;
                    sTime = DateTime.Now - sStartTime;
                    int wait = 0;
                    string etriccServiceStatus = svcEtricc.Status.ToString().ToLower();
                    Console.WriteLine(" time is :" + sTime.TotalSeconds);
                    while (!etriccServiceStatus.StartsWith("running") && wait < 150)
                    {
                        Thread.Sleep(2000);
                        wait = wait + 2;
                        etriccServiceStatus = svcEtricc.Status.ToString().ToLower();
                        Console.WriteLine("wait etricc service status running:  time is (sec) : " + wait + "  and status is:" + etriccServiceStatus);
                    }

                    if (svcEtricc.Status != ServiceControllerStatus.Running)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Etricc Service start up failed: " + etriccServiceStatus, sOnlyUITest);
                        throw new Exception("Etricc service start up failed:"); //   get message from log file sErrorMessage//
                    }

                    Console.WriteLine("ETRICC SERVER Service Started : ");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "ETRICC SERVER Service Started:", sOnlyUITest);
                    Thread.Sleep(2000);

                }
                #endregion

                #region  check location overviewv screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    root = EtriccUtilities.GetMainWindow("MainForm");
                    AutomationElement aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, LOCATION_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Input.MoveToAndClick(aePanelLink);
                    }

                    Thread.Sleep(10000);
                    // Find Location Overview Window
                    System.Windows.Automation.Condition c = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, LOCATION_OVERVIEW_TITLE),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                    AutomationElement aeGrid = null;
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        // Find the LocationOverview element.
                        AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                        if (aeOverview == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = LOCATION_OVERVIEW + " Window not found";
                            Console.WriteLine("FindElementByID failed:" + "Locations");
                        }
                        else
                        {
                            // Find Location GridView
                            aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                            if (aeGrid == null)
                            {
                                sErrorMessage = "Find LocationDataGridView failed:" + "LocationDataGridView";
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                    }

                    // find cellelement FLV_L5    PRK_FLV
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        bool validate = EtriccUtilities.ValidateGridData(aeGrid, "Id", "FLV_L5", "Mode", "Disabled", 14, ref sErrorMessage);
                        if (validate == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            validate = EtriccUtilities.ValidateGridData(aeGrid, "Id", "PRK_FLV", "Mode", "Disabled", 14, ref sErrorMessage);
                            if (validate == false)
                            {
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }

                    }
                }
                #endregion

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    sBeforeAfterActivateScriptOK = false;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                Thread.Sleep(3000);
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
            }
        }
        #endregion ScriptBeforeAfterDeactivate
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ScriptBeforeAfterDeactivate
        public static void ScriptBeforeAfterDeactivate(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            string EtriccServerPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server";
            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }

            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 5, 11), ref sErrorMessage) == true)
            {
                Console.WriteLine(sErrorMessage);
                result = ConstCommon.TEST_UNDEFINED;
                return;
            }

            if (sOnlyUITest)
                sBeforeAfterActivateScriptOK = true;

            if (sBeforeAfterActivateScriptOK == false)
            {
                sErrorMessage = "BeforeAfterActivateScript test failed, ScriptBeforeAfterDeactivate cannot be tested";
                return;
            }

            try
            {

                #region stop etricc service
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Stop ETRICC SERVER as Service : "
                        + EtriccServerPath + " - " + ConstCommon.EGEMIN_ETRICC_SERVER_EXE);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Stop ETRICC SERVER as Service : "
                        + EtriccServerPath + " - " + ConstCommon.EGEMIN_ETRICC_SERVER_EXE, sOnlyUITest);
                    TestTools.ProcessUtilities.StartProcessWaitForExit(EtriccServerPath,
                        ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /stop");
                    Thread.Sleep(2000);

                    ServiceController svcEtricc = new ServiceController("Egemin Etricc Server");
                    Console.WriteLine(svcEtricc.ServiceName + " has status " + svcEtricc.Status.ToString());
                    Thread.Sleep(2000);
                    //svcEtricc.WaitForStatus(ServiceControllerStatus.Stopped);
                    // wait until etricc service status is Stoped
                    sStartTime = DateTime.Now;
                    sTime = DateTime.Now - sStartTime;
                    int wait = 0;
                    string etriccServiceStatus = svcEtricc.Status.ToString().ToLower();
                    Console.WriteLine(" time is :" + sTime.TotalSeconds);
                    while (!etriccServiceStatus.StartsWith("stopped") && wait < 150)
                    {
                        Thread.Sleep(2000);
                        wait = wait + 2;
                        etriccServiceStatus = svcEtricc.Status.ToString().ToLower();
                        Console.WriteLine("wait etricc service status stopped:  time is (sec) : " + wait + "  and status is:" + etriccServiceStatus);
                    }

                    if (svcEtricc.Status != ServiceControllerStatus.Stopped)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Etricc Service stop up failed: " + etriccServiceStatus, sOnlyUITest);
                        throw new Exception("Etricc service stop up failed:"); //   get message from log file sErrorMessage//
                    }

                    Console.WriteLine("ETRICC SERVER Service Stopped : ");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "ETRICC SERVER Service Stopped:", sOnlyUITest);
                    Thread.Sleep(2000);
                }
                #endregion

                string configFile2 = Path.Combine(EtriccServerPath, "Egemin.Etricc.Server.exe.config");
                string configFileBackup2 = Path.Combine(Directory.GetCurrentDirectory(), "Egemin.Etricc.Server.exe.config");
                System.IO.File.Copy(configFileBackup2, configFile2, true);

                /*
              #region update etricc config file
              var xDoc2 = new XmlDocument();
              //xDoc.Load("C:\\Etricc\\Server\\Egemin.Etricc.Server.exe.config");
              string EtriccPath2 = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server";
              xDoc2.Load(Path.Combine(EtriccPath2, "Egemin.Etricc.Server.exe.config"));

              var xPathNav2 = xDoc2.CreateNavigator();
              xPathNav2.MoveToFirstChild();
              xPathNav2.MoveToFirstChild();
              while (!xPathNav2.LocalName.StartsWith("epia.componentconfiguration"))
              {
                  xPathNav2.MoveToNext();
              }

              xPathNav2.MoveToFirstChild();
              //while (!xPathNav.LocalName.Equals("parameter"))
              while (!xPathNav2.LocalName.Equals("parameters"))
              {
                  xPathNav2.MoveToFirstChild();
                  //System.Windows.Forms.MessageBox.Show("xPathNav2.GetAttribute(name)=" + xPathNav2.GetAttribute("name", ""), xPathNav2.LocalName);
              }
                
              while (xPathNav2.LocalName.Equals("parameters"))
              {
                  xPathNav2.MoveToFirstChild();
                  //System.Windows.Forms.MessageBox.Show("2xPathNav2.GetAttribute(name)=" + xPathNav2.GetAttribute("name", ""), xPathNav2.LocalName);
                  while (xPathNav2.LocalName.Equals("parameter"))
                  {
                      //System.Windows.Forms.MessageBox.Show("xPathNav2.GetAttribute(name)=" + xPathNav2.GetAttribute("name", ""));
                      if (xPathNav2.GetAttribute("name", "").ToLower().IndexOf("script") > 0)
                      {
                          //System.Windows.Forms.MessageBox.Show("xPathNav2.GetAttribute(name)=" + xPathNav2.GetAttribute("name", ""));
                          xPathNav2.DeleteSelf();
                      }
                      xPathNav2.MoveToNext();
                      //System.Windows.Forms.MessageBox.Show("CCCCCCCCCCPathNav2.GetAttribute(name)=" + xPathNav2.GetAttribute("name", ""), xPathNav2.LocalName);
                      //XmlNode fieldNode = xDoc2.SelectSingleNode(@"/root/field[@name = '3']");
                      xPathNav2.MoveToNext();
                      //System.Windows.Forms.MessageBox.Show("DDDDDDDDDDDPathNav2.GetAttribute(name)=" + xPathNav2.GetAttribute("name", ""), xPathNav2.LocalName);
                  }
              }
              xDoc2.Save(Path.Combine(EtriccPath2, "Egemin.Etricc.Server.exe.config"));
              xDoc2.Save(Path.Combine(slogFilePath, "Egemin.Etricc.Server.exe.config"));
              #endregion
              */
                #region start etricc service
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Start ETRICC SERVER as Service : "
                        + EtriccServerPath + " - " + ConstCommon.EGEMIN_ETRICC_SERVER_EXE);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Star ETRICC SERVER as Service : "
                        + EtriccServerPath + " - " + ConstCommon.EGEMIN_ETRICC_SERVER_EXE, sOnlyUITest);
                    TestTools.ProcessUtilities.StartProcessWaitForExit(EtriccServerPath,
                        ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /start");
                    Thread.Sleep(2000);

                    ServiceController svcEtricc = new ServiceController("Egemin Etricc Server");
                    Console.WriteLine(svcEtricc.ServiceName + " has status " + svcEtricc.Status.ToString());
                    Thread.Sleep(2000);
                    //svcEtricc.WaitForStatus(ServiceControllerStatus.Running);
                    // wait until etricc service status is Running
                    sStartTime = DateTime.Now;
                    sTime = DateTime.Now - sStartTime;
                    int wait = 0;
                    string etriccServiceStatus = svcEtricc.Status.ToString().ToLower();
                    Console.WriteLine(" time is :" + sTime.TotalSeconds);
                    while (!etriccServiceStatus.StartsWith("running") && wait < 150)
                    {
                        Thread.Sleep(2000);
                        wait = wait + 2;
                        etriccServiceStatus = svcEtricc.Status.ToString().ToLower();
                        Console.WriteLine("wait etricc service status running:  time is (sec) : " + wait + "  and status is:" + etriccServiceStatus);
                    }

                    if (svcEtricc.Status != ServiceControllerStatus.Running)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Etricc Service start up failed: " + etriccServiceStatus, sOnlyUITest);
                        throw new Exception("Etricc service start up failed:"); //   get message from log file sErrorMessage//
                    }

                    Console.WriteLine("ETRICC SERVER Service Started : ");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "ETRICC SERVER Service Started:", sOnlyUITest);
                    Thread.Sleep(2000);

                }
                #endregion

                #region  check location overviewv screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    root = EtriccUtilities.GetMainWindow("MainForm");
                    AutomationElement aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, LOCATION_OVERVIEW, ref sErrorMessage);
                    if (aePanelLink == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Input.MoveToAndClick(aePanelLink);
                    }

                    Thread.Sleep(10000);
                    // Find Location Overview Window
                    System.Windows.Automation.Condition c = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, LOCATION_OVERVIEW_TITLE),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                    AutomationElement aeGrid = null;
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        // Find the LocationOverview element.
                        AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                        if (aeOverview == null)
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = LOCATION_OVERVIEW + " Window not found";
                            Console.WriteLine("FindElementByID failed:" + "Locations");
                        }
                        else
                        {
                            // Find Location GridView
                            aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
                            if (aeGrid == null)
                            {
                                sErrorMessage = "Find LocationDataGridView failed:" + "LocationDataGridView";
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                    }

                    // find cellelement FLV_L5    PRK_FLV
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        bool validate = EtriccUtilities.ValidateGridData(aeGrid, "Id", "FLV_L5", "Mode", "Automatic", 14, ref sErrorMessage);
                        if (validate == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            validate = EtriccUtilities.ValidateGridData(aeGrid, "Id", "PRK_FLV", "Mode", "Automatic", 14, ref sErrorMessage);
                            if (validate == false)
                            {
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }

                    }
                }
                #endregion

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
                Thread.Sleep(3000);
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
            }

        }
        #endregion ScriptBeforeAfterActivate
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Epia4Close
        public static void Epia3Close(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            string BtnCloseID = "Close";
            sErrorMessage = string.Empty;

            if (sEtriccServerStartupOK == false)
            {
                sErrorMessage = "Etricc Server startup failed, this testcase cannot be tested";
                return;
            }
            try
            {
                //AUICommon.ClearDisplayedScreens(root);
                Thread.Sleep(5000);
                Console.WriteLine(testname + ": try to find  aeClose: " + System.DateTime.Now.ToString("HH:mm:ss"));
                AutomationElement aeClose = AUIUtilities.FindElementByID(BtnCloseID, root);
                if (aeClose == null)
                {
                    Console.WriteLine(testname + " failed to find aeClose at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine(testname + " aeClose found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    Input.MoveTo(aeClose);
                }

                Thread.Sleep(2000);

                InvokePattern ip =
                   aeClose.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                ip.Invoke();

                Thread.Sleep(10000);

                System.Diagnostics.Process proc = null;
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
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
        }
        #endregion Epia3Close
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region EtriccExplorerOverview
        public static void EtriccExplorerOverview(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            AutomationEventHandler UIOpenExplorerEventHandler = new AutomationEventHandler(OnOpenExplorerEvent);
            string path = m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server";
            try
            {
                //========================   Explorer =================================================
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                  AutomationElement.RootElement, TreeScope.Descendants, UIOpenExplorerEventHandler);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    TestTools.ProcessUtilities.CloseProcess("EPIA.Explorer");
                    Thread.Sleep(2000);
                    sEventEnd = false;
                    sErrorMessage = string.Empty;
                    TestTools.ProcessUtilities.StartProcessNoWait(path, ConstCommon.EGEMIN_ETRICC_EXPLORER_EXE, string.Empty);
                }
                Console.WriteLine("Start Etricc Explorer ......");

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Application is started : ");

                    sStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - sStartTime;
                    aeForm = null;
                    string smsg = GetErrorMessage();
                    while (aeForm == null && mTime.TotalSeconds <= 120)
                    {
                        aeForm = EtriccUtilities.GetMainWindow("frmMain", 30);
                        Thread.Sleep(2000);
                        mTime = DateTime.Now - sStartTime;
                    }

                    if (aeForm == null)
                    {
                        smsg = GetErrorMessage();
                        if (smsg.Length > 20)
                        {
                            sErrorMessage = smsg;
                            TestCheck = ConstCommon.TEST_FAIL;
                            Console.WriteLine("Explorer sErrorMessage. : " + sErrorMessage);
                        }
                        else
                        {
                            AutomationElement aeError = AUIUtilities.FindElementByID("ErrorScreen", AutomationElement.RootElement);
                            if (aeError != null)
                                AUICommon.ErrorWindowHandling(aeError, ref sErrorMessage);
                            else
                                sErrorMessage = "Application Startup failed.";

                            throw new Exception(sErrorMessage);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Explorer aeForm name : " + aeForm.Current.Name);
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    sEtriccExplorerStartupOK = false;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine(testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (System.ComponentModel.Win32Exception iex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = System.IO.Path.Combine(path, ConstCommon.EGEMIN_ETRICC_EXPLORER_EXE) +" not found" + iex.Message + "  ---  " + iex.StackTrace;
                Console.WriteLine(testname + " === " + System.IO.Path.Combine(path, ConstCommon.EGEMIN_ETRICC_EXPLORER_EXE) + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
                return;
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "  ---  " + ex.StackTrace;
                Console.WriteLine(testname + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
                return;
            }
            finally
            {
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                          AutomationElement.RootElement, UIOpenExplorerEventHandler);
                Thread.Sleep(3000);
                //Input.SendKeyboardInput(System.Windows.Input.Key.Space, true);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region EtriccExplorerLoadProject
        public static void EtriccExplorerLoadProject(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            AutomationElement aeWindow = null;
            string WindowID = "frmMain";
            string EtriccServerPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server";

            if (sEtriccExplorerStartupOK == false)
            {
                sErrorMessage = "Etricc Explorer startup failed, this testcase cannot be tested";
                return;
            }
            try
            {
                sStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - sStartTime;
                //========================   Explorer =================================================
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Find Explorer Window... : ");
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mAppTime, 120);
                    if (aeWindow == null)
                    {
                        sErrorMessage = "Explorer window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        TransformPattern tranform = aeWindow.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                        if (tranform != null)
                        {
                            tranform.Resize(System.Windows.Forms.SystemInformation.VirtualScreen.Width - System.Windows.Forms.SystemInformation.VirtualScreen.Width*0.3,
                                System.Windows.Forms.SystemInformation.VirtualScreen.Height);
                            Thread.Sleep(1000);
                            tranform.Move(0, 0);
                        }

                        aeWindow.SetFocus();
                        Console.WriteLine("Explorer aeWindow name : " + aeWindow.Current.Name);
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                }

                // open load xmlwindow             
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region
                    string menubarName = "Application";
                    Console.WriteLine("find Explorer Menubar ------------:" + menubarName);

                    System.Windows.Automation.Condition cMenubar = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, menubarName),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuBar)
                    );

                    // Find the manebar.
                    AutomationElement aeMenubar = aeWindow.FindFirst(TreeScope.Descendants, cMenubar);
                    if (aeMenubar != null)
                    {
                        // Find Root Menuitem
                        System.Windows.Automation.Condition cRoot = new AndCondition(
                           new PropertyCondition(AutomationElement.NameProperty, "Root"),
                           new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                        );

                        // Find the MenuItem Root
                        AutomationElement aeMenuItemRoot = aeMenubar.FindFirst(TreeScope.Element | TreeScope.Descendants, cRoot);
                        if (aeMenuItemRoot != null)
                        {
                            Input.MoveToAndClick(aeMenuItemRoot);
                            Thread.Sleep(3000);
                            // Find Load XML... Menuitem
                            System.Windows.Automation.Condition cLoadXML = new AndCondition(
                               new PropertyCondition(AutomationElement.NameProperty, "Load XML..."),
                               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                            );

                            AutomationElement aeMenuItemLoadXML = aeMenuItemRoot.FindFirst(TreeScope.Element | TreeScope.Descendants, cLoadXML);
                            if (aeMenuItemLoadXML != null)
                            {
                                Input.MoveToAndClick(aeMenuItemLoadXML);
                            }
                            else
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = "Load XML... menu item not found ------------:";
                                Console.WriteLine(sErrorMessage);
                            }
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "Root menu item not found ------------:";
                            Console.WriteLine(sErrorMessage);
                        }
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Explorer menubar not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    #endregion
                }

                // select project 
                /*AutomationElement aeTypeWiodow = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mAppTime, 300);
                    if (aeWindow == null)
                    {
                        sErrorMessage = "Explorer window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("Explorer window found ------------:");
                        Thread.Sleep(300);
                        // Find root type window
                        System.Windows.Automation.Condition cTypeWindow = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Select root type"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                        );

                        sStartTime = DateTime.Now;
                        mTime = DateTime.Now - sStartTime;
                        while (aeTypeWiodow == null && mTime.TotalSeconds <= 120)
                        {
                            aeTypeWiodow = aeWindow.FindFirst(TreeScope.Children, cTypeWindow);
                            mTime = DateTime.Now - sStartTime;
                            Thread.Sleep(500);
                        }

                        if (aeTypeWiodow  != null)
                        {    
                            Console.WriteLine("aeTypeWiodow found ------------:");                                                                                                                                                                                                                         
                        }
                        else
                        {
                            //TestCheck = ConstCommon.TEST_FAIL;
                            //sErrorMessage = "aeTypeWiodow  not found ------------:";
                            //Console.WriteLine(sErrorMessage);
                            Console.WriteLine("aeTypeWiodow NOT found ------------:");   
                        }
                    }                   
                    #endregion
                }
                /*
                // SCROOL TO CORE PROJECT
                AutomationElement aeTreeView = null;
                AutomationElement aeEgeminEpiaWcs = null;
                AutomationElement aeEgemin = null;
                AutomationElement aeProject = null;
                if (aeTypeWiodow != null)    // if Type Window open, process it first 
                //if (TestCheck == ConstCommon.TEST_PASS)
                {
                    string treeViewId = "tvTypes";
                    #region // processs tree item
                    DateTime mTime2 = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(aeTypeWiodow, ref aeTreeView, treeViewId, mTime2, 60);
                    if (aeTreeView == null)
                    {
                        sErrorMessage = "aeTreeView not found name";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("Tree view found  .........");
                        Thread.Sleep(5000);
                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition cTreeItem = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.TreeItem);

                        AutomationElementCollection aeAllItems = aeTreeView.FindAll(TreeScope.Children, cTreeItem);
                        Console.WriteLine("All items count ..." + aeAllItems.Count);
                        for (int i = 0; i < aeAllItems.Count; i++)
                        {
                            try
                            {
                                ExpandCollapsePattern ep = (ExpandCollapsePattern)aeAllItems[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                ep.Collapse();
                                Thread.Sleep(300);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("aeAllItems[i] can not collasped1: " + aeAllItems[i].Current.Name);
                                Console.WriteLine("aeAllItems[i] can not collasped2: " +ex.Message);
                            }
                        }

                        for (int i = 0; i < aeAllItems.Count; i++)
                        {
                            if (aeAllItems[i].Current.Name.Equals("Egemin.EPIA.WCS"))
                            {
                                aeEgeminEpiaWcs = aeAllItems[i];
                                try
                                {
                                    ExpandCollapsePattern ep = (ExpandCollapsePattern)aeAllItems[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                    ep.Expand();
                                    Thread.Sleep(300);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("aeAllItems[i] can not expanded3: " + aeAllItems[i].Current.Name);
                                    Console.WriteLine("aeAllItems[i] can not expanded4: " + ex.Message);
                                }
                            }
                        }

                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition cItems = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.TreeItem);

                        aeAllItems = aeEgeminEpiaWcs.FindAll(TreeScope.Descendants, cItems);

                        for (int i = 0; i < aeAllItems.Count; i++)
                        {
                            if (aeAllItems[i].Current.Name.Equals("Egemin"))
                            {
                                aeEgemin = aeAllItems[i];
                                try
                                {
                                    ExpandCollapsePattern ep = (ExpandCollapsePattern)aeAllItems[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                    ep.Expand();
                                    Thread.Sleep(300);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("aeAllItems[i] can not expanded5: " + aeAllItems[i].Current.Name);
                                    Console.WriteLine("aeAllItems[i] can not expanded6: " + ex.Message);
                                }
                            }
                        }

                        aeAllItems = aeEgemin.FindAll(TreeScope.Descendants, cItems);
                        AutomationElement aeEPIA = null;
                        for (int i = 0; i < aeAllItems.Count; i++)
                        {
                            if (aeAllItems[i].Current.Name.Equals("EPIA"))
                            {
                                aeEPIA = aeAllItems[i];
                                try
                                {
                                    ExpandCollapsePattern ep = (ExpandCollapsePattern)aeAllItems[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                    ep.Expand();
                                    Thread.Sleep(300);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("aeAllItems[i] can not expanded7: " + aeAllItems[i].Current.Name);
                                    Console.WriteLine("aeAllItems[i] can not expanded8: " + ex.Message);
                                }
                            }
                        }

                        aeAllItems = aeEPIA.FindAll(TreeScope.Descendants, cItems);
                        AutomationElement aeWCS = null;
                        for (int i = 0; i < aeAllItems.Count; i++)
                        {
                            if (aeAllItems[i].Current.Name.Equals("WCS"))
                            {
                                aeWCS = aeAllItems[i];
                                try
                                {
                                    ExpandCollapsePattern ep = (ExpandCollapsePattern)aeAllItems[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                    ep.Expand();
                                    Thread.Sleep(300);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("aeAllItems[i] can not expanded9: " + aeAllItems[i].Current.Name);
                                    Console.WriteLine("aeAllItems[i] can not expandedA: " + ex.Message);
                                }
                            }
                        }

                        aeAllItems = aeWCS.FindAll(TreeScope.Descendants, cItems);
                        AutomationElement aeCore = null;
                        for (int i = 0; i < aeAllItems.Count; i++)
                        {
                            if (aeAllItems[i].Current.Name.Equals("Core"))
                            {
                                aeCore = aeAllItems[i];
                                try
                                {
                                    ExpandCollapsePattern ep = (ExpandCollapsePattern)aeAllItems[i].GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                    ep.Expand();
                                    Thread.Sleep(300);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("aeAllItems[i] can not expandedB: " + aeAllItems[i].Current.Name);
                                    Console.WriteLine("aeAllItems[i] can not expandedC: " + ex.Message);
                                }
                            }
                        }

                        AutomationElementCollection aeAllLeafItems = aeWCS.FindAll(TreeScope.Descendants, cItems);

                        Console.WriteLine("All items count ..." + aeAllItems.Count);
                        for (int i = 0; i < aeAllLeafItems.Count; i++)
                        {
                            Console.WriteLine("aeAllLeafItems[i].Current.Name: " + aeAllLeafItems[i].Current.Name);
                            if (aeAllLeafItems[i].Current.Name.Equals("Project"))
                            {
                                aeProject = aeAllLeafItems[i];
                                break;
                            }
                        }

                        Thread.Sleep(1000);
                        #region // click OK button
                        if (aeProject == null)
                        {
                            sErrorMessage = " aeProject not found, ";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Input.MoveToAndClick(aeProject);
                            Thread.Sleep(300);
                            AutomationElement aeBtnOK = AUIUtilities.FindElementByID("btnOK", aeTypeWiodow);
                            if (aeBtnOK == null)
                            {
                                sErrorMessage = " aeBtnOK not found, ";
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                Input.MoveToAndClick(aeBtnOK);
                            }
                        }
                        #endregion
                    }
                    #endregion
                }
                */
                // process open folder window   
                AutomationElement aeOpenFolderWiodow = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // find open folder window
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mAppTime, 300);
                    if (aeWindow == null)
                    {
                        sErrorMessage = "Explorer window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("Explorer window found ------------:");
                        Thread.Sleep(5000);
                        // Find root type window
                        System.Windows.Automation.Condition cTypeWindow = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Open"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                        );
                        
                        aeOpenFolderWiodow = aeWindow.FindFirst(TreeScope.Children, cTypeWindow);
                        int kx = 0;
                        while (aeOpenFolderWiodow ==null && kx++ < 20)
                        {
                            Thread.Sleep(5000);
                            Console.WriteLine("try to find OpenFolderWiodow ------------:"+kx);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "try to find OpenFolderWiodow ------------:" + kx, sOnlyUITest);
                            aeOpenFolderWiodow = aeWindow.FindFirst(TreeScope.Children, cTypeWindow);
                        }

                        if (aeOpenFolderWiodow != null)
                        {
                            Console.WriteLine("eOpenFolderWiodow found ------------:");                                                                                                                                     
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "eOpenFolderWiodow  not found ------------:";
                            Console.WriteLine(sErrorMessage);
                        }
                    }                   
                    #endregion
                }

                // find root disk
                AutomationElement aeRootDisk =null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // processs tree item
                    DateTime mTime2 = DateTime.Now;
                    // Find root type window
                    System.Windows.Automation.Condition cTree = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Namespace Tree Control"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree)
                    );
                   
                    AutomationElement  aeTree  = aeOpenFolderWiodow.FindFirst(TreeScope.Descendants, cTree);
                    if (aeTree != null)
                    {
                        Console.WriteLine(" aeTree found ------------:");
                        Thread.Sleep(1000);
                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition cTreeItem = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.TreeItem);

                        // first expande Computer item
                        AutomationElementCollection aeAllItems = aeTree.FindAll(TreeScope.Descendants, cTreeItem);
                        AutomationElement aeComputer = null;
                        Console.WriteLine("All items count ..." + aeAllItems.Count);
                        for (int i = 0; i < aeAllItems.Count; i++)
                        {
                            Console.WriteLine("aeAllLeafItems[i].Current.Name: " + aeAllItems[i].Current.Name);
                            if (aeAllItems[i].Current.Name.IndexOf("Computer") >= 0)
                            {
                                aeComputer = aeAllItems[i];
                                try
                                {
                                    ExpandCollapsePattern ep = (ExpandCollapsePattern)aeComputer.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                    ep.Expand();
                                    Thread.Sleep(1000);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("aeLegendeTreeView can not expanedD: " + aeComputer.Current.Name);
                                    Console.WriteLine("aeLegendeTreeView can not expanedE: " + ex.Message);
                                }
                                break;
                            }
                        }

                        aeAllItems = aeTree.FindAll(TreeScope.Descendants, cTreeItem);
                        Console.WriteLine("All items count ..." + aeAllItems.Count);
                        for (int i = 0; i < aeAllItems.Count; i++)
                        {
                            Console.WriteLine("aeAllLeafItems[i].Current.Name: " + aeAllItems[i].Current.Name);
                            if (aeAllItems[i].Current.Name.IndexOf("C:") >= 0)
                            {
                                aeRootDisk = aeAllItems[i];
                                break;
                            }
                        }

                        if (aeRootDisk == null)
                        {
                            sErrorMessage = "aeRootDisk not found name";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Input.MoveToAndClick(aeRootDisk);
                        }
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = " aeTree  not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    #endregion
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(1000);
                    if ( EtriccUtilities.ScrollToThisFolderItemAndDoubleClick(OSVersionInfoClass.ProgramFilesx86FolderName(), ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + OSVersionInfoClass.ProgramFilesx86FolderName());
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if ( EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("Dematic", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Dematic");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("Etricc Server", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Etricc Server");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("Data", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Data");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("Etricc", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Etricc");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("Demo", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Demo");
                        Thread.Sleep(2000);

                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                //Thread.Sleep(5000);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("TestCkeck : " + TestCheck.ToString());
                    Console.WriteLine("Application is started : ");
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mAppTime, 300);
                    if (aeWindow == null)
                    {
                        sErrorMessage = "Explorer window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Make sure our window is usable.
                        // WaitForInputIdle will return before the specified time 
                        // if the window is ready.
                        //WindowPattern windowPattern = aeWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                        //if (false == windowPattern.WaitForInputIdle(120000))
                        //{
                       //     System.Windows.Forms.MessageBox.Show("Object not responding in a timely manner, click OK continue", "CreateDatabase");
                        //}

                        Thread.Sleep(5000);
                        try
                        {
                            Thread.Sleep(5000);
                            mAppTime = DateTime.Now;
                            AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mAppTime, 300);

                            WindowPattern windowPattern = aeWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                           // bool inputIdel = false;
                            while (windowPattern.WaitForInputIdle(100) == false)
                            {
                                Console.WriteLine("HasKeyboardFocus false");
                                //Console.WriteLine("aeWindow.Current.IsEnabled:" + aeWindow.Current.IsEnabled);
                                //Console.WriteLine("aeWindow.Current.IsKeyboardFocusable:" + aeWindow.Current.IsKeyboardFocusable);
                                Console.WriteLine("windowPattern.WaitForInputIdle(100):" + windowPattern.WaitForInputIdle(100));
                                Thread.Sleep(1000);
                            }

                            Thread.Sleep(20000);
                            aeWindow.SetFocus();
                            if (aeWindow.Current.Name.IndexOf("Demo") < 0)
                            {
                                sErrorMessage = "Explorer window demo file not loaded";
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                        catch (Exception ex)
                        {
                            sErrorMessage = "load project failed: "+ex.Message+"---"+ex.StackTrace;
                            TestCheck = ConstCommon.TEST_FAIL;
                            //System.Windows.Forms.MessageBox.Show(ex.Message+"---"+ex.StackTrace, "aeWindow.Current.Name.IndexOf(Demo) < 0");
                            // maybe close excption form
                            Console.WriteLine("maybe close excption form..." + sErrorMessage);
                            Thread.Sleep(5000);
                            mAppTime = DateTime.Now;
                            AutomationElement aeExceptionForm = null;
                            AUIUtilities.WaitUntilElementByIDFound(aeWindow, ref aeExceptionForm, "ExceptionBoxForm", mAppTime, 120);
                            if (aeExceptionForm == null)
                            {
                                sErrorMessage = "Explorer aeExceptionForm not found";
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                //TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                //aeExceptionForm.SetFocus();
                                Console.WriteLine("Explorer aeExceptionForm name : " + aeExceptionForm.Current.Name);
                                AutomationElement aeBtn = AUIUtilities.FindElementByName("OK", aeExceptionForm);
                                // OK button
                                if (aeBtn != null)
                                {
                                    Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeBtn));
                                    Thread.Sleep(3000);
                                }
                            }
                        }
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    sExploreProjectLoadOK = false;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine(testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "  ---  " + ex.StackTrace;
                Console.WriteLine(testname + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
                return;
            }
            finally
            {
                Thread.Sleep(3000);
                //Input.SendKeyboardInput(System.Windows.Input.Key.Space, true);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region EtriccExplorerEditSaveProject
        public static void EtriccExplorerEditSaveProject(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            AutomationElement aeWindow = null;
            string WindowID = "frmMain";
            string EtriccServerPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server";

            if (sExploreProjectLoadOK == false)
            {
                sErrorMessage = "EtriccExplorerLoadProject test failed, EtriccExplorerEditSaveProject cannot be tested";
                return;
            }

            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 6, 19), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                sStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - sStartTime;
                //========================   Explorer =================================================
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Find Explorer window... : ");
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mAppTime, 120);
                    if (aeWindow == null)
                    {
                        sErrorMessage = "Explorer window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        aeWindow.SetFocus();
                        Console.WriteLine("Explorer aeWindow name : " + aeWindow.Current.Name);
                    }
                }

                // focus Tree View
                //string logDirectory = Directory.GetCurrentDirectory();
                string logDirectory = @"C:\ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region
                    // Find C:... Treeitem
                    System.Windows.Automation.Condition cTree = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, "Project - Demo"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem)
                    );

                    AutomationElement aeTable = null;
                    AutomationElement aeTreeItem = aeWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cTree);
                    AutomationElement aeLogDir = null;
                    if (aeTreeItem != null)
                    {    
                        Console.WriteLine("Tree view found  .........");
                        Input.MoveToAndDoubleClick(TestTools.AUIUtilities.GetElementCenterPoint(aeTreeItem));
                        #region
                        Thread.Sleep(2000);
                        // scroll properties Window first
                        System.Windows.Automation.Condition cScrollBar = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, ""),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                        );

                        AutomationElement aeVScrollBar = aeWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cScrollBar);
                        if (aeVScrollBar == null)
                        {
                            Point pt = new Point(
                                (aeVScrollBar.Current.BoundingRectangle.Left + aeVScrollBar.Current.BoundingRectangle.Right)/2,
                                (aeVScrollBar.Current.BoundingRectangle.Top + 5)
                                );
                            int k = 0;
                            while (k++ < 20)
                                Input.MoveToAndClick(pt);

                        }

                        // Set a property condition that will be used to find the control.
                        // Find  item in Properties Window
                        System.Windows.Automation.Condition cTABLE = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Properties Window"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Table)
                        );

                        aeTable = aeWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cTABLE);
                        if (aeTable == null)
                        {
                            sErrorMessage = " aeTable not found, ";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            // Find "BaseDir" 
                            System.Windows.Automation.Condition cCus = new AndCondition(
                                new PropertyCondition(AutomationElement.NameProperty, "BaseDir"),
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom)
                            );

                            aeLogDir = aeTable.FindFirst(TreeScope.Element | TreeScope.Descendants, cCus);
                            if (aeLogDir == null)
                            {
                                sErrorMessage = " aeBaseDir not found, ";
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                if (aeLogDir.Current.IsOffscreen)
                                    Console.WriteLine("aeLogDir.Current.IsOffscreen");
                                else
                                    Console.WriteLine("aeBaseDir.Current.IsOnscreen");

                                //Thread.Sleep(1000);
                                //ValuePattern vp = aeLogDir.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                //vp.SetValue(logDirectory);

                                //aeLogDir.SetFocus();
                                Input.MoveToAndClick(aeLogDir);
                                Thread.Sleep(2000);
                                System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                                Thread.Sleep(1000);
                                System.Windows.Forms.SendKeys.SendWait(logDirectory);
                                //System.Windows.Forms.SendKeys.SendWait(logDirectory);
                                Thread.Sleep(1000);


                                Thread.Sleep(1000);
                                //Input.MoveToAndClick(aeLogDir);
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Tree view not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    #endregion
                }

                // click Save XML As...menu item             
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region
                    string menubarName = "Application";
                    Console.WriteLine("find Explorer Menubar ------------:" + menubarName);

                    System.Windows.Automation.Condition cMenubar = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, menubarName),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuBar)
                    );

                    // Find the manebar.
                    AutomationElement aeMenubar = aeWindow.FindFirst(TreeScope.Descendants, cMenubar);
                    if (aeMenubar != null)
                    {
                        // Find Root Menuitem
                        System.Windows.Automation.Condition cRoot = new AndCondition(
                           new PropertyCondition(AutomationElement.NameProperty, "Root"),
                           new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                        );

                        // Find the MenuItem Root
                        AutomationElement aeMenuItemRoot = aeMenubar.FindFirst(TreeScope.Element | TreeScope.Descendants, cRoot);
                        if (aeMenuItemRoot != null)
                        {
                            Input.MoveToAndClick(aeMenuItemRoot);
                            Thread.Sleep(3000);
                            // Find Load XML... Menuitem
                            System.Windows.Automation.Condition cLoadXML = new AndCondition(
                               new PropertyCondition(AutomationElement.NameProperty, "Save XML As..."),
                               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                            );

                            AutomationElement aeMenuItemLoadXML = aeMenuItemRoot.FindFirst(TreeScope.Element | TreeScope.Descendants, cLoadXML);
                            if (aeMenuItemLoadXML != null)
                            {
                                Input.MoveToAndClick(aeMenuItemLoadXML);
                            }
                            else
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = "Save XML... menu item not found ------------:";
                                Console.WriteLine(sErrorMessage);
                            }
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "Root menu item not found ------------:";
                            Console.WriteLine(sErrorMessage);
                        }
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Explorer menubar not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    #endregion
                }

                // delte demofile if necessary
                string demofile = OSVersionInfoClass.ProgramFilesx86() + @"\Dematic\Etricc Server\Data\Etricc\EditSaveTest.xml";
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if ( System.IO.File.Exists(demofile))
                        File.Delete(demofile);
                }

                // save new xml
                AutomationElement aeSaveAsWindow = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeWindow = null;
                    Thread.Sleep(1000);
                    DateTime mTimex = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mTimex, 300);
                    string saveAsName = "Save As";
                    aeSaveAsWindow = AUIUtilities.FindElementByName(saveAsName, aeWindow);
                    if (aeSaveAsWindow == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "aeSaveAsWindow NOT found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    else
                    {
                        DateTime sTime = DateTime.Now;
                        Console.WriteLine("Find aeFilenameEdit : " + System.DateTime.Now);
                        AutomationElement aeFilenameEdit = AUIUtilities.FindElementByID("FileNameControlHost", aeSaveAsWindow);
                        if (aeFilenameEdit == null)
                        {
                            sErrorMessage = "aeFilenameEdit not found";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Thread.Sleep(500);
                            Point pnt = TestTools.AUIUtilities.GetElementCenterPoint(aeFilenameEdit);
                            Input.MoveTo(pnt);
                            Thread.Sleep(1000);

                            aeFilenameEdit.SetFocus();
                            Thread.Sleep(2000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("EditSaveTest");
                            Thread.Sleep(1000);
                            //vp.SetValue(vp.Current.Value + "Test");

                            AutomationElement aeSaveBtn = AUIUtilities.FindElementByName("Save", aeSaveAsWindow);
                            if (aeSaveBtn == null)
                            {
                                sErrorMessage = "aeSaveBtn not found";
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                Input.MoveToAndClick(aeSaveBtn);
                            }
                        }
                    }
                }
                
                // valivation saved xml
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(20000);
                    Console.WriteLine("demofile: " + demofile);
                    if (FileManipulation.CheckSearchTextExistInFile(demofile, logDirectory, ref sErrorMessage))
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        sErrorMessage = "CheckSearchTextExistInFile failed";
                        TestCheck = ConstCommon.TEST_FAIL;
                        Console.WriteLine(sErrorMessage);
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine(testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "  ---  " + ex.StackTrace;
                Console.WriteLine(testname + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
                return;
            }
            finally
            {
                Thread.Sleep(3000);
                //Input.SendKeyboardInput(System.Windows.Input.Key.Space, true);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        static public void DeleteRecursiveFolder(DirectoryInfo dirInfo)
        {
            foreach (var subDir in dirInfo.GetDirectories())
            {
                DeleteRecursiveFolder(subDir);
            }

            foreach (var file in dirInfo.GetFiles())
            {
                file.Attributes = FileAttributes.Normal;
                file.Delete();
            }

            dirInfo.Delete();
        }
        // ————————————————————————————————————————————————————————————————————————————————————————————————————————————
        #region EtriccExplorerBuildProject
        public static void EtriccExplorerBuildProject(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sErrorMessage = string.Empty;

            AutomationElement aeWindow = null;
            string WindowID = "frmMain";
            AutomationElement aeScriptWindow = null;
            string scriptWindowId = "frmScript";

            AutomationEventHandler UICloseExplorerEventHandler = new AutomationEventHandler(OnCloseExplorerEvent);
            string EtriccServerPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server";

            if (sEtriccExplorerStartupOK == false)
            {
                sErrorMessage = "Etricc Explorer startup failed, this testcase cannot be tested";
                Console.WriteLine(sErrorMessage);
                return;
            }

            if (sExploreProjectLoadOK == false)
            {
                sErrorMessage = "EtriccExplorerLoadProject test failed, EtriccExplorerBuildProject cannot be tested";
                return;
            }

            /*if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc.Production.Release", new DateTime(2012, 6, 25), ref sErrorMessage) == true)
            {
                Console.WriteLine(sErrorMessage);
                result = ConstCommon.TEST_UNDEFINED;
                return;
            }*/

            try
            {
                string sampleProjectPath = @"C:\Etricc 5.0.0\Sample";
                DirectoryInfo dirInfo = new DirectoryInfo(sampleProjectPath);
                // WORKING Progress
                // copy sample project from dropfolders to C:\Etricc 5.0.0\Sample
                /*while (System.IO.Directory.Exists(sampleProjectPath))
                {
                    DeleteRecursiveFolder(dirInfo);
                    Thread.Sleep(2000);
                }

                if (System.IO.Directory.Exists(sampleProjectPath))
                {
                    sErrorMessage = "SampleProject not deleted ";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;

                }             

                string msgX = "Get Sample project";
                if (TestCheck == ConstCommon.TEST_PASS)
                {

                    TFSConnected = TfsUtilities.CheckTFSConnection( ref msgX, ref tfsProjectCollection );
                    while (TFSConnected == false)
                    {
                        TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                "Get Sample project", (uint)Tfs.ReconnectDelay );
                        System.Threading.Thread.Sleep( Tfs.ReconnectDelay );
                        TFSConnected = TfsUtilities.CheckTFSConnection( ref msgX, ref tfsProjectCollection );
                    }

                    if (EtriccUtilities.GetSampleProject(tfsProjectCollection, ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                */
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.UpdateCreateScriptNoRunApplication(ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }

                }

                sStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - sStartTime;
                // open build script  frmScript again
                AutomationElement aeRunScriptBtn = null;
                AutomationElement aeWindowSave = null;
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                  AutomationElement.RootElement, TreeScope.Descendants, UICloseExplorerEventHandler);

                //========================   Explorer =================================================
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Find Explorer window... : ");
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mAppTime, 120);
                    if (aeWindow == null)
                    {
                        sErrorMessage = "Explorer window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        aeWindow.SetFocus();
                        Console.WriteLine("Explorer aeWindow name : " + aeWindow.Current.Name);
                    }
                }

                // try to open new script window             
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region
                    string menubarName = "Application";
                    Console.WriteLine("find Explorer Menubar ------------:" + menubarName);

                    System.Windows.Automation.Condition cMenubar = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, menubarName),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuBar)
                    );

                    // Find the manebar.
                    AutomationElement aeMenubar = aeWindow.FindFirst(TreeScope.Descendants, cMenubar);
                    if (aeMenubar != null)
                    {
                        // Find Tools Menuitem
                        System.Windows.Automation.Condition cRoot = new AndCondition(
                           new PropertyCondition(AutomationElement.NameProperty, "Tools"),
                           new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                        );

                        // Find the MenuItem Tools
                        AutomationElement aeMenuItemTools = aeMenubar.FindFirst(TreeScope.Element | TreeScope.Descendants, cRoot);
                        if (aeMenuItemTools != null)
                        {
                            Input.MoveToAndClick(aeMenuItemTools);
                            Thread.Sleep(3000);
                            // Find Load XML... Menuitem
                            System.Windows.Automation.Condition cMenuItemScript = new AndCondition(
                               new PropertyCondition(AutomationElement.NameProperty, "Script..."),
                               new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                            );

                            AutomationElement aeMenuItemScript = aeMenuItemTools.FindFirst(TreeScope.Element | TreeScope.Descendants, cMenuItemScript);
                            if (aeMenuItemScript != null)
                            {
                                Input.MoveToAndClick(aeMenuItemScript);
                            }
                            else
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = "MenuItem Script... menu item not found ------------:";
                                Console.WriteLine(sErrorMessage);
                            }
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "Tools menu item not found ------------:";
                            Console.WriteLine(sErrorMessage);
                        }
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Explorer menubar not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    #endregion
                }
                // try to find opened build script window 
                
                AutomationElement aeOpenScriptBtn = null;
                AutomationElement aeScriptToolbar = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region
                    Thread.Sleep(3000);
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeScriptWindow, scriptWindowId, mAppTime, 300);
                    if (aeScriptWindow == null)
                    {
                        sErrorMessage = "Script Window  not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("Script window found ------------:");
                        Thread.Sleep(5000);
                        aeScriptToolbar = AUIUtilities.FindElementByType(ControlType.ToolBar, aeScriptWindow);
                        if (aeScriptToolbar == null)
                        {
                            sErrorMessage = "aeScriptToolbar not found";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine("aeScriptToolbar found");
                            aeOpenScriptBtn = AUIUtilities.FindElementByName("Open script", aeScriptToolbar);
                            if (aeScriptWindow != null)
                            {
                                Console.WriteLine("aeOpenScriptBtn found : ");
                                Input.MoveToAndClick(aeOpenScriptBtn);
                            }
                            else
                            {
                                sErrorMessage = "aeOpenScriptBtr not found";
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            
                        }
                    }
                    #endregion
                }

                // process open folder window   
                AutomationElement aeOpenFolderWiodow = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // find open folder window
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeScriptWindow, scriptWindowId, mAppTime, 300);
                    if (aeScriptWindow == null)
                    {
                        sErrorMessage = "Script window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("Script window found ------------:");
                        Thread.Sleep(5000);
                        // Find root type window
                        System.Windows.Automation.Condition cTypeWindow = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Open"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                        );

                        aeOpenFolderWiodow = aeScriptWindow.FindFirst(TreeScope.Children, cTypeWindow);
                        if (aeOpenFolderWiodow != null)
                        {
                            Console.WriteLine("eOpenFolderWiodow found ------------:");
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "eOpenFolderWiodow  not found ------------:";
                            Console.WriteLine(sErrorMessage);
                        }
                    }
                    #endregion
                }

                // find root disk
                //AutomationElement aeOpenFolderWiodow = null;
                AutomationElement aeRootDisk = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // processs tree item
                    DateTime mTime2 = DateTime.Now;
                    // Find root type window
                    System.Windows.Automation.Condition cTree = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Namespace Tree Control"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree)
                    );

                    AutomationElement aeTree = aeOpenFolderWiodow.FindFirst(TreeScope.Descendants, cTree);
                    if (aeTree != null)
                    {
                        Console.WriteLine(" aeTree found ------------:");
                        Thread.Sleep(1000);
                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition cTreeItem = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.TreeItem);

                        AutomationElementCollection aeAllItems = aeTree.FindAll(TreeScope.Descendants, cTreeItem);
                        Console.WriteLine("All items count ..." + aeAllItems.Count);
                        for (int i = 0; i < aeAllItems.Count; i++)
                        {
                            Console.WriteLine("aeAllLeafItems[i].Current.Name: " + aeAllItems[i].Current.Name);
                            if (aeAllItems[i].Current.Name.IndexOf("C:") >= 0)
                            {
                                aeRootDisk = aeAllItems[i];
                                break;
                            }
                        }

                        if (aeRootDisk == null)
                        {
                            sErrorMessage = "aeRootDisk not found name";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Input.MoveToAndClick(aeRootDisk);
                        }
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = " aeTree  not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    #endregion
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("frmScript", "Etricc 5.0.0", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Etricc 5.0.0");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("frmScript", "Sample", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Sample");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("frmScript", "Source", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Source");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("frmScript", "Script", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Script");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("frmScript", "Main", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Main");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("frmScript", "Full", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "Full");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // open build script  frmScript again
                //AutomationElement aeRunScriptBtn = null;
                //AutomationElement aeSaveWindow = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region
                    Thread.Sleep(3000);
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeScriptWindow, scriptWindowId, mAppTime, 300);
                    if (aeScriptWindow == null)
                    {
                        sErrorMessage = "Script Window  not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("Script window found2 ------------:");
                        //Thread.Sleep(3000);
                        aeScriptWindow.SetFocus();
                        Thread.Sleep(3000);
                        // Find aeRunScriptBtn
                        System.Windows.Automation.Condition cRunButton = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Run script [F5]"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                        );

                        // new UI version RunScriptBtn name changed to "Run Script [F5]"   -- added at 22-07-2014
                        System.Windows.Automation.Condition cRunButtonNew = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Run Script [F5]"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                        );

                        aeRunScriptBtn = aeScriptWindow.FindFirst(TreeScope.Descendants, cRunButton);
                        if (aeRunScriptBtn == null)   // 
                        {   // find in New version UI
                            Console.WriteLine("cRunButton find in New version UI.. :");
                            aeRunScriptBtn = aeScriptWindow.FindFirst(TreeScope.Descendants, cRunButtonNew);
                        }

                        if (aeRunScriptBtn != null)
                        {
                            Console.WriteLine("cRunButton found ------------:");
                            Point pt = AUIUtilities.GetElementCenterPoint(aeRunScriptBtn);
                            Thread.Sleep(3000);
                            Input.MoveToAndClick(pt);
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "cRunButton  not found ------------:";
                            Console.WriteLine(sErrorMessage);
                        }
                    }
                    #endregion
                }


                // wait until save project  "Save the project to file." window open
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sStartTime = DateTime.Now;
                    mTime = DateTime.Now - sStartTime;

                    // Find root type window
                      System.Windows.Automation.Condition cWindowSplash = new AndCondition(
                      new PropertyCondition(AutomationElement.AutomationIdProperty, "frmSplash"),
                      new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                      );


                    System.Windows.Automation.Condition cSaveWindow = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Save the project to file."),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                    );

                    while (aeWindowSave == null && mTime.TotalSeconds <= 600 )
                    {
                        Thread.Sleep(5000);
                        mTime = DateTime.Now - sStartTime;
                        Console.WriteLine(" find  aeSaveWindow ------------:"+mTime.TotalSeconds);
                        AutomationElement aeWindowSplash = AutomationElement.RootElement.FindFirst(TreeScope.Descendants, cWindowSplash);
                        if (aeWindowSplash == null)
                        {
                            Console.WriteLine("111111111111111   aeWindowSplash NOT FOUND");
                        }
                        else
                        {
                            Console.WriteLine("111111111111111x   aeWindowSplash FOUND -- continue to find  aeSaveWindow");
                            aeWindowSave = AutomationElement.RootElement.FindFirst(TreeScope.Descendants, cSaveWindow);
                            if (aeWindowSave == null)
                            {
                                Console.WriteLine("22222222222222   aeWindowSave NOT FOUND");
                            }
                            else
                            {
                                Console.WriteLine("22222222222222x   aeWindowSave FOUND");
                                Console.WriteLine(" find  aeWindowSave ------------mTime.TotalSeconds:" + mTime.TotalSeconds);
                            }
                        }

                    
                        
                        if (TestCheck == ConstCommon.TEST_PASS)
                        {
                            Console.WriteLine("   Test PPPPPPPPPPPPAAAAAASSSSSSSSS...............");
                        }
                        else
                        {
                            Console.WriteLine("----------------   Test Failed...............");
                        }
                        //----------------------------------- HHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHHH
                        #region 
                        if (mTime.TotalSeconds > 60)
                        {
                            Console.WriteLine(" check script running status ------------:");
                            DateTime StartTime2 = DateTime.Now;
                            TimeSpan mAppTime2 = DateTime.Now - StartTime2;
                            AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeScriptWindow, scriptWindowId, DateTime.Now, 300);
                            if (aeScriptWindow == null)
                            {
                                sErrorMessage = "Script Window  not found";
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                Console.WriteLine("Script window found3 ------------time1: " + mTime.TotalSeconds);
                                Thread.Sleep(3000);
                                // Find status bar running status
                                AutomationElement aeRunningStatusBar = AUIUtilities.FindElementByType(ControlType.StatusBar, aeScriptWindow);
                                if (aeRunningStatusBar != null)
                                {
                                    Console.WriteLine("caeRunningStatusBar found ------------:"+aeRunningStatusBar.Current.Name);
                                    if (aeRunningStatusBar.Current.Name.IndexOf("Script has errors") >= 0)
                                    {
                                        sErrorMessage = "running script failed,  Script has errors:";
                                        Console.WriteLine(sErrorMessage);
                                        TestTools.ProcessUtilities.CloseProcess("EPIA.Explorer");
                                        TestCheck = ConstCommon.TEST_FAIL;
                                        break;
                                    }
                                    else if (aeRunningStatusBar.Current.Name.IndexOf("Script running...") >= 0)
                                    {
                                        if (mTime.TotalSeconds > 600)
                                        {
                                            sErrorMessage = "Script is still running ......time2: " + mTime.TotalSeconds;
                                            Console.WriteLine(sErrorMessage);
                                            TestTools.ProcessUtilities.CloseProcess("EPIA.Explorer");
                                            TestCheck = ConstCommon.TEST_FAIL;
                                            break;
                                        }
                                        else
                                        {
                                            mTime = DateTime.Now - sStartTime;
                                            Console.WriteLine("mTime.TotalSeconds =  " + mTime.TotalSeconds);
                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                    }  // End while

                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        if (aeWindowSave != null)
                        {
                            // enlarge windows, root disk aleays clickable
                            System.Drawing.Rectangle rec = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
                            int Width = rec.Width;
                            int Height = rec.Height;

                            TransformPattern tranform = aeWindowSave.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                            if (tranform != null)
                                tranform.Move(0, 0);

                            Thread.Sleep(3000);
                            tranform.Resize((Width - Width * 0.3), (Height - Height * 0.2));

                            Console.WriteLine(" aeWindowSave found ------------:");
                            Thread.Sleep(1000);
                        }
                        else
                        {
                            sErrorMessage = " aeWindowSave NOT found ------------:";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Test Failed...............");
                        Epia3Common.WriteTestLogFail(slogFilePath, "TestCheck Failed, ????????? have a look!!!!! ", sOnlyUITest);
                    }
                }

                // find root disk           
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // processs tree item
                    DateTime mTime2 = DateTime.Now;
                    // Find root type window
                    System.Windows.Automation.Condition cTree = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Namespace Tree Control"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tree)
                    );

                    AutomationElement aeTree = aeWindowSave.FindFirst(TreeScope.Descendants, cTree);
                    if (aeTree != null)
                    {
                        Console.WriteLine(" aeTree found ------------:");
                        Thread.Sleep(1000);
                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition cTreeItem = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.TreeItem);

                        AutomationElementCollection aeAllItems = aeTree.FindAll(TreeScope.Descendants, cTreeItem);
                        Console.WriteLine("All items count ..." + aeAllItems.Count);
                        for (int i = 0; i < aeAllItems.Count; i++)
                        {
                            Console.WriteLine("aeAllLeafItems[i].Current.Name: " + aeAllItems[i].Current.Name);
                            if (aeAllItems[i].Current.Name.IndexOf("C:") >= 0)
                            {
                                aeRootDisk = aeAllItems[i];
                                break;
                            }
                        }

                        if (aeRootDisk == null)
                        {
                            sErrorMessage = "aeRootDisk not found name";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Input.MoveToAndClick(aeRootDisk);
                            //if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick(aeRootDisk.Current.Name, ref sErrorMessage))
                            /*if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("frmSplash", "Save the project to file.", aeRootDisk.Current.Name, ref sErrorMessage))
                            {
                                Console.WriteLine("successufully scroll to: " + "aeRootDisk.Current.Name");
                            }
                            else
                            {
                                Console.WriteLine(sErrorMessage);
                                Console.WriteLine("failed scroll to: " + "aeRootDisk.Current.Name");
                                Thread.Sleep(50000);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }*/
                        }
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = " aeTree  not found ------------:";
                        Console.WriteLine(sErrorMessage);
                    }
                    #endregion
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (EtriccUtilities.ScrollToThisFolderItemAndDoubleClick("frmSplash", "Save the project to file.","EtriccTests", ref sErrorMessage))
                    {
                        Console.WriteLine("successufully scroll to: " + "EtriccTests");
                    }
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // delte old project file if necessary
                string buildfile = @"C:\EtriccTests\NewBuildProject.xml";
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (System.IO.File.Exists(buildfile))
                        File.Delete(buildfile);
                }

                // save new xml
                AutomationElement aeSaveAsWindow = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeWindow = null;
                    Thread.Sleep(1000);
                    DateTime mTimex = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mTimex, 300);
                    string saveAsName = "Save As";
                    aeSaveAsWindow = AUIUtilities.FindElementByName(saveAsName, aeWindow);
                    //if (aeSaveAsWindow == null)
                    //{
                    //    TestCheck = ConstCommon.TEST_FAIL;
                    //    sErrorMessage = "aeSaveAsWindow NOT found ------------:";
                    //    Console.WriteLine(sErrorMessage);
                   // }
                    //else
                   // {
                        DateTime sTime = DateTime.Now;
                        Console.WriteLine("Find aeFilenameEdit : " + System.DateTime.Now);
                        AutomationElement aeFilenameEdit = AUIUtilities.FindElementByID("FileNameControlHost", aeWindowSave);
                        if (aeFilenameEdit == null)
                        {
                            sErrorMessage = "aeFilenameEdit not found";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Thread.Sleep(500);
                            Point pnt = TestTools.AUIUtilities.GetElementCenterPoint(aeFilenameEdit);
                            Input.MoveTo(pnt);
                            Thread.Sleep(1000);

                            aeFilenameEdit.SetFocus();
                            Thread.Sleep(2000);
                            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                            Thread.Sleep(1000);
                            System.Windows.Forms.SendKeys.SendWait("NewBuildProject");
                            Thread.Sleep(1000);
                            //vp.SetValue(vp.Current.Value + "Test");

                            AutomationElement aeSaveBtn = AUIUtilities.FindElementByName("Save", aeWindowSave);
                            if (aeSaveBtn == null)
                            {
                                sErrorMessage = "aeSaveBtn not found";
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                Input.MoveToAndClick(aeSaveBtn);
                            }
                        }
                    //}
                }

             

                // validate "Script ran successfully";
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region
                    Thread.Sleep(3000);
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeScriptWindow, scriptWindowId, mAppTime, 300);
                    if (aeScriptWindow == null)
                    {
                        sErrorMessage = "Script Window  not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("Script window found ------------:");
                        //Thread.Sleep(3000);
                        bool scriptWindowSetFocus = false;
                        while (scriptWindowSetFocus == false)
                        {
                            try
                            {
                                aeScriptWindow.SetFocus();
                                scriptWindowSetFocus = true;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("scriptWindowSetFocus exception ------------:"+ex.Message);
                                Thread.Sleep(5000);
                            }
                            
                        }

                        Thread.Sleep(3000);
                        // Find root type window
                        System.Windows.Automation.Condition cRunButton = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Script ran successfully"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.StatusBar)
                        );

                        aeRunScriptBtn = aeScriptWindow.FindFirst(TreeScope.Descendants, cRunButton);
                        if (aeRunScriptBtn != null)
                        {
                            Console.WriteLine("Script ran successfully found ------------:");
                            Point pt = AUIUtilities.GetElementCenterPoint(aeRunScriptBtn);
                            Thread.Sleep(3000);
                            //Input.MoveToAndClick(pt);
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "Script ran successfully  not found ------------:";
                            Console.WriteLine(sErrorMessage);
                        }
                    }
                    #endregion
                }

                // valivation saved xml
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(20000);
                    Console.WriteLine("buildfile: " + buildfile);
                    if (System.IO.File.Exists(buildfile))
                        TestCheck = ConstCommon.TEST_PASS;
                    else
                    {
                        sErrorMessage = "Build new project file failed";
                        TestCheck = ConstCommon.TEST_FAIL;
                        Console.WriteLine(sErrorMessage);
                    }
                }

                if (sErrorMessage.Length > 10)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    //TestCheck = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine(testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "  ---  " + ex.StackTrace;
                Console.WriteLine(testname + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
                return;
            }
            finally
            {
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                         AutomationElement.RootElement, UICloseExplorerEventHandler);

                Thread.Sleep(3000);
                //Input.SendKeyboardInput(System.Windows.Input.Key.Space, true);
            }
        }
        #endregion
        // ————————————————————————————————————————————————————————————————————————————————————————————————————————————
        #region EtriccExplorerClose
        public static void EtriccExplorerClose(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            AutomationElement aeWindow = null;
            string WindowID = "frmMain";

            AutomationElement aeScriptWindow = null;
            string scriptWindowId = "frmScript";

            AutomationElement aeCloseBtn = null;
            string closeBtnId = "Close";

            if (sEtriccExplorerStartupOK == false)
            {
                sErrorMessage = "Etricc Explorer startup failed, this testcase cannot be tested";
                return;
            }

            string EtriccServerPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server";
            AutomationEventHandler UICloseExplorerEventHandler = new AutomationEventHandler(OnCloseExplorerEvent);
            try
            {
                sStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - sStartTime;
                // open build script  frmScript again
              
                //========================   Close script window =================================================
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Find script window... : ");
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeScriptWindow, scriptWindowId, mAppTime, 120);
                    if (aeScriptWindow != null)
                    {
                        aeScriptWindow.SetFocus();
                        Console.WriteLine("aeScriptWindow aeWindow name : " + aeScriptWindow.Current.Name);
                        aeCloseBtn = AUIUtilities.FindElementByID(closeBtnId, aeScriptWindow);
                        if (aeCloseBtn != null)
                        {
                            Input.MoveToAndClick(aeCloseBtn);
                        }
                    }
                }

                
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                  AutomationElement.RootElement, TreeScope.Descendants, UICloseExplorerEventHandler);



                //========================   Close Explorer window =================================================
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Find script window... : ");
                    DateTime mAppTime = DateTime.Now;
                    AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mAppTime, 120);
                    if (aeWindow == null)
                    {
                        sErrorMessage = "aeWindow window not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        aeWindow.SetFocus();
                        Console.WriteLine("aeWindow aeWindow name : " + aeWindow.Current.Name);
                        aeCloseBtn = AUIUtilities.FindElementByID(closeBtnId, aeWindow);
                        if (aeCloseBtn != null)
                        {
                            Input.MoveToAndClick(aeCloseBtn);
                        }
                    }
                }
                
                // valivation saved xml
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(1000);
                    System.Diagnostics.Process proc;
                    int ipid = TestTools.ProcessUtilities.GetApplicationProcessID("EPIA.Explorer", out proc);
                    sStartTime = DateTime.Now;
                    mTime = DateTime.Now - sStartTime;
                    while (ipid > 0 && mTime.TotalSeconds < 120)
                    {
                        sErrorMessage = "EPIA.Explorer not closed, wait 2 seconds again  ...  ipid :" + ipid;
                        Thread.Sleep(2000);
                        mTime = DateTime.Now - sStartTime;
                        Console.WriteLine("wait extra time(sec) : " + mTime.TotalSeconds + "  ... ipid :" + ipid);
                        ipid = TestTools.ProcessUtilities.GetApplicationProcessID("EPIA.Explorer", out proc);
                    }

                    if (ipid > 0)  // check Etricc 5 testing is running
                    {
                        sErrorMessage = "EPIA.Explorer not closed";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine(testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "  ---  " + ex.StackTrace;
                Console.WriteLine(testname + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
                return;
            }
            finally
            {
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                          AutomationElement.RootElement, UICloseExplorerEventHandler);

                Thread.Sleep(3000);
                //Input.SendKeyboardInput(System.Windows.Input.Key.Space, true);
            }
        }
        #endregion
        // ————————————————————————————————————————————————————————————————————————————————————————————————————————————
        #region Event
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnUIAServerEvent
        public static void OnUIAServerEvent(object src, AutomationEventArgs args)
        {
            string message = string.Empty;
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
                Epia3Common.WriteTestLogFail(slogFilePath, "OnUIAServerEvent: Shell start exception: " + sErrorMessage, sOnlyUITest);
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
            else if (name.IndexOf("Etricc Server") > 10)
            {
                message = "OnUIAServerEvent: Etricc Server is starting ...; Name is ------------:" + name;
                Console.WriteLine(message);
                Epia3Common.WriteTestLogMsg(slogFilePath, message, sOnlyUITest);
                // try to close all extra windows
                Console.WriteLine("find this Name again ------------:" + name);
                System.Windows.Automation.Condition c = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, name),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );
            }
            else if (name.Equals(m_SystemDrive + "WINDOWS\\system32\\cmd.exe"))
            {
                Console.WriteLine("SERVER OOOOOOOOOOOO Name is ------------:" + name);
                Console.WriteLine("SERVER OOOOOOOOOOOO Name is ------------:" + name);
                Console.WriteLine("SERVER OOOOOOOOOOOO Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAServerEvent: open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("Etricc Service"))
            {
                message = "OnUIAServerEvent: : Etricc Service start up failed; Name is ------------:" + name;
                Console.WriteLine(message);
                Epia3Common.WriteTestLogMsg(slogFilePath, message, sOnlyUITest);
                // try to close all extra windows
                Console.WriteLine("find this Name again ------------:" + name);
                System.Windows.Automation.Condition c = new AndCondition(
                    //new PropertyCondition(AutomationElement.NameProperty, "Close program"),
                    new PropertyCondition(AutomationElement.NameProperty, "Cancel"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                );

                Thread.Sleep(9000);
                // Find the element.
                AutomationElement aeCloseBtn = element.FindFirst(TreeScope.Descendants, c);
                if (aeCloseBtn != null)
                {
                    Console.WriteLine("aeCloseBtn:" + aeCloseBtn.Current.Name);
                    //DEV 04 CI 0605.2
                    Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeCloseBtn));
                }
                else
                {
                    Console.WriteLine("find cmd window NOT found ------------:" + name);
                }
            }
            else if (name.Equals("ThemeManagerNotification"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAServerEvent: open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("Epia security"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAServerEvent:open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("Egemin e'pia User Interface Shell"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAServerEvent:open window name: " + name, sOnlyUITest);
                Thread.Sleep(3000);
            }
            else if (name.Equals("Egemin Shell"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAServerEvent:open window name: " + name, sOnlyUITest);
                Thread.Sleep(3000);
            }
            else if (name.Equals("Open File - Security Warning"))
            {
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAServerEvent:open window name: " + name, sOnlyUITest);
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
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAServerEvent: open other window name: " + name, sOnlyUITest);
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnUIAShellEvent
        public static void OnUIAShellEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnUIAShellEvent-Begin");
            AutomationElement element;
            string message = string.Empty;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            string automationId = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.Current.Name;
                automationId = element.Current.AutomationId;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnUIAShellEvent:name={0} : AutomationId={1}", name, automationId);
            Console.WriteLine(str);

            Thread.Sleep(2000);
            if (name.Equals("Error"))
            {
                TestCheck = ConstCommon.TEST_FAIL;
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, "shell start exception: " + sErrorMessage, sOnlyUITest);
                Thread.Sleep(6000);
            }
            else if (name.Equals("<NoName>"))
            {
                Console.WriteLine("Window with empty name, maybe a error window ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "Window with empty name, maybe a error window ------------:", sOnlyUITest);
                
                AutomationElement aeBtn = AUIUtilities.FindElementByName("OK", element);
                // find error text
                System.Windows.Automation.Condition c = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Text);

                AutomationElement aeErrorText = element.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                if (aeErrorText != null)
                {
                    sErrorMessage = aeErrorText.Current.Name;
                    if (sErrorMessage.StartsWith("Builder succeeded,")
                        || sErrorMessage.StartsWith("Builder")
                        || sErrorMessage.StartsWith("Carriers")
                        || sErrorMessage.StartsWith("Drawings"))
                    {
                        Console.WriteLine("xx2 BUILDER succeeded...     : " + sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "xx2 BUILDER succeeded... " + sErrorMessage, sOnlyUITest);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Console.WriteLine("xx shell start exception: " + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, "xx shell start exception: " + sErrorMessage, sOnlyUITest);
                    }
                }
                else
                {
                    Console.WriteLine("Not a Error window ------------:" + name);
                    return;
                }
                // OK button
                if (aeBtn != null)
                {
                    Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeBtn));
                    Thread.Sleep(3000);
                }
            }
            else if (name.Equals("Egemin.Epia.Shell"))  // in case of shell crash
            {
                Console.WriteLine("epia shell crash Window, click  Cancel or Close the program button------------:" + name);
                AutomationElement aeBtn = AUIUtilities.FindElementByName("Cancel", element);
                if (aeBtn != null)
                {
                    Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeBtn));
                    Thread.Sleep(3000);
                }
                else
                {
                    AutomationElement aeCloseProgram = AUIUtilities.FindElementByName("Close the program", element);
                    if (aeCloseProgram != null)
                    {
                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeCloseProgram));
                        Thread.Sleep(3000);
                    }
                }
            }
            else if (name.Equals(m_SystemDrive + "WINDOWS\\system32\\cmd.exe"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAShellEvent open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("Epia security"))
            {
                Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAShellEvent open window name: " + name, sOnlyUITest);
                return;
            }
            else if (name.Equals("Egemin e'pia User Interface Shell"))
            {
                Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAShellEvent open window name: " + name, sOnlyUITest);
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
                else if (element.Current.AutomationId.Equals("Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen"))
                {
                    AutomationElement aeBtn = AUIUtilities.FindElementByID("m_BtnRetryRegistration", element);
                    if (aeBtn != null)
                    {
                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeBtn));
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        Console.WriteLine("m_BtnRetryRegistration not found ------------:");
                        Epia3Common.WriteTestLogFail(slogFilePath, "m_BtnRetryRegistration not found: ", sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                else
                {
                    Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                    Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                    Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAShellEvent: open window name: " + name, sOnlyUITest);
                    Thread.Sleep(3000);
                }
            }
            else if (name.Equals("Open File - Security Warning"))
            {
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAShellEvent open window name: " + name, sOnlyUITest);
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
                Epia3Common.WriteTestLogMsg(slogFilePath, "OnUIAShellEvent Do ELSE OTHER is  name: " + name, sOnlyUITest);
            }
            sEventEnd = true;
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnOpenExplorerEvent
        public static void OnOpenExplorerEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnOpenExplorerEvent");
            AutomationElement element;
            string message = string.Empty;
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
            string str = string.Format("OnOpenExplorerEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(2000);
            if (name.Equals("Save"))
            {
                // Find the element.
                AutomationElement aeBtnNo = AUIUtilities.FindElementByName("No", element);
                if (aeBtnNo != null)
                {
                    Point pt = AUIUtilities.GetElementCenterPoint(aeBtnNo);
                    Input.MoveTo(pt);
                    Thread.Sleep(1000);
                    Input.ClickAtPoint(pt);
                }
            }
            else if (name.Equals("Exit"))
            {
                // Find the element.
                AutomationElement aeBtnYes = AUIUtilities.FindElementByName("Yes", element);
                if (aeBtnYes != null)
                {
                    Point pt = AUIUtilities.GetElementCenterPoint(aeBtnYes);
                    Input.MoveTo(pt);
                    Thread.Sleep(1000);
                    Input.ClickAtPoint(pt);
                }
            }
            else if (name.EndsWith("Exception"))
            {
                // Find the element.
                Thread.Sleep(3000);
                AutomationElement aeTxt = AUIUtilities.FindElementByID("txbException", element);
                if (aeTxt != null)
                {
                    TextPattern tp = (TextPattern)aeTxt.GetCurrentPattern(TextPattern.Pattern);
                    Thread.Sleep(1000);
                    sErrorMessage = tp.DocumentRange.GetText(-1).Trim().ToString(); ;
                    Console.WriteLine("Error Message Catched ------------:" + sErrorMessage);
                    //Console.WriteLine("Error Message is ------------:" + sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, "start explorer failed: " + sErrorMessage, sOnlyUITest);

                    AutomationElement aeBtnYes = AUIUtilities.FindElementByName("OK", element);
                    if (aeBtnYes != null)
                    {
                        Point pt = AUIUtilities.GetElementCenterPoint(aeBtnYes);
                        Input.MoveTo(pt);
                        Thread.Sleep(1000);
                        Console.WriteLine("Click OK ------------:");
                        Input.ClickAtPoint(pt);
                    }
                }
                else
                {
                    Console.WriteLine("Error Message not found ------------:");
                    Epia3Common.WriteTestLogFail(slogFilePath, "Error Message pane not found: ", sOnlyUITest);
                }
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
            else
            {
                Console.WriteLine("Do ELSE OTHER is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
            }
            sEventEnd = true;
        }
        #endregion OnOpenExplorerEvent

        public static string GetErrorMessage()
        {
            return sErrorMessage;
        }

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnCloseExplorerEvent
        public static void OnCloseExplorerEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnCloseExplorerEvent");
            AutomationElement element;
            string name = "";
            string message = string.Empty;
            try
            {
                element = src as AutomationElement;
                if (element == null)
                {
                    name = "null";
                    return;
                }
                else
                {
                    name = element.GetCurrentPropertyValue(
                        AutomationElement.NameProperty) as string;
                }
            }
            catch
            {
                return;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnCloseExplorerEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(2000);
            if (name.Equals("Save"))
            {
                // Find the element.
                AutomationElement aeBtnNo = AUIUtilities.FindElementByName("No", element);
                if (aeBtnNo != null)
                {
                    Point pt = AUIUtilities.GetElementCenterPoint(aeBtnNo);
                    Input.MoveTo(pt);
                    Thread.Sleep(1000);
                    Input.ClickAtPoint(pt);
                }
            }
            else if (name.Equals("Exit"))
            {
                // Find the element.
                AutomationElement aeBtnYes = AUIUtilities.FindElementByName("Yes", element);
                if (aeBtnYes != null)
                {
                    Point pt = AUIUtilities.GetElementCenterPoint(aeBtnYes);
                    Input.MoveTo(pt);
                    Thread.Sleep(1000);
                    Input.ClickAtPoint(pt);
                }
            }
            else if (name.EndsWith("Exception"))
            {
                // Find the element.
                AutomationElement aeBtnYes = AUIUtilities.FindElementByName("OK", element);
                if (aeBtnYes != null)
                {
                    Point pt = AUIUtilities.GetElementCenterPoint(aeBtnYes);
                    Input.MoveTo(pt);
                    Thread.Sleep(1000);
                    Input.ClickAtPoint(pt);
                }
            }
            else if (name.EndsWith("Builder") || name.EndsWith("Router") || name.EndsWith("Finisher"))
            {
                Epia3Common.WriteTestLogMsg(slogFilePath, " <------->   NEW Build message:"+name, sOnlyUITest);
                AutomationElement aeText = AUIUtilities.FindElementByType(ControlType.Text, element);
                if (aeText != null)
                {
                    sErrorMessage = aeText.Current.Name;
                    Epia3Common.WriteTestLogFail(slogFilePath, aeText.Current.Name, sOnlyUITest);
                }
              
                // Find the element.
                AutomationElement aeBtnYes = AUIUtilities.FindElementByName("Ignore", element);
                if (aeBtnYes != null)
                {
                    Point pt = AUIUtilities.GetElementCenterPoint(aeBtnYes);
                    Input.MoveTo(pt);
                    Thread.Sleep(1000);
                    Input.ClickAtPoint(pt);
                    Thread.Sleep(3000);
                }
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
                    if (sErrorMessage.StartsWith("Builder") 
                        || sErrorMessage.StartsWith("Agvs")
                        || sErrorMessage.StartsWith("Carriers")
                        || sErrorMessage.StartsWith("Drawings"))
                    {
                        Console.WriteLine("This is a Build window ------------:" + name);
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        Console.WriteLine("321 shell start exception: " + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, "321 shell start exception: " + sErrorMessage, sOnlyUITest);
                    }
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
            /*else if (name.Equals("Egemin e'pia User Interface Shell"))
            {
                Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
                Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
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
            }*/
            else
            {
                Console.WriteLine("Do ELSE OTHER is ------------:" + name);
                Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
            }
            sEventEnd = true;
        }
        #endregion OnCloseExplorerEvent
        #endregion End Event
        //===============================================================================================================
        public static void SendEmail(string resultFile)
        {
            string str1 = DeployUtilities.GetTestReportContentString(sTotalCounter, sTotalPassed, sTotalFailed, sTotalException, sTotalUntested,
               sCurrentPlatform, sInstallMsiDir); // AnyCPU 

            TestTools.ProcessUtilities.SendTestResultToDevelopers(resultFile, sProjectFile, sBuildDef, logger, sTotalFailed,
               sBuildNr /*used for email title*/, str1/*content*/, sSendMail);
        }
    }
}
