using System;
using System.IO;
using System.Configuration;
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

namespace EtriccGUIAutoTest
{
	#pragma warning disable 0162 // Disable warning for Unreachable Code Detected.
	class EtriccProgram
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
        static IBuildServer m_BuildSvc;
        static bool TFSConnected = true;
        //static BuildStore buildStore = null;
        static string sInstallScriptsDir = string.Empty;
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
        public static string slogFilePath = @"C:\";
        static string sErrorMessage;
        static string sExcelVisible = string.Empty;
        static string sServerRunAs = string.Empty;
        static string sOutFilename = string.Empty;
        static string sOutFilePath = string.Empty;
        static StreamWriter Writer;

        // Test Param. =========================================================
        static string sTFSServer = "http://teamApplication.teamSystems.egemin.be:8080";
        static string sProjectFile  = "demo.xml";
        static AutomationElement aeForm = null;
        static int Counter          = 0;
        static string[] sTestCaseName = new string[100];
        static DateTime sTestStartUpTime = DateTime.Now;
        static int sTotalCounter    = 0;
        static int sTotalException  = 0;
        static int sTotalFailed     = 0;
        static int sTotalPassed     = 0;
        static int sTotalUntested   = 0;
        static int TestCheck = ConstCommon.TEST_UNDEFINED;
        static public string TimeOnPC;
        static bool sEventEnd = false;
        static bool sAutoTest = true;
        static bool sFunctionalTest = true;
        static bool sDemo = false;
        static string sSendMail = "false";
        static string m_SystemDrive = string.Empty;

        static DateTime sStartTime = DateTime.Now;
        static TimeSpan sTime;

        static string sFuncTotalFailed = "0";
        private static int sNumAgvs = 2;

        static bool sOnlyUITest = false;
        static string sTestType = "all"; 
        // excel 	--------------------------------------------------------
        static Excel.Application xApp;
        static Excel.Workbook xBook;
        static Excel.Workbooks xBooks;
        static Excel.Range xRange;
        static dynamic xSheet;

        static Excel.Application xAppFunc;

        #region TestCase Name
        private const string DISPLAY_SYSTEM_OVERVIEW        = "SystemOverviewDisplay";
        private const string DISPLAY_AGV_OVERVIEW           = "AgvOverviewDisplay";
        private const string DISPLAY_LOCATION_OVERVIEW_START_NODE      = "LocationOverviewScreenSelectNodeAsStart";
        private const string DISPLAY_LOCATION_OVERVIEW_END_NODE = "LocationOverviewScreenSelectNodeAsEnd";
        private const string DISPLAY_STATION_OVERVIEW_START_NODE = "StationOverviewScreenSelectNodeAsStart";
        private const string DISPLAY_STATION_OVERVIEW_END_NODE = "StationOverviewScreenSelectNodeAsEnd";
        private const string DISPLAY_TRANSPORT_OVERVIEW     = "TransportOverviewDisplay";
        private const string MULTI_LANGUAGE_CHECK = "MultiLanguageCheck";
        private const string LOCATION_OVERVIEW_OPEN_DETAIL  = "LocationOverviewOpenDetail";
        private const string LOCATION_MODE_MANUAL           = "LocationModeManual";
        private const string AGV_OVERVIEW_OPEN_DETAIL       = "AgvOverviewOpenDetail";
        private const string AGV_JOB_OVERVIEW               = "AgvJobsOverview";
        private const string AGV_JOB_OVERVIEW_OPEN_DETAIL   = "AgvJobOverviewOpenDetail";
        private const string AGV_RESTART                    = "AgvRestart";
        private const string AGV_MODE                       = "AgvMode";
        private const string CREATE_NEW_TRANSPORT           = "TransportCreateNew";
        private const string EDIT_TRANSPORT                 = "TransportEdit";
        private const string SUSPEND_TRANSPORT              = "TransportSuspend";
        private const string RELEASE_TRANSPORT              = "TransportRelease";
        private const string CANCEL_TRANSPORT               = "TransportCancel";
        private const string TRANSPORT_OVERVIEW_OPEN_DETAIL = "TransportOverviewOpenDetail";
        private const string AGV_OVERVIEW_REMOVE_ALL        = "AgvsAllModeRemoved";
        private const string AGV_OVERVIEW_ID_SORTING        = "AgvsIdSorting";
        private const string SYSTEM_OVERVIEW_QUERY          = "SystemOverviewQuery";
        private const string EPIA4_CLOSE                    = "Epia4Close";
        private const string SCRIPT_BEFORE_AFTER_ACTIVATE   = "BeforeAfterActivateScript";
        private const string SCRIPT_BEFORE_AFTER_DEACTIVATE = "BeforeAfterDeactivateScript";
        #endregion TestCase Name
        

        private const string INFRASTRUCTURE = "E'tricc®";
        private const string SYSTEM_OVERVIEW    = "System Overview";
        private const string AGV_OVERVIEW       = "Agvs";
        private const string LOCATION_OVERVIEW  = "Locations";
        private const string STATION_OVERVIEW = "Stations";
        private const string TRANSPORT_OVERVIEW = "Transports";
        private const string NEW_TRANSPORT      = "New Transport";

        private const string SYSTEM_OVERVIEW_TITLE      = "System overview";
        private const string AGV_OVERVIEW_TITLE         = "Agvs";
        private const string LOCATION_OVERVIEW_TITLE = "Locations";
        private const string STATION_OVERVIEW_TITLE = "Stations";
        private const string TRANSPORT_OVERVIEW_TITLE   = "Transports";

        private const string DATAGRIDVIEW_ID = "m_GridData";
        private const string AGV_GRIDDATA_ID = "m_GridData";
        private const string MESSAGESCREEN_ID = "MessageScreen";
        private static string sScreenResolution = string.Empty;

        // Test Case Status. =========================================================
        private static bool sTransportSuspendOK = false;
        private static bool sBeforeAfterActivateScriptOK = true;

        [STAThread]
        static void Main(string[] args)
        {
            try  // Get test PC info======================================
            {
                m_CurrentDrive = Path.GetPathRoot(Directory.GetCurrentDirectory());
                HelpUtilities.SavePCInfo("y");
                HelpUtilities.GetPCInfo(out PCName, out OSName, out OSVersion, out UICulture, out TimeOnPC);
                Console.WriteLine("PCName : " + PCName);
                Console.WriteLine("OSName : " + OSName);
                Console.WriteLine("OSVersion : " + OSVersion);
                Console.WriteLine("UICulture : " + UICulture);
                Console.WriteLine("TimeOnPC : " + TimeOnPC);

                int w = System.Windows.Forms.SystemInformation.VirtualScreen.Width;
                int h = System.Windows.Forms.SystemInformation.VirtualScreen.Height;
                sScreenResolution = w + "x" + h;
                //System.Windows.MessageBox.Show("screen resolution: " + w + "x" + h);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            sOnlyUITest = false;
            string x = System.Configuration.ConfigurationManager.AppSettings.Get("OnlyUITest");
            if (x.ToLower().StartsWith("true"))
                sOnlyUITest = true;

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

                        sInstallScriptsDir = args[0];
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

                        sTestResultFolder = sBuildDropFolder + "\\TestResults";
                        if (!System.IO.Directory.Exists(sTestResultFolder))
                            System.IO.Directory.CreateDirectory(sTestResultFolder);

                        //Epia3Common.CreateOutputFileInfo(args, PCName, ref sOutFilePath, ref sOutFilename);
                        CreateOutputFileInfo(args, sCurrentPlatform, PCName, ref sOutFilePath, ref sOutFilename);

                      

                        sOutFilePath = Path.Combine(sBuildDropFolder, "TestResults");
                        Console.WriteLine("sOutFilePath : " + sOutFilePath);

                        Epia3Common.CreateTestLog(ref slogFilePath, sOutFilePath, sOutFilename, ref Writer);

                        Epia3Common.WriteTestLogMsg(slogFilePath, "sReportDirectory : " + sTestResultFolder, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "sOutFilePath : " + sOutFilePath, sOnlyUITest);

                        Epia3Common.WriteTestLogMsg(slogFilePath, "0) Install msi file path: " + sInstallScriptsDir, sOnlyUITest);
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

                        Console.WriteLine("slogFilePath : " + slogFilePath);
                        Console.WriteLine("sOutFilePath : " + sOutFilePath);
                        Console.WriteLine("sOutFilename : " + sOutFilename);
                        logger = new Logger(slogFilePath);

                        string windir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
                        m_SystemDrive = Path.GetPathRoot(windir);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "m_SystemDrive: " + m_SystemDrive, sOnlyUITest);

                        Console.WriteLine("0) sInstall msi files Dir : " + sInstallScriptsDir);
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
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            if (sAutoTest)
            {
                if (!sOnlyUITest)
                {
                    try
                    {
                        // Get TFS Server
                        string serverUrl = "http://team2010app.teamsystems.egemin.be:8080/tfs/Development";
                        Uri serverUri = new Uri(serverUrl);
                        System.Net.ICredentials tfsCredentials
                            = new System.Net.NetworkCredential("TfsBuild", "Egemin01", "TeamSystems.Egemin.Be");

                        TfsTeamProjectCollection tfsProjectCollection
                            = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                        tfsProjectCollection.Connect(Microsoft.TeamFoundation.Framework.Common.ConnectOptions.IncludeServices);
                        int kTime = 0;
                        bool conn = false;
                        while (conn == false)
                        {
                            try
                            {
                                m_BuildSvc = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));
                                conn = true;
                            }
                            catch (Microsoft.TeamFoundation.TeamFoundationServiceUnavailableException ex)
                            {
                                TestTools.MessageBoxEx.Show("Team Foundation services are not available from server\nWill try to reconnect the Server after 10 minutes",
                                kTime++ + " During E'tricc UI Testing, please not touch the screen, time: " + DateTime.Now.ToLongTimeString(), 10 * 60000);
                                System.Threading.Thread.Sleep(10 * 60000);
                                conn = false;
                            }
                            catch (Exception ex)
                            {
                                TestTools.MessageBoxEx.Show("TeamFoundation getService Exception:" + ex.Message + " ----- " + ex.StackTrace,
                                     kTime++ + " This is automatic testing, please not touch the screen: exception time:" + DateTime.Now.ToLongTimeString(), 10 * 60000);
                                System.Threading.Thread.Sleep(10 * 60000);
                                conn = false;
                            }
                            
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + "--" + ex.StackTrace);
                        TFSConnected = false;
                    }
                }
            }
            else
                TFSConnected = false;

            Console.WriteLine("Test started:");
            sTestCaseName[0] = DISPLAY_SYSTEM_OVERVIEW;
            sTestCaseName[1] = SYSTEM_OVERVIEW_QUERY;
            sTestCaseName[2] = DISPLAY_AGV_OVERVIEW;
            sTestCaseName[3] = DISPLAY_LOCATION_OVERVIEW_START_NODE;
            sTestCaseName[4] = DISPLAY_LOCATION_OVERVIEW_END_NODE;
            sTestCaseName[5] = DISPLAY_STATION_OVERVIEW_START_NODE;
            sTestCaseName[6] = DISPLAY_STATION_OVERVIEW_END_NODE;
            sTestCaseName[7] = DISPLAY_TRANSPORT_OVERVIEW;
            sTestCaseName[8] = LOCATION_OVERVIEW_OPEN_DETAIL;
            sTestCaseName[9] = LOCATION_MODE_MANUAL;
            sTestCaseName[10] = AGV_OVERVIEW_OPEN_DETAIL;
            sTestCaseName[11] = AGV_RESTART;
            sTestCaseName[12] = AGV_MODE;
            sTestCaseName[13] = AGV_JOB_OVERVIEW;
            sTestCaseName[14] = AGV_JOB_OVERVIEW_OPEN_DETAIL;
            sTestCaseName[15] = CREATE_NEW_TRANSPORT;
            sTestCaseName[16] = EDIT_TRANSPORT;
            sTestCaseName[17] = SUSPEND_TRANSPORT;
            sTestCaseName[18] = RELEASE_TRANSPORT;
            sTestCaseName[19] = CANCEL_TRANSPORT;
            sTestCaseName[20] = TRANSPORT_OVERVIEW_OPEN_DETAIL;
            sTestCaseName[21] = AGV_OVERVIEW_REMOVE_ALL;
            sTestCaseName[22] = AGV_OVERVIEW_ID_SORTING;
            sTestCaseName[23] = MULTI_LANGUAGE_CHECK;
            sTestCaseName[24] = SCRIPT_BEFORE_AFTER_ACTIVATE;
            sTestCaseName[25] = SCRIPT_BEFORE_AFTER_DEACTIVATE;
            sTestCaseName[26] = EPIA4_CLOSE;
           
            try
            {
                if (!sOnlyUITest)
                {
                    Utilities.CloseProcess("EXCEL");
                    Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                    Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    Thread.Sleep(1000);

                    //========================   SERVER =================================================
                    #region SERVER
                    if ( sServerRunAs.ToLower().IndexOf("service") >= 0)
                    {
                        // uninstall Egemin.Epia.server Service
                        Console.WriteLine("UNINSTALL EPIA SERVER Service : ");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "UNINSTALL EPIA SERVER Service : "
                            +m_SystemDrive+ ConstCommon.EPIA_SERVER_ROOT, sOnlyUITest);
                        TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT, 
                            ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /u");
                        Thread.Sleep(2000);

                        // uninstall Egemin.Etricc.server Service
                        Console.WriteLine("UNINSTALL ETRICC SERVER Service : ");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "UNINSTALL ETRICC SERVER Service : "
                            + sEtriccServerRoot, sOnlyUITest);
                        TestTools.Utilities.StartProcessWaitForExit(sEtriccServerRoot, 
                            ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /u");
                        Thread.Sleep(2000);

                        Console.WriteLine("INSTALL EPIA SERVER Service : ");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "INSTALL EPIA SERVER Service : "
                           + m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT, sOnlyUITest);
                        TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT, 
                            ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /i");
                        Thread.Sleep(2000);

                        Console.WriteLine("INSTALL ETRICC SERVER Service : ");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "INSTALL ETRICC SERVER Service : "
                           + sEtriccServerRoot, sOnlyUITest);
                        TestTools.Utilities.StartProcessWaitForExit(sEtriccServerRoot, 
                            ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /i");
                        Thread.Sleep(2000);

                        Console.WriteLine("Start EPIA SERVER as Service : ");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Start EPIA SERVER as Service : "
                           + m_SystemDrive+ConstCommon.EPIA_SERVER_ROOT + " - " + ConstCommon.EGEMIN_EPIA_SERVER_EXE, 
                           sOnlyUITest);
                        TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT, 
                            ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /start");
                        Thread.Sleep(2000);

                        Console.WriteLine("Start ETRICC SERVER as Service : ");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Start ETRICC SERVER as Service : "
                           + sEtriccServerRoot + " - " + ConstCommon.EGEMIN_ETRICC_SERVER_EXE, sOnlyUITest);
                        TestTools.Utilities.StartProcessWaitForExit(sEtriccServerRoot, 
                            ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /start");
                        Thread.Sleep(2000);

                        ServiceController svcEpia = new ServiceController("Egemin Epia Server");
                        Console.WriteLine(svcEpia.ServiceName + " has status " + svcEpia.Status.ToString());
                        Thread.Sleep(2000);

                        //svcEpia.WaitForStatus(ServiceControllerStatus.Running);
                        // wait until epia service status is Running
                        sStartTime = DateTime.Now;
                        sTime = DateTime.Now - sStartTime;
                        int wat = 0;
                        string epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                        Console.WriteLine(" time is :" + sTime.TotalMilliseconds);
                        while (!epiaServiceStatus.StartsWith("running") && wat < 60)
                        {
                            Thread.Sleep(2000);
                            wat = wat + 2;
                            epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                            Console.WriteLine("wait epia service status running:  time is (sec) : " + wat + "  and status is:" + epiaServiceStatus);
                        }

                        if (svcEpia.Status != ServiceControllerStatus.Running)
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, "Epia Service start up failed: " + epiaServiceStatus, sOnlyUITest);
                            throw new Exception("Epia service start up failed:"); //   get message from log file sErrorMessage//
                        }
                        /*while (svcEpia.Status != ServiceControllerStatus.Running)
                        {
                            Console.WriteLine(svcEpia.ServiceName + " ==has status == " + svcEpia.Status.ToString());
                            Thread.Sleep(2000);
                        }
                        */

                        ServiceController svcEtricc = new ServiceController("Egemin Etricc Server");
                        Console.WriteLine(svcEtricc.ServiceName + " has status " + svcEtricc.Status.ToString());
                        Thread.Sleep(2000);
                        //svcEtricc.WaitForStatus(ServiceControllerStatus.Running);
                        // wait until etricc service status is Running
                        sStartTime = DateTime.Now;
                        sTime = DateTime.Now - sStartTime;
                        int wait = 0;
                        string etriccServiceStatus = svcEtricc.Status.ToString().ToLower();
                        Console.WriteLine(" time is :" + sTime.TotalMilliseconds);
                        while (!etriccServiceStatus.StartsWith("running") && wait < 150)
                        {
                            Thread.Sleep(2000);
                            wait = wait + 2;
                            etriccServiceStatus = svcEtricc.Status.ToString().ToLower();
                            Console.WriteLine("wait etricc service status running:  time is (sec) : " + wait + "  and status is:" + etriccServiceStatus);
                        }

                        if (svcEtricc.Status != ServiceControllerStatus.Running)
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, "Etricc Service start up failed: " + etriccServiceStatus, sOnlyUITest);
                            throw new Exception("Etricc service start up failed:"); //   get message from log file sErrorMessage//
                        }

                   
                        /*while (svcEtricc.Status != ServiceControllerStatus.Running)
                        {
                            Console.WriteLine(svcEtricc.ServiceName + " ==has status == " + svcEtricc.Status.ToString());
                            Thread.Sleep(2000);
                        }
                        */
                        Console.WriteLine("EPIA and ETRICC SERVER Service Started : ");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "EPIA and ETRICC SERVER Service Started:", sOnlyUITest);
                        Thread.Sleep(2000);
                    }

                    if (sServerRunAs.ToLower().IndexOf("console") >= 0)
                    {
                        Console.WriteLine("Start EPIA and ETRICC Server as console applications : ");
                        // Start Epia SERVER as Console
                        TestTools.Utilities.StartProcessNoWait(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT, 
                            ConstCommon.EGEMIN_EPIA_SERVER_EXE, string.Empty);

                        // Start Etricc SERVER as Console
                        TestTools.Utilities.StartProcessNoWait(sEtriccServerRoot, 
                            ConstCommon.EGEMIN_ETRICC_SERVER_EXE, string.Empty);
                        Thread.Sleep(90000);
                    }
                    #endregion
                    Thread.Sleep(5000);

                    sStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - sStartTime;

                    //========================   SHELL =================================================
                    #region  Shell
                    AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
                    // Add Open window Event Handler
                    Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                        AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);
                    sEventEnd = false;
                    TestCheck = ConstCommon.TEST_PASS;

                    Thread.Sleep(45000);

                    // Start Shell
                    //TestTools.Utilities.StartProcessNoWait(sInstallScriptsDir, Constants.SHELL_BAT, string.Empty);
                    TestTools.Utilities.StartProcessNoWait(
                        m_SystemDrive+ConstCommon.EPIA_CLIENT_ROOT, ConstCommon.EGEMIN_EPIA_SHELL_EXE, string.Empty);
                   
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
                    
                    #endregion

                    Console.WriteLine("Shell started after (sec) : " + mTime.Seconds);

                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                           AutomationElement.RootElement,
                          UIAShellEventHandler);

                    if (TestCheck == ConstCommon.TEST_FAIL)
                    {
                        throw new Exception("shell start up failed:" + sErrorMessage);
                    }
                    Thread.Sleep(8000);
                    aeForm = null;
                    DateTime mAppTime = DateTime.Now;
                    mTime = DateTime.Now - mAppTime;
                    while (aeForm == null && mTime.Minutes < 5)
                    {
                        Console.WriteLine("Find Application aeForm : " + System.DateTime.Now);
                        aeForm = AUIUtilities.FindElementByID("MainForm", AutomationElement.RootElement);
                        Console.WriteLine("Application aeForm name : " + System.DateTime.Now);
                        mTime = DateTime.Now - mAppTime;
                        Console.WriteLine(" find time is :" + mTime.TotalMilliseconds/1000);
                    }
                    // if after 5 minutes still no mainform,throw exception 
                    if (aeForm == null)
                    {
                        AutomationElement aeError = AUIUtilities.FindElementByID("ErrorScreen", AutomationElement.RootElement);
                        if (aeError != null)
                            AUICommon.ErrorWindowHandling(aeError, ref sErrorMessage);
                        else
                            sErrorMessage = "Application Startup failed,see logging";

                        throw new Exception(sErrorMessage);
                    }
                    else
                        Console.WriteLine("Application maeForm name : " + aeForm.Current.Name);
                }

                // Excel file not for EpiaTestPC3
                if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
                        || PCName.ToUpper().Equals("EPIATESTSRV3V1")  )
                    Epia3Common.WriteTestLogMsg(slogFilePath, "No Excel due to: " + PCName, sOnlyUITest);
                else
                {
                    xApp = new Excel.Application();
                    xBooks = xApp.Workbooks;
                    xBook = xBooks.Add(Type.Missing);
                    //xSheet = (Excel.Worksheet)xBook.Worksheets[1];
                    xSheet = xApp.ActiveSheet;
                    if (sExcelVisible == string.Empty)
                        xApp.Visible = Constants.VISIBLE;
                    else
                    {
                        if (sExcelVisible.StartsWith("Visible"))
                            xApp.Visible = true;
                        else
                            xApp.Visible = false;
                    }

                    xApp.Interactive = true;
                    string today = System.DateTime.Now.ToString("MMMM-dd");
                    xSheet.Cells[1, 1] = today;
                    xSheet.Cells[1, 2] = "Etricc UI Test Scenarios";

                    xSheet.Cells[2, 1] = "Test Machine:";
                    xSheet.Cells[2, 2] = PCName;
                    xSheet.Cells[3, 1] = "Tester:";
                    xSheet.Cells[3, 2] = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                    xSheet.Cells[4, 1] = "OSName:";
                    xSheet.Cells[4, 2] = OSName;
                    xSheet.Cells[5, 1] = "OS Version:";
                    xSheet.Cells[5, 2] = OSVersion;
                    xSheet.Cells[6, 1] = "UI Culture";
                    xSheet.Cells[6, 2] = UICulture;
                    xSheet.Cells[7, 1] = "Time On PC";
                    xSheet.Cells[7, 2] = "local time:" + TimeOnPC;
                    xSheet.Cells[8, 1] = "Test Tool Version:";
                    xSheet.Cells[8, 2] = sTestToolsVersion;
                    if (sInstallScriptsDir != null)
                    {
                        xSheet.Cells[9, 1] = "Build Location:";
                        xSheet.Cells[9, 2] = sInstallScriptsDir;
                    }
                    xSheet.Cells[10, 1] = "Screen Resolution:";
                    xSheet.Cells[10, 2] = sScreenResolution;
                }
                // check Shell proc. if exist , get Proc id
                System.Diagnostics.Process ShellProcess = null;
                int pID = Utilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out ShellProcess);
                Console.WriteLine("Proc ID:" + pID);
                aeForm = AutomationElement.FromHandle(ShellProcess.MainWindowHandle);

                if (aeForm == null)
                {
                    Console.WriteLine("aeForm  not found : ");
                    return;
                }
                else
                    Console.WriteLine("aeForm found name : " + aeForm.Current.Name);

                // start test----------------------------------------------------------
                int sResult = ConstCommon.TEST_UNDEFINED;
                int aantal = 27;
               
                if ( sDemo )
                    aantal = 2;

                if (sOnlyUITest)
                {
                    aantal = 27;
                    sTestType = System.Configuration.ConfigurationManager.AppSettings.Get("TestType");
                    if (sTestType.ToLower().StartsWith("all"))
                    {
                        aantal = 27;
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
                    }
                }
                else
                {
                    if (TFSConnected)
                    {
                        Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                            TestTools.TfsUtilities.GetProjectName("Etricc UI"), sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                        if (quality.Equals("GUI Tests Failed"))
                        {
                            string msg = sBuildNr + " has build quality: " + quality + " , no update needed";
                            Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has build quality: " + quality + " , no update needed", sOnlyUITest);
                        }
                        else
                        {
                            TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName("Etricc UI"),
                                "GUI Tests Started", m_BuildSvc, sDemo ? "true" : "false");
                        }

                        if (sAutoTest)
                        {
                            if (sParentProgram.StartsWith("TFS"))
                            {
                                FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", ConstCommon.ETRICCUI + "+" + sCurrentPlatform + "Normal");
                                Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.ETRICCUI, sOnlyUITest);
                            }
                            else
                            {
                                FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", ConstCommon.ETRICC_UI +"+" +sCurrentPlatform + "Normal");
                                Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.ETRICC_UI, sOnlyUITest);
                            }
                        }
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
                    sResult = ConstCommon.TEST_UNDEFINED;
                    switch (sTestCaseName[Counter])
                    {
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
                            break;
                        case CREATE_NEW_TRANSPORT:
                            CreateNewTransport(CREATE_NEW_TRANSPORT, aeForm, out sResult);
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
                            break;
                        case SYSTEM_OVERVIEW_QUERY:
                            SystemOverviewQuery(SYSTEM_OVERVIEW_QUERY, aeForm, out sResult);
                            break;
                        default:
                            break;
                    }

                    if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3") 
                        || PCName.ToUpper().Equals("EPIATESTSRV3V1")  )
                        Console.WriteLine("No Excel due to: " + PCName);
                    else
                        WriteResult(sResult, Counter, sTestCaseName[Counter], xSheet, sErrorMessage);

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

                if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3") 
                    || PCName.ToUpper().Equals("EPIATESTSRV3V1")  )
                    Epia3Common.WriteTestLogMsg(slogFilePath, "No Excel due to: " + PCName, sOnlyUITest);
                else
                {
                    xSheet.Cells[Counter + 2 + 9, 1] = "Total tests: ";
                    xSheet.Cells[Counter + 3 + 9, 1] = "Total Passes: ";
                    xSheet.Cells[Counter + 4 + 9, 1] = "Total Failed: ";

                    xSheet.Cells[Counter + 2 + 9, 2] = sTotalCounter;
                    xSheet.Cells[Counter + 3 + 9, 2] = sTotalPassed;
                    xSheet.Cells[Counter + 4 + 9, 2] = sTotalFailed;

                    xSheet.Cells[Counter + 5 + 9, 2] =  "Project is: "+ sProjectFile;

                    // Add Legende
                    xSheet.Cells[Counter + 6 + 9, 2] = "Legende";
                    xRange = xApp.get_Range("B" + (Counter + 6 + 9));
                    //xRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                    xSheet.Cells[Counter + 7 + 9, 2] ="Pass";
                    xRange = xApp.get_Range("B" + (Counter + 7 + 9));
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_GREEN;
                    xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                   xSheet.Cells[Counter + 8 + 9, 2] ="Fail";
                    xRange = xApp.get_Range("B" + (Counter + 8 + 9));
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_RED;
                    xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                    xSheet.Cells[Counter + 9 + 9, 2] ="Exception";
                    xRange = xApp.get_Range("B" + (Counter + 9 + 9));
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_PINK;
                    xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                    xSheet.Cells[Counter + 10 + 9, 2] ="Untested";
                    xRange = xApp.get_Range("B" + (Counter + 10 + 9));
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_YELLOW;
                    xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }

                if (!sOnlyUITest)
                {
                    if (TFSConnected)
                    {
                        Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                           TestTools.TfsUtilities.GetProjectName("Etricc UI"), sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                        if (quality.Equals("GUI Tests Failed"))
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                        }
                        else
                        {
                            if (sTotalFailed == 0)
                                TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName("Etricc UI"),
                                "GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
                            else
                                TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName("Etricc UI"),
                                "GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
                        }

                        // update testinfo file
                        string testout = "-->" + sOutFilename + ".xls";
                        if (sAutoTest)
                        {
                            if (sTotalFailed == 0)
                            {
                                if (sParentProgram.StartsWith("TFS"))
                                {
                                    FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Passed", ConstCommon.ETRICCUI + "+" + sCurrentPlatform + "Normal");
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Passed:" + ConstCommon.ETRICCUI, sOnlyUITest);
                                }
                                else
                                {
                                    FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Passed" + testout, ConstCommon.ETRICC_UI);
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Passed:" + ConstCommon.ETRICC_UI, sOnlyUITest);
                                }   
                            }
                            else
                            {
                                if (sParentProgram.StartsWith("TFS"))
                                {
                                    FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Failed", ConstCommon.ETRICCUI + "+" + sCurrentPlatform + "Normal");
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Failed:" + ConstCommon.ETRICCUI, sOnlyUITest);
                                }
                                else
                                {
                                    FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Failed" + testout, ConstCommon.ETRICC_UI);
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Failed:" + ConstCommon.ETRICC_UI, sOnlyUITest);
                                }
                            }
                        }
                    }

                    if (sAutoTest)
                    {
                        if (sTotalFailed == 0 && sFunctionalTest == true && sBuildDef.ToLower().StartsWith("nightly"))
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, "Working file not updated now, will continue do Functional testing", sOnlyUITest);
                        }
                        else
                            FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                    }
                }

                #region Save excel file and send email
                xSheet.Columns.AutoFit();
                xSheet.Rows.AutoFit();
               
                // save Excel to Local machine
                string sXLSPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                    sOutFilename + ".xls");

                Epia3Common.WriteTestLogMsg(slogFilePath, "Save excel to local : " + sXLSPath, sOnlyUITest);
                // Save the Workbook locally  --- not for PC EPIATESTPC3
                object missing = System.Reflection.Missing.Value;
                
                xBook.SaveAs(sXLSPath, Excel.XlFileFormat.xlWorkbookNormal,
                                missing, missing, missing, missing,
                                Excel.XlSaveAsAccessMode.xlNoChange,
                                missing, missing, missing, missing, missing);

                // Save to remote machine
                if (System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().StartsWith("TEAMSYSTEMS\\JIEMINSHI"))
                {
                    Console.WriteLine("\n   not write to remote machine");
                }
                else
                {
                    //string ReportDirectory = sBuildDropFolder + "\\TestResults";
                    //string dir = System.IO.Path.GetDirectoryName(sTestResultFolder);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "ReportDirectory : " + sTestResultFolder, sOnlyUITest);

                    string sXLSPath3 = System.IO.Path.Combine(sTestResultFolder, sOutFilename + ".xls");
                       
                    //string sXLSPath2 = System.IO.Path.Combine(sOutFilePath, sOutFilename + ".xls");
                    //Epia3Common.WriteTestLogMsg(slogFilePath, "sOutFilePath : " + sOutFilePath, sOnlyUITest);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Save excel to server : " + sXLSPath3, sOnlyUITest);


                    xBook.SaveAs(sXLSPath3, Excel.XlFileFormat.xlWorkbookNormal,
                                missing, missing, missing, missing,
                                Excel.XlSaveAsAccessMode.xlNoChange,
                                missing, missing, missing, missing, missing);
                }
                // quit Excel.
                if (xBook != null) xBook.Close(true, sOutFilename, false);
                if (xBooks != null) xBooks.Close();
                xApp.Quit();

                releaseObject(xSheet);
                releaseObject(xBook);
                releaseObject(xApp);

                // Send Result via Email
                SendEmail(sXLSPath);
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                try
                {
                    if (sAutoTest)
                    {
                        if (sParentProgram.StartsWith("TFS"))
                        {
                            FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Exception -->" + sOutFilename + ".log", ConstCommon.ETRICCUI + "+" + sCurrentPlatform + "Normal");
                            Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Exception -->" + sOutFilename + ".log" + ConstCommon.ETRICCUI, sOnlyUITest);
                        }
                        else
                        {
                            FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Exception -->" + sOutFilename + ".log", ConstCommon.ETRICC_UI);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Exception -->" + sOutFilename + ".log :: " + ConstCommon.ETRICC_UI, sOnlyUITest);

                        }   

                        
                        Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace, sOnlyUITest);

                        Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                        Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                        Utilities.CloseProcess("cmd");
                        FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");

                        if (TFSConnected)
                        {
                            Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                            TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICC_UI), sBuildNr);
                            string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                            if (quality.Equals("GUI Tests Failed"))
                            {
                                Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                            }
                            else
                            {
                                TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                    TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICC_UI),
                                    "GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
                            }
                        }
                    }
                }
                catch (Exception)
                { }
               
                return;                  
            }

            try
            {
                #region // Functional Testing
                if (sTotalFailed == 0 && sFunctionalTest == true && sBuildDef.ToLower().StartsWith("nightly")) 
                {
                    Console.WriteLine("Start Functional Testing: ");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Start Functional Testing... ", sOnlyUITest);

                    Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                    Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);

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
                    string msg = "Recompile TestRuns:delete files:"+deletePathDll;
                    if (!FileManipulation.DeleteFilesWithWildcards(deletePathDll, ref msg))
                        throw new Exception(msg);

                    string deletePathPdb = @"C:\EtriccTests\TestRuns\bin\debug\Egemin*.pdb";
                    string msg1 = "Recompile TestRuns:delete files:" + deletePathDll;
                    if (!FileManipulation.DeleteFilesWithWildcards(deletePathPdb, ref msg1))
                        throw new Exception(msg1);
                    
                    Thread.Sleep(3000);

                    string origPath = sEtriccServerRoot+@"\Egemin*.dll";
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
                    string output = Utilities.RunProcessAndGetOutput(exePath, arg);
                    if (output.IndexOf("error") >= 0 )
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
                        Utilities.CloseProcess("EXCEL");
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
                                        Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                                        TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICC_UI), sBuildNr);
                                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                                        //if (quality.Equals("GUI Tests Failed"))
                                        //{
                                        //    Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                                        //}
                                        //else
                                        //{
                                        TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                            TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICC_UI),
                                            "Functional Tests Started", m_BuildSvc, sDemo ? "true" : "false");
                                        //}
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message+"---"+ ex.StackTrace, "TFSConnected Exception");
                                        Console.WriteLine(ex.Message);
                                        Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message+"-final check -"+ex.StackTrace, sOnlyUITest);
                                    }
                                }
                            }

                            // uninstall Egemin.Epia.server Service
                            Console.WriteLine("UNINSTALL EPIA SERVER Service : ");
                            TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
                                ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /u");
                            Thread.Sleep(2000);

                            // uninstall Egemin.Etricc.server Service
                            Console.WriteLine("UNINSTALL ETRICC SERVER Service : ");
                            TestTools.Utilities.StartProcessWaitForExit(sEtriccServerRoot,
                                ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /u");
                            Thread.Sleep(2000);

                            Console.WriteLine("INSTALL EPIA SERVER Service : ");
                            TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
                                ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /i");
                            Thread.Sleep(2000);

                            Console.WriteLine("INSTALL ETRICC SERVER Service : ");
                            TestTools.Utilities.StartProcessWaitForExit(sEtriccServerRoot,
                                ConstCommon.EGEMIN_ETRICC_SERVER_EXE, " /i");
                            Thread.Sleep(2000);

                            Console.WriteLine("Start EPIA SERVER as Service : ");
                            TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
                                ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /start");
                            Thread.Sleep(2000);

                            Console.WriteLine("Start ETRICC SERVER as Service : ");
                            TestTools.Utilities.StartProcessWaitForExit(sEtriccServerRoot,
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
                            TestTools.Utilities.StartProcessNoWait(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT, 
                                ConstCommon.EGEMIN_EPIA_SERVER_EXE, string.Empty);

                            // Start Etricc SERVER as Console
                            TestTools.Utilities.StartProcessNoWait(sEtriccServerRoot, ConstCommon.EGEMIN_ETRICC_SERVER_EXE, string.Empty);
                            Thread.Sleep(90000);
                        }
                        #endregion
                        Console.WriteLine("----- time 90 seconds .....");
                        Thread.Sleep(15000);

                        sStartTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - sStartTime;

                        //========================   SHELL =================================================
                        #region  Shell
                        AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
                        // Add Open window Event Handler
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);
                        sEventEnd = false;
                        TestCheck = ConstCommon.TEST_PASS;

                        Thread.Sleep(45000);

                        // Start Shell
                        TestTools.Utilities.StartProcessNoWait(m_SystemDrive + ConstCommon.EPIA_CLIENT_ROOT, 
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
                        MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Start Server and Shell");
                    }
                        //------------------------
                        System.Diagnostics.Process proc = null;
                        int pIDx = Utilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out proc);
                        Console.WriteLine("Proc ID:" + pIDx);

                        Thread MyNewThread = new Thread(new ThreadStart(ClickScreenThreadProc));
                        MyNewThread.Start();
                   
                    //-------------------------------- Check Functional Testing End?
                    string directoryPath = ConstCommon.ETRICC_TESTS_DIRECTORY;
                    string fileNameExcel = "*.xls";
                    string[] allFilesExcel = System.IO.Directory.GetFiles(directoryPath, fileNameExcel);
                    while(allFilesExcel.Length == 0)
                    {
                        Thread.Sleep(30000);
                        Console.WriteLine("No xls files found   :"+System.DateTime.Now);
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
                        Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message+"-final check -"+ex.StackTrace, sOnlyUITest);
                    }

                    xAppFunc.Quit();

                    if (!sOnlyUITest)
                    {
                        if (TFSConnected)
                        {
                            try
                            {
                                Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                                    TestTools.TfsUtilities.GetProjectName(ConstCommon.ETRICC_UI), sBuildNr);
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
                                    TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                        TestTools.TfsUtilities.GetProjectName(ConstCommon.ETRICC_UI),
                                        "Functional Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
                                else
                                    TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                        TestTools.TfsUtilities.GetProjectName(ConstCommon.ETRICC_UI),
                                        "Functional Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
                                 //}


                             }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                                Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message+"-final check -"+ex.StackTrace, sOnlyUITest);
                            }

                        }
                    }

                }  //end sTotalFailed == 0 || sFunctionalTest == true
                #endregion
                
                // Close LogFile
                Epia3Common.CloseTestLog(slogFilePath, sOnlyUITest);

                Console.WriteLine("\nClosing application in 10 seconds");
                if (sOnlyUITest)
                    Thread.Sleep(10000000);
                else
                    Thread.Sleep(10000);
                
                // close CommandHost
                Thread.Sleep(10000);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                Console.WriteLine("\nEnd test run\n");

                if (sAutoTest)
                {
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Working file updated now, Functional testing finished ", sOnlyUITest);
                    FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                if (sAutoTest)
                {
                    if (sParentProgram.StartsWith("TFS"))
                    {
                        FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "FUNCTIONAL Tests Exception -->" + sOutFilename + ".log", ConstCommon.ETRICCUI + "+" + sCurrentPlatform + "Normal");
                        Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.ETRICCUI, sOnlyUITest);
                    }
                    else
                    {
                        FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "FUNCTIONAL Tests Exception -->" + sOutFilename + ".log", ConstCommon.ETRICC_UI);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.ETRICC_UI, sOnlyUITest);
                    }   

                    
                    Epia3Common.WriteTestLogFail(slogFilePath, "FUNCTIONAL Tests Exception -->" + sOutFilename + ".log:"+ConstCommon.ETRICC_UI, sOnlyUITest);

                    Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                    Utilities.CloseProcess("cmd");
                    FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");

                    if (TFSConnected)
                    {
                        Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                                TestTools.TfsUtilities.GetProjectName("Epia"), sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;
                        if (quality.Equals("Functional Tests Failed"))
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                        }
                        else
                        {
                            TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName("Epia"),
                                "Functional Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
                        }
                    }
                }
            }
        }

        private static void releaseObject(object obj)
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
        }

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
        
        #region Excel ------------------------------------------------------------------------------------------------
        public static void WriteResult(int result, int counter, string name,
            Excel.Worksheet sheet, string errorMSG)
        {
            string time = System.DateTime.Now.ToString("HH:mm:ss");
            xSheet.Cells[counter + 2 + 9, 1]= time;
            xSheet.Cells[counter + 2 + 9, 2]= name;
            xSheet.Cells[counter + 2 + 9, 3]= errorMSG;

            xRange = sheet.get_Range("A" + (Counter + 2 + 9), "A" + (Counter + 2 + 9));
            xRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

            xRange = sheet.get_Range("C" + (Counter + 2 + 9), "C" + (Counter + 2 + 9));
            xRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

            xRange = sheet.get_Range("B" + (Counter + 2 + 9), "B" + (Counter + 2 + 9));
            xRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            switch (result)
            {
                case ConstCommon.TEST_PASS:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_GREEN;
                    break;
                case ConstCommon.TEST_FAIL:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_RED;
                    break;
                case ConstCommon.TEST_EXCEPTION:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_PINK;
                    break;
                case ConstCommon.TEST_UNDEFINED:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_YELLOW;
                    //xSheet.Cells.set_Item(row, 3, testData);
                    break;
            }
        }
        #endregion Excel +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
        #region SystemOverviewQuery
        public static void SystemOverviewQuery(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            TestCheck = ConstCommon.TEST_PASS;
            result = ConstCommon.TEST_UNDEFINED;

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
                    sErrorMessage = SYSTEM_OVERVIEW+" Window not found";
                    Console.WriteLine(sErrorMessage);
                }
                else
                {
                    string ms = SYSTEM_OVERVIEW+" window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
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
                    tranform.Resize(Width, Height-60);

                    Thread.Sleep(3000);

                    if (root.Current.BoundingRectangle.Width == Width &&
                        root.Current.BoundingRectangle.Height == (Height-60) )
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
                    }
                }
                #endregion

                #region open query screen
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
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);
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
                    AutomationElement aeQueryRouteCostTrackingItem = aeQueryWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cQueryRouteCostTrackingItem);
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
            }
        }
        #endregion SystemOverviewQuery

        #region SystemOverviewDisplay
        public static void SystemOverviewDisplay(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

           
            AutomationElement aePanelLink = null;
            AutomationElement aeWindow = null;
            AutomationElement aeOverview = null;
            try
            {
                #region // open System overview
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
                #endregion

                #region// Find System Overview Window
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Condition cWindow = new AndCondition( new PropertyCondition(AutomationElement.NameProperty, SYSTEM_OVERVIEW_TITLE),
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

                #region // Validate LegendeInfo UI
                AutomationElement aeLegendeButton = null;
                AutomationElement aeLegendeTreeView = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // Map legende
                    string legendeBtnId = "LegendeButton";
                    string mapLegendTreeViewId = "MapLegend"; // check IsOffScreen 
                    aeLegendeButton = AUIUtilities.FindElementByID( legendeBtnId, aeOverview );
                    aeLegendeTreeView = AUIUtilities.FindElementByID(mapLegendTreeViewId, aeOverview);
                    if ( aeLegendeButton == null || aeLegendeTreeView == null )
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
                    aeWindow =  EtriccUtilities.GetCategoryWindow("System overview", ref sErrorMessage);
                    aeLegendeTreeView = AUIUtilities.FindElementByID("MapLegend", aeWindow);
                    if (aeLegendeTreeView.Current.IsOffscreen == true)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "aeLegendeTreeView is stell OffScreen after legende button clicked";
                        Console.WriteLine(sErrorMessage);
                    }
                    else
                    {
                        Condition cTreeItem = new AndCondition( new PropertyCondition(AutomationElement.IsEnabledProperty, true),
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
                            }
                            //elementNode = walker.GetNextSibling(elementNode);
                        }
                    }
                }
                #endregion


                // check mini map                
                AutomationElement aeMiniMap = null;

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
        
        #region AgvOverviewDisplay
        public static void AgvOverviewDisplay(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

            AutomationElement aePanelLink = null;
            try
            {
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

                Thread.Sleep(10000);
                 // Find Location Overview Window
                System.Windows.Automation.Condition c = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, LOCATION_OVERVIEW_TITLE),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window) );

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
                    aeSelectWindow = AUIUtilities.FindElementByID(selectWindowId,root);
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

            AutomationElement aeOverview = null;
            AutomationElement aePanelLink = null;
            if (sOnlyUITest)
                root = EtriccUtilities.GetMainWindow("MainForm");

            try
            {
                AUICommon.ClearDisplayedScreens(root);
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
        #region TransportOverviewDisplay
        public static void TransportOverviewDisplay(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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

            try
            {
                string epiaDataResourceFolder = OSVersionInfoClass.ProgramFilesx86() + "\\Egemin\\Epia Server\\Data\\Resources";
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
        #region AgvOverviewOpenDetail
        public static void AgvOverviewOpenDetail(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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
                    Input.MoveToAndClick(aePanelLink);

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
                
                // Find AGV GridView
                AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
                if (aeGrid == null)
                {
                    sErrorMessage = "Find AgvDataGrid failed:" + AGV_GRIDDATA_ID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Console.WriteLine("Button AgvDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

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
                    sErrorMessage = "AgvDataGrid aeCell Value not found:" + cellname;
                    Console.WriteLine("AgvDataGridView aeCell Value not found:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                // Open AGV Detail Screen
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
                    Epia3Common.WriteTestLogMsg(slogFilePath, testname +": " + ex.Message + "---" + ex.StackTrace, sOnlyUITest);
                    return;
                }

                // Open Detail screen
                Input.MoveToAndDoubleClick(point);
                Thread.Sleep(2000);
 
                // Check AGV Detail Screen Opened
                string detailScreenName = "Agv detail - " + AgvValue;
                System.Windows.Automation.Condition c2 = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, detailScreenName),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                    );

                AutomationElement aeDetailScreen = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                if (aeDetailScreen == null)
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = detailScreenName + " Window not found";
                    Console.WriteLine("FindElementByName failed:" + detailScreenName);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    return;
                }
                
                // Check AGV text value
                string textID = "m_IdValueLabel";
                AutomationElement aeAgvText = AUIUtilities.FindElementByID(textID, aeDetailScreen);
                if (aeAgvText == null)
                {
                    sErrorMessage = "Find AgvTextElement failed:" + textID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage,sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                string agvTextValue = aeAgvText.Current.Name;
                if ( agvTextValue.Equals(AgvValue))
                {
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
                else
                {
                    result = ConstCommon.TEST_FAIL;
                    sErrorMessage = detailScreenName + "Agv Value should be " + AgvValue + ", but  " + agvTextValue;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath,"===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        
        #region LocationOverviewOpenDetail
        public static void LocationOverviewOpenDetail(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage,sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                        return;
                    }

                    string locTextValue = aeLocText.Current.Name;
                    if ( locTextValue.Equals(cellValue))
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

        #region LocationModeManual
        public static void LocationModeManual(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                while (!StateValue.Equals("Manual") && mTime.Seconds < 30)
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
        
        #region RestartAgv
        public static void RestartAgv(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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

                // Find AGV GridView
                AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
                if (aeGrid == null)
                {
                    sErrorMessage = "Find AgvDataGrid failed:" + AGV_GRIDDATA_ID;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                    Console.WriteLine("AgvDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

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
                System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
               
                Thread.Sleep(3000);

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
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("cellState AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                // find cellState value
                string StateValue = string.Empty;
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
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                Thread.Sleep(3000);
                System.Windows.Point point2 = AUIUtilities.GetElementCenterPoint(aeCellState);
                
                Thread.Sleep(3000);
                //______________________________________________________________________________
                /*string StateValue = AUICommon.GetDataGridViewCellValueAt(0, "State", aeGrid);

                if (StateValue == null || StateValue == string.Empty)
                {
                    sErrorMessage = "AgvDataGridView aeCell Value not found:" + "State Row 0";
                    Console.WriteLine("AgvDataGridView aeCell Value not found:" + "State Row 0");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + "State Row 0", sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                */
                Input.MoveToAndRightClick(point);
                Thread.Sleep(3000);

                // find Agv Restart menu item point
                System.Windows.Automation.Condition cAgvRestart = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, Constants.AGV_MENUITEM_Restart_Agv),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
               );

                // Find the MenuItemMode element
                AutomationElement aeMenuItemRestartAgv = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cAgvRestart);
                if (aeMenuItemRestartAgv != null)
                {
                    Console.WriteLine("new element found: " + aeMenuItemRestartAgv.Current.Name);
                    Console.WriteLine("err.Current.AutomationId: " + aeMenuItemRestartAgv.Current.AutomationId);
                    Console.WriteLine(" err.Current.ControlType.ToString(): " + aeMenuItemRestartAgv.Current.ControlType.ToString());
                    Console.WriteLine(" err.Current.ControlType.ProgrammaticName: " + aeMenuItemRestartAgv.Current.ControlType.ProgrammaticName);
                    Input.MoveToAndClick(aeMenuItemRestartAgv);
                }
                else
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Restart Agv menu iteme not found ------------:";
                    Console.WriteLine(sErrorMessage);
                    return;
                }

                Thread.Sleep(2000);

                // Find Restart Agv Dialog Window
                System.Windows.Automation.Condition c2 = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, "Restart Agvs?"),
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
                Thread.Sleep(5000);

                StateValue = AUICommon.GetDataGridViewCellValueAt(0, "State", aeGrid);
                sStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - sStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                while (!StateValue.Equals("Ready") && mTime.Seconds < 30)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - sStartTime;
                    StateValue = AUICommon.GetDataGridViewCellValueAt(0, "State", aeGrid);
                    Console.WriteLine("time is (sec) : " + mTime.Seconds + " and state is " + StateValue);
                }

                if (StateValue.Equals("Ready"))
                {
                    result = ConstCommon.TEST_PASS;
                    Console.WriteLine(testname + " ---pass --- "+StateValue);
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
                sErrorMessage = sErrorMessage.Trim();
                Console.WriteLine("Fatal error: " + sErrorMessage);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, sOnlyUITest);
            }
        }
        #endregion
        
        #region AgvJobOverview
        public static void AgvJobOverview(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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
        
        #region AgvJobOverviewOpenDetail
        public static void AgvJobOverviewOpenDetail(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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
               
                string cellname = "Id Row "+row;
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

                string JobCellname = "Id Row "+row;
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
                string JobScreenName = "Job detail - " + AgvValue+" - " +JobValue;
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
        
        #region AgvModeSemiAutomatic
        public static void AgvModeSemiAutomatic(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            AutomationElement aePanelLink = null;
            AutomationElement aeOverview = null;
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
                    Thread.Sleep(10000);
                    // Find AGV Overview Window
                    System.Windows.Automation.Condition c = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, AGV_OVERVIEW_TITLE),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                   );

                    // Find the AGVOverview element.
                    aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeOverview == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = TRANSPORT_OVERVIEW + " Window not found";
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
                    aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
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
                            sErrorMessage = Modes[k]+ " not found ------------:";
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
                           new PropertyCondition(AutomationElement.NameProperty, "Put Agvs in mode "+Modes[k]+"?"),
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
        
        #region AgvsAllModeRemoved
        public static void AgvsAllModeRemoved(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                while (!StateValue.Equals("Removed") && mTime.Seconds < 30)
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
                    catch (  System.NullReferenceException)
                    {
                        break;
                    }

                    sStartTime = DateTime.Now;
                    mTime = DateTime.Now - sStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                    while (!StateValue.Equals("Removed") && mTime.Seconds < 30)
                    {
                        Thread.Sleep(2000);
                        mTime = DateTime.Now - sStartTime;
                        StateValue = AUICommon.GetDataGridViewCellValueAt(1, "Mode", aeGrid);
                        Console.WriteLine("time is (sec) : " + mTime.Seconds + " and mode is " + StateValue);
                    }

                    if (StateValue.Equals("Removed"))
                    {
                        result = ConstCommon.TEST_PASS;
                        Console.WriteLine(testname +" "+i+ " de ---pass --- " + StateValue);
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
        
        #region AgvsIdSorting
        public static void AgvsIdSorting(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            
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
                    string cellname = "Id Row "+i;
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
                if (AgvsIdCells[0].CompareTo(AgvsIdCells[1]) < 0 )
                    PreSortAscending = true;
                else
                    PreSortAscending = false;


                Console.WriteLine("Sort Ascending = " + PreSortAscending);
                Thread.Sleep(3000);
                // Click ID Header Cell
                double x = aeGrid.Current.BoundingRectangle.Left;
                double y = aeGrid.Current.BoundingRectangle.Top;
                Point headpoint = new Point(x + 5, y + 5);

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

                Console.WriteLine( " ---Total agvs --- "+sNumAgvs);
                Epia3Common.WriteTestLogMsg(slogFilePath, " ---Total agvs --- " + sNumAgvs, sOnlyUITest);
                        
                bool sortResult = true;
                // Check result
                for (int i = 0; i < sNumAgvs-1; i++)
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

                if ( sortResult )
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
        
        #region CreateNewTransport
        public static void CreateNewTransport(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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

                Console.WriteLine("\n=== Find new Transport ===");
                System.Windows.Forms.TreeNode treeNode = new TreeNode();
                AutomationElement aeNode = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, NEW_TRANSPORT, ref sErrorMessage);
                if (aeNode == null)
                {
                    sErrorMessage = string.Empty;
                    Console.WriteLine("\n=== New Transport NOT Exist ===");
                    Input.MoveToAndDoubleClick(aePanelLink.GetClickablePoint());
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
                    Input.MoveToAndClick(aeNode);

                Thread.Sleep(2000); 

                //find new transport screen
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
                    CommandID   = "Drop";
                    MoverID     = "FLV";
                    SourceID    = "FLV_L101";
                    DestID      = "FLV_L203";
                }
                else if (sProjectFile.ToLower().IndexOf("eurobaltic") >= 0)
                {
                    CommandID   = "Move";
                    MoverID     = "AGV1";
                    SourceID    = "0040-01-01";
                    DestID      = "0030-01-01";
                }
                else
                {
                    CommandID   = "Move";
                    MoverID     = "AGV1";
                    SourceID    = "M_01_01_01_01";
                    DestID      = "ABF_1_1_T";
                }
                
                AutomationElement item
                    = AUIUtilities.FindElementByName(CommandID, aeCommandList);
                if (item != null)
                {
                    Console.WriteLine(CommandID+" item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
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
        #endregion CreateNewTransport
        
        #region EditTransport
        public static void EditTransport(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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
                    Console.WriteLine("TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

                Thread.Sleep(3000);
                // Construct the Grid Cell Element Name
                string cellname = "Id Row 0";
                // Get the Element with the Row Col Coordinates
                AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                if (aeCell == null)
                {
                    Console.WriteLine("Find TransportDataGridView aeCell failed:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView cell failed:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("cell TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                Thread.Sleep(3000);
                System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
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
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = "Edit Transport menu item not found ------------:";
                    Console.WriteLine(sErrorMessage);
                    return;
                }

                Thread.Sleep(2000);
                
                //find edit transport screen
                System.Windows.Automation.Condition c2 = new AndCondition(
                  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );

                // Find the NewTransport Screen element.
                AutomationElement aeNewT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
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
                while (aeDestRadio == null && mTime.TotalMilliseconds < 120000)
                {
                    aeDestRadio = AUIUtilities.FindElementByID(destLocId, aeNewT);
                    Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
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
                
                // Find Save element.
                string id = "m_btnSave";
                AutomationElement aeCancel = AUIUtilities.FindElementByID(id, aeNewT);

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
                    sErrorMessage = " Transport destination is not changed to "+DestID+" , but:" + StateValue;
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


        }
        #endregion EditTransport
        
        #region CancelTransport
        public static void CancelTransport(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            AutomationElement aePanelLink = null;
            AutomationElement aeOverview = null;
            AutomationElement aeGrid = null;
            try
            {
                if (sOnlyUITest)
                    root = EtriccUtilities.GetMainWindow("MainForm");

                AUICommon.ClearDisplayedScreens(root);
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


                AutomationElement aeCell = null;
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
                        string cellname = "Id Row 0";
                        // Get the Element with the Row Col Coordinates
                        aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                        if (aeCell == null)
                        {
                            sErrorMessage = "Find TransportDataGridView aeCell failed:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Thread.Sleep(3000);
                            System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
                            //System.Windows.Point point = AUICommon.GetDataGridViewCellPointAt(0, "Id", aeGrid);
                            // find cell value
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
                                sErrorMessage = "TransportDataGridView aeCell Value not found:" + cellname;
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


                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // find Cancel Transport point
                    System.Windows.Automation.Condition cM = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, Constants.TRANSPORT_MENUITEM_Cancel_Transport),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
                   );

                    // Find the MenuItem Cancel Transport element
                    AutomationElement aeMenuItemTransport = root.FindFirst(TreeScope.Element | TreeScope.Descendants, cM);
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

        #region SuspendTransport
        public static void SuspendTransport(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            TestCheck = ConstCommon.TEST_PASS;
            result = ConstCommon.TEST_UNDEFINED;
            sTransportSuspendOK = true;

            AutomationElement aePanelLink = null;
            AutomationElement aeOverview = null;
            try
            {
                #region // open transport overview
                AUICommon.ClearDisplayedScreens(root);
                aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
                if (aePanelLink == null)
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck  = ConstCommon.TEST_FAIL;
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
                AutomationElement aeCell = null;
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
                        string cellname = "Id Row 0";
                        // Get the Element with the Row Col Coordinates
                        aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                        if (aeCell == null)
                        {
                            sErrorMessage = "Find TransportDataGridView aeCell failed:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine("cell TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(3000);
                            cellPoint = AUIUtilities.GetElementCenterPoint(aeCell);
                            // find cell value
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
                                sErrorMessage = "TransportDataGridView aeCell Value not found:" + cellname;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, "TransportDataGridView cell value not found:" + cellname, sOnlyUITest);
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
                    Input.MoveToAndRightClick(cellPoint);
                    Thread.Sleep(3000);

                    // find Suspend Transport point
                    Condition cM = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, Constants.TRANSPORT_MENUITEM_Suspend_Transport),
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
                            Thread.Sleep(3000);
                        }
                    }
                }

                // validation Suspending
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // suspend transport menu action
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
                        string editId = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.AddEditTransport";
                        aeEditTransportWindow = AUIUtilities.FindElementByID(editId,aeWindow);
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


        }
        #endregion SuspendTransport

        #region ReleaseTransport
        public static void ReleaseTransport(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            TestCheck = ConstCommon.TEST_PASS;
            result = ConstCommon.TEST_UNDEFINED;

            if (sTransportSuspendOK == false)
            {
                sErrorMessage = "Transport suspend test failed, Transport release cannot be tested";
                return;
            }

            AutomationElement aePanelLink = null;
            AutomationElement aeOverview = null;
            try
            {
                #region // open transport overview
                AUICommon.ClearDisplayedScreens(root);
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
                AutomationElement aeCell = null;
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
                        string cellname = "Id Row 0";
                        // Get the Element with the Row Col Coordinates
                        aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                        if (aeCell == null)
                        {
                            sErrorMessage = "Find TransportDataGridView aeCell failed:" + cellname;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine("cell TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(3000);
                            cellPoint = AUIUtilities.GetElementCenterPoint(aeCell);
                            // find cell value
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
                                sErrorMessage = "TransportDataGridView aeCell Value not found:" + cellname;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, "TransportDataGridView cell value not found:" + cellname, sOnlyUITest);
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
                    Input.MoveToAndRightClick(cellPoint);
                    Thread.Sleep(3000);

                    // find Suspend Transport point
                    Condition cM = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, Constants.TRANSPORT_MENUITEM_Release_Transport),
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
                            Thread.Sleep(3000);
                        }
                    }
                }

                // validation Suspending
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // suspend transport menu action
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
                        string editId = "Dialog - Egemin.Etricc.Presentation.ShellModule.Screens.AddEditTransport";
                        aeEditTransportWindow = AUIUtilities.FindElementByID(editId, aeWindow);
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


        }
        #endregion ReleaseTransport

        #region TransportOverviewOpenDetail
        public static void TransportOverviewOpenDetail(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;

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
                string cellname = "Id Row 0";
                // Get the Element with the Row Col Coordinates
                AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

                if (aeCell == null)
                {
                    Console.WriteLine("Find TransportDataGridView aeCell failed:" + cellname);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView cell failed:" + cellname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }
                else
                {
                    Console.WriteLine("cell TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                }

                Thread.Sleep(3000);
                System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);
               
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
                    Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage,sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

                string TrnTextValue = aeTrnText.Current.Name;
                if ( TrnTextValue.Equals(cellValue))
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

        #region ScriptBeforeAfterActivate
        public static void ScriptBeforeAfterActivate(string testname, AutomationElement root, out int result)
        {
            

            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sBeforeAfterActivateScriptOK = true;
            string EtriccServerPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Etricc Server";

            if (sBuildNr.IndexOf("Dev01") > 0
                || sBuildNr.IndexOf("Dev02") > 0
                || sBuildNr.IndexOf("Dev04") > 0
                 || sBuildNr.IndexOf("Dev07") > 0
                || sBuildNr.IndexOf("Main") > 0
                )
                Console.WriteLine("will be tested : " + sBuildNr);
            else
                return;
                 
            try
            {
                #region // unzip test scripts first
                if (System.IO.Directory.Exists(System.IO.Directory.GetCurrentDirectory() + @"\EtriccShell"))
                {
                    Console.WriteLine("EtriccShell folder exist, do nothing: ");
                }
                else
                {
                    Console.WriteLine("EtriccShell folder not exist, unzip test scripts data: ");
                    try
                    {
                        // unzip project file
                        //string zipFile = EtriccStatistics.zip;
                        string zipFile = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "EtriccShell.zip");
                        FastZip fz = new FastZip();
                        fz.ExtractZip(zipFile, System.IO.Directory.GetCurrentDirectory(), "");
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show("EtriccShell.zip unzip exception: " + ex.Message + "--------------" + ex.StackTrace);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                #endregion
                
                string EtriccServerDataPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Etricc Server\Data\Etricc";
                string EtriccCurrentDataPath = System.IO.Directory.GetCurrentDirectory() + @"\EtriccShell\Scripts";
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
                    TestTools.Utilities.StartProcessWaitForExit(EtriccServerPath, 
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
                    Console.WriteLine(" time is :" + sTime.TotalMilliseconds);
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
                        string EtriccPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Etricc Server";
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
                    TestTools.Utilities.StartProcessWaitForExit(EtriccServerPath,
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
                    Console.WriteLine(" time is :" + sTime.TotalMilliseconds);
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
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window) );

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

        #region ScriptBeforeAfterDeactivate
        public static void ScriptBeforeAfterDeactivate(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            string EtriccServerPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Etricc Server";

            if (sBuildNr.IndexOf("Dev01") > 0
                || sBuildNr.IndexOf("Dev02") > 0
                 || sBuildNr.IndexOf("Dev04") > 0
                 || sBuildNr.IndexOf("Dev07") > 0
                || sBuildNr.IndexOf("Main") > 0
                )
                Console.WriteLine("will be tested : " + sBuildNr);
            else
                return;


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
                    TestTools.Utilities.StartProcessWaitForExit(EtriccServerPath,
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
                    Console.WriteLine(" time is :" + sTime.TotalMilliseconds);
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
                string EtriccPath2 = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Etricc Server";
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
                    TestTools.Utilities.StartProcessWaitForExit(EtriccServerPath,
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
                    Console.WriteLine(" time is :" + sTime.TotalMilliseconds);
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

        #region Epia4Close
        public static void Epia3Close(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); 
            result = ConstCommon.TEST_UNDEFINED;
            string BtnCloseID = "Close";
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
                int pID = Utilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SHELL, out proc);
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
        #region Event
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
            else if (name.Equals(m_SystemDrive+"WINDOWS\\system32\\cmd.exe"))
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

        #endregion

        public static void SendEmail(string resultFile)
        {
            string str1 = "<html><body><b><center>Test Overview</center></b><br><br><table col=" +
                        '"' + "5" + '"' + " > <tr><th></th><th>Total Tests:</th><th>&nbsp;</th><th>"
                        + sTotalCounter
                        + "</th>  <th></th> </tr>"
                        + "<tr><td><br></td><td></td><td></td><td></td><td></td></tr>"
                        + "                            <tr><td></td>	<td>Pass:     </td> <td></td>       <td>"
                        + sTotalPassed
                        + "</td><td></td><td></td></tr><tr><td></td>	<td>Fail:     </td>  <td></td>	    <td>"
                        + sTotalFailed
                        + "</td><td></td><td></td></tr><tr><td></td>	<td>Exception:</td>  <td></td>	    <td>"
                        + sTotalException
                        + "</td><td></td><td></td></tr><tr><td></td>	<td>Untested:</td>   <td></td>	    <td>"
                        + sTotalUntested
                        + "</td><td></td>	<td></td></tr></table><br><br></body></html>";

            string TextStatistics = "       Test Overview   " + System.Environment.NewLine;
            TextStatistics = TextStatistics + "Total Test Cases:     " + sTotalCounter + System.Environment.NewLine;
            TextStatistics = TextStatistics + "Pass:                 " + sTotalPassed + System.Environment.NewLine;
            TextStatistics = TextStatistics + "Fail:                 " + sTotalFailed + System.Environment.NewLine;
            TextStatistics = TextStatistics + "Exception:            " + sTotalException + System.Environment.NewLine;
            TextStatistics = TextStatistics + "Untested:             " + sTotalUntested + System.Environment.NewLine;
            TextStatistics = TextStatistics + System.Environment.NewLine;

            TextStatistics = str1;

            TestTools.Utilities.SendTestResultToDevelopers(resultFile, sProjectFile, sBuildDef, logger, sTotalFailed,
               sBuildNr /*used for email title*/, str1/*content*/, sSendMail);
        }

        public static void CreateOutputFileInfo(string[] args, string currentPlatform, string PCName, ref string outPath, ref string outFilename)
        {

            // out filename 
            outFilename = System.DateTime.Now.ToString("yyyyMMdd-HH-mm-ss") + "-GUITESTS";

            if (args[0].IndexOf("Debug") > 0)
                outFilename = currentPlatform + "Debug-" + outFilename + "-" + PCName;
            else
                outFilename = currentPlatform + "Release-" + outFilename + "-" + PCName;

            if (args[10].ToLower().StartsWith("false"))
                outFilename = "Manual-" + currentPlatform + outFilename;

            outPath = args[1] + "\\TestResults";
        }
    }
}
