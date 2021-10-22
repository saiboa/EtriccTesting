using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.Win32;
using TestTools;

namespace TFSQATestTools
{
    public class Tester
    {
        #region fields
        internal static TestTools.Logger logger = null;
        internal static string sLogFilename = string.Empty;
        internal string sTestDefinitionFilesPath = string.Empty;

        // GUI test params
        static string m_installScriptDir = string.Empty;                        // "0) Install msi file path: "
        static string m_ValidatedBuildDropFolder = string.Empty;                // "1) Build Drop folder: "
        static string m_BuildNumber = string.Empty;                             // "2) build nr: "
        static string mTeamProject = "Epia 4";                              // "3) test Project: "
        static string mTestApp = TestApp.EPIA4;                                  // "4) test Application: "
        static string mTargetPlatform = string.Empty;                           // "5) targeted platform: " --->  AnyCPU x86 X64    AnyCPU+X86
        static string mCurrentPlatform = string.Empty;                          // "6) current platform: "  --->  AnyCPU+X86   --> AnyCPU  than x86
        static string mTestDef = string.Empty;                                  // "7) test def: " -->  CI, Nightly, Weekly, Version
        const string mCalledProgram = "TFSQATesttTool";                         // "8) Called by: "
        internal string TESTTOOL_VERSION = "3.10.08.05";                        // "9) TestTool version: "
        static internal bool m_TestAutoMode = true;                             // "10) Auto test: "
        //static string sTFSServerUrl = Constants.sTFSServerUrl;                // "11) TFSServerUrl: "
        string mServerRunAs = "Service";                                        // "12) Server Run As: "  --> read from configuration file
        string mExcelVisible = "Visible";                                       // "13) Excel Visible: "  --> read from configuration file
        static string sDemonstration = Constants.sDemonstration;                // "14) Demo test: "         --> read from App.config file
        static string mMail = "false";                                          // "15) Mail: "         --> read from configuration file
        static string mTestDefFile = string.Empty;                              // "16) System.IO.Path.Combine(sTestDefinitionFilesPath, m_TFS_Settings.Element.TestDefinition);
        static string sInfoFileKey = string.Empty;                              // "17) Windows7.32.x86.Debug.EPIAAUTOTEST1
        static string sNetworkMap = string.Empty;                               // "18) \\FileCluster2.Ecorp.Int\TFSDROPFOLDER\Builds
        static string sDemoCaseCount = Constants.sDemoCaseCount;
        static int sDemoCaseCountINT = Convert.ToInt32(sDemoCaseCount);          // "19) 

        static bool mInstallOldEtricc5Service = false;
        static bool mInstallEtriccLauncher = false;
        static string sEtriccMsiName = "Etricc ?.msi";

        List<string> mBuildDefs = new List<string>();
        static string mDateFilter = "Today";
        static string mCurrentConfiguration = "Debug";          // Debug, Release, Protected
        static string mTestDefinitionFilename = string.Empty;

        // --- TEST PARAMS
        //private static Settings m_Settings;
        private static TestsConfigSection m_TestConfigSettings = new TestsConfigSection();
        private static TfsSettingsSection m_TFS_Settings = new TfsSettingsSection();
        private static string sCurrentBuildInTesting = "Searching ......";
        static internal string m_TestPC = string.Empty;
        static string m_CurrentDrive = string.Empty;
        static string m_SystemDrive = string.Empty;
        public DateTime sTestStartUpTime;
        static internal string m_TestWorkingDirectory = string.Empty;

        static string sEpia4InstallerName = "Epia.msi";
        private string mEpiaPath = string.Empty;
        internal StringCollection m_Logging = new StringCollection();
        string sMsiRelativePath = string.Empty;
        public static string sProjectFile = string.Empty;

        internal STATE m_State;
        public bool sIsDeployed = false;
        public static DateTime sDeploymentEndTime;
        static int sLogCount = 0;
        static int sLogInterval = 0;
        public static string sMsgDebug = Constants.sMsgDebug;

        private static string sErrorMessage = string.Empty;

        static string sTestResultFolder = string.Empty;
        // tested build info
        BuildObject m_ValidatedTestBuildObject = new BuildObject();
        private Uri m_Uri = null;

        // --- Etricc 5
        static string mPreviousSetupPathEpia = string.Empty;
        static string mCurrentSetupPathEpia = string.Empty;
        static string mPreviousSetupPathEtriccShell = string.Empty;
        static string mCurrentSetupPathEtriccShell = string.Empty;
        static string mPreviousSetupPathEtricc5 = string.Empty;
        static string mCurrentSetupPathEtricc5 = string.Empty;
        static string mPreviousSetupPathKC = string.Empty;
        static string mCurrentSetupPathKC = string.Empty;
        static string mPreviousSetupPathEwms = string.Empty;
        static string mCurrentSetupPathEwms = string.Empty;
        static string mTestRunsDirectory = string.Empty;

        // statistics
        static string mPreviousSetupPathEtriccStatisticsUI = string.Empty;
        static string mCurrentSetupPathEtriccStatisticsUI = string.Empty;
        static string mPreviousSetupPathEtriccStatisticsParser = string.Empty;
        static string mCurrentSetupPathEtriccStatisticsParser = string.Empty;
        static string mPreviousSetupPathEtriccStatisticsParserConfigurator = string.Empty;
        static string mCurrentSetupPathEtriccStatisticsParserConfigurator = string.Empty;

        // --- BUILD
        TfsTeamProjectCollection tfsProjectCollection = null;
        IBuildServer m_BuildSvc;
        private bool TFSConnected = true;

        string sEtricc5InstallationFolder = string.Empty;
        private const string REGKEY = "Software\\Egemin\\Automatic testing\\";
        //static bool sEventEnd = false;

        /// <summary>
        /// network related struct
        /// </summary>
        public struct NETRESOURCEA
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

        [DllImport("mpr.dll")]
        public static extern int WNetAddConnection2A(
            [MarshalAs(UnmanagedType.LPArray)] NETRESOURCEA[] lpNetResource,
            [MarshalAs(UnmanagedType.LPStr)] string lpPassword,
            [MarshalAs(UnmanagedType.LPStr)] string UserName,
            int dwFlags);

        [DllImport("mpr.dll")]
        public static extern int WNetCancelConnection2A(string sharename, int dwFlags, int fForce);
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Events of Tester (1)
        public event EventHandler OnLoggingChanged;
        #endregion // —— Events •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Constructors/Destructors/Cleanup of Tester (1)
        public Tester()
        {
            string logFilename = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "-"
                + Constants.sDeploymentLogFilename;
            ;
            sLogFilename = System.IO.Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, logFilename);
            logger = new Logger(sLogFilename);

            m_State = STATE.PENDING;

            //m_TestWorkingDirectory = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY;
            m_TestWorkingDirectory = System.IO.Directory.GetCurrentDirectory(); ;
            // prepare test directory
            m_CurrentDrive = Path.GetPathRoot(Directory.GetCurrentDirectory());
            string windir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
            m_SystemDrive = Path.GetPathRoot(windir);

            // Epia
            mPreviousSetupPathEpia = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Epia\\Previous";
            mCurrentSetupPathEpia = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Epia\\Current";

            mPreviousSetupPathEtriccShell = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Etricc\\Previous";
            mCurrentSetupPathEtriccShell = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Etricc\\Current";

            mPreviousSetupPathEtricc5 = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Etricc5\\Previous";
            mCurrentSetupPathEtricc5 = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Etricc5\\Current";

            // Statistics 
            mPreviousSetupPathEtriccStatisticsUI = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\EtriccStatisticsUI\\Previous";
            mCurrentSetupPathEtriccStatisticsUI = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\EtriccStatisticsUI\\Current";

            mPreviousSetupPathEtriccStatisticsParser = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\EtriccStatisticsParser\\Previous";
            mCurrentSetupPathEtriccStatisticsParser = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\EtriccStatisticsParser\\Current";

            mPreviousSetupPathEtriccStatisticsParserConfigurator = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\EtriccStatisticsParserConfigurator\\Previous";
            mCurrentSetupPathEtriccStatisticsParserConfigurator = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\EtriccStatisticsParserConfigurator\\Current";

            // KC
            mPreviousSetupPathKC = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\KC\\Previous";
            mCurrentSetupPathKC = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\KC\\Current";
            // EWMS
            mPreviousSetupPathEwms = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Ewms\\Previous";
            mCurrentSetupPathEwms = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Ewms\\Current";

            // create Shared folder c:\EtriccTests
            //ShareFolderPermission("C:\\EtriccTests", "EtriccTests", "");

            mTestRunsDirectory = ConstCommon.ETRICC_TESTS_DIRECTORY + "\\TestRuns\\bin\\Debug";

            if (!Directory.Exists(mPreviousSetupPathEpia))
                Directory.CreateDirectory(mPreviousSetupPathEpia);

            if (!Directory.Exists(mCurrentSetupPathEpia))
                Directory.CreateDirectory(mCurrentSetupPathEpia);
            //-------------ETRICC SHELL-----------------------------
            if (!Directory.Exists(mPreviousSetupPathEtriccShell))
                Directory.CreateDirectory(mPreviousSetupPathEtriccShell);

            if (!Directory.Exists(mCurrentSetupPathEtriccShell))
                Directory.CreateDirectory(mCurrentSetupPathEtriccShell);
            //------------------------------------------------------
            if (!Directory.Exists(mPreviousSetupPathEtricc5))
                Directory.CreateDirectory(mPreviousSetupPathEtricc5);

            if (!Directory.Exists(mCurrentSetupPathEtricc5))
                Directory.CreateDirectory(mCurrentSetupPathEtricc5);
            //--------------------STATISTICS -------------------------
            if (!Directory.Exists(mPreviousSetupPathEtriccStatisticsUI))
                Directory.CreateDirectory(mPreviousSetupPathEtriccStatisticsUI);

            if (!Directory.Exists(mCurrentSetupPathEtriccStatisticsUI))
                Directory.CreateDirectory(mCurrentSetupPathEtriccStatisticsUI);
            //---
            if (!Directory.Exists(mPreviousSetupPathEtriccStatisticsParser))
                Directory.CreateDirectory(mPreviousSetupPathEtriccStatisticsParser);

            if (!Directory.Exists(mCurrentSetupPathEtriccStatisticsParser))
                Directory.CreateDirectory(mCurrentSetupPathEtriccStatisticsParser);
            //---
            if (!Directory.Exists(mPreviousSetupPathEtriccStatisticsParserConfigurator))
                Directory.CreateDirectory(mPreviousSetupPathEtriccStatisticsParserConfigurator);

            if (!Directory.Exists(mCurrentSetupPathEtriccStatisticsParserConfigurator))
                Directory.CreateDirectory(mCurrentSetupPathEtriccStatisticsParserConfigurator);


            //------------------------------------------------------
            if (!Directory.Exists(mPreviousSetupPathEwms))
                Directory.CreateDirectory(mPreviousSetupPathEwms);

            if (!Directory.Exists(mCurrentSetupPathEwms))
                Directory.CreateDirectory(mCurrentSetupPathEwms);

            if (!Directory.Exists(mTestRunsDirectory))
                Directory.CreateDirectory(mTestRunsDirectory);

            if (!Directory.Exists(ConstCommon.ETRICC_TESTS_DIRECTORY))
                Directory.CreateDirectory(ConstCommon.ETRICC_TESTS_DIRECTORY);

            if (!Directory.Exists(@"C:\Etricc 5.0.0\Greatview Aseptic Packaging"))
                Directory.CreateDirectory(@"C:\Etricc 5.0.0\Greatview Aseptic Packaging");

            // Get tfs server 
            try
            {
                sIsDeployed = false;
                sDeploymentEndTime = DateTime.Now;
                m_TestPC = System.Environment.MachineName;

                if (TFSConnected)
                {
                    Log("Connect to TFS");

                    Uri serverUri = new Uri(Tfs.ServerUrl);
                    System.Net.ICredentials tfsCredentials
                        = new System.Net.NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

                    tfsProjectCollection
                        = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                    tfsProjectCollection.Connect(Microsoft.TeamFoundation.Framework.Common.ConnectOptions.IncludeServices);
                    TfsConfigurationServer tfsConfigurationServer = tfsProjectCollection.ConfigurationServer;

                    m_BuildSvc = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));
                }

                //TeamFoundationServer TFS = TeamFoundationServerFactory.GetServer(Constants.sTFSServer);
                WorkItemStore store = (WorkItemStore)tfsProjectCollection.GetService(typeof(WorkItemStore));
                //WorkItemType wiType = store.Projects[8].WorkItemTypes[1];
                // project nedds to be checked 
                //MessageBox.Show("store projects: " + store.Projects[8].ToString() + "  ");

                /*
                int ret = Disconnect(Constants.TEST_DEFINITION_DRIVE_MAP_LETTER);
                if (ret == 0)
                {
                    logger.LogMessageToFile(m_TestPC + "Disconnect MAP DRIVE OK:", sLogCount, sLogInterval);
                }
                else if (ret == 2250)
                    logger.LogMessageToFile(m_TestPC
                        + "Disconnnet: MAP DRIVE The Network connection could not be found :" + ret,
                        sLogCount, sLogInterval);
                else
                    System.Windows.MessageBox.Show("Disconnect  DriveMap failed with error code:" + ret);

                Thread.Sleep(3000);

                ret = OpenDriveMap(@"\\Teamsystem\Team Systems Builds", ConstCommon.DRIVE_MAP_LETTER);
                if (ret == 0)
                {
                    logger.LogMessageToFile(m_TestPC + "Open MAP DRIVE OK:", sLogCount, sLogInterval);
                }
                else if (ret == 85)
                    logger.LogMessageToFile(m_TestPC + "Open MAP DRIVE not connected due to existing connection:", sLogCount, sLogInterval);
                else
                    System.Windows.MessageBox.Show("OpenDriveMap failed with error code:" + ret);
                 */
            }
            catch (TeamFoundationServerUnauthorizedException ex1)
            {
                System.Windows.Forms.MessageBox.Show(ex1.Message + System.Environment.NewLine + ex1.StackTrace, "Tester Constructor");
                Log(ex1.Message + System.Environment.NewLine + ex1.StackTrace);
                TFSConnected = false;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message + System.Environment.NewLine + ex.StackTrace, "Tester Constructor");
                Log(ex.Message + System.Environment.NewLine + ex.StackTrace);
                TFSConnected = false;
            }
        }
        #endregion // —— Constructors/Destructors/Cleanup ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••*

        /// <summary>
        /// Method will start new tests
        /// </summary>
        public void Start(ref DateTime StartUpTime)
        {
            Start2(string.Empty, ref StartUpTime);
        }


        /// <summary>
        /// 1 Get All Builds From TFS depend on the setting from configuration file
        /// 2 Check Build Quality
        /// 3 CheckBuildSucceeding
        /// 4 IsTestWorking
        /// 5 IsThisPCTested
        /// 6 deployment
        /// </summary>
        /// <param name="BuildPath">The build Path</param>
        /// <param name="upTime"></param>
        public void Start2(string manualBuildInfo, ref DateTime upTime)
        {
            try
            {
              
            #region validate input
            try
            {
                
                mBuildDefs.Clear();   // List<string>
                string buildDefs = m_TFS_Settings.Element.BuildDefinitions;
                string[] strArray = buildDefs.Split(';');
                for (int i = 0; i < strArray.Length; i++)
                {    //;Epia.Development.Dev02.CI; in this case buildDefs.count = 3, two of them are empty, should check length > 0
                    if (strArray[i].Length > 0)
                        mBuildDefs.Insert(i, strArray[i]);
                }

                mDateFilter = m_TFS_Settings.Element.DateFilter;
                //mTestDefinitionFilename = m_TFS_Settings.Element.TestDefinition;
                //mTestDefFile = System.IO.Path.Combine(sTestDefinitionFilesPath, m_TFS_Settings.Element.TestDefinition);

                TimeSpan mTime = DateTime.Now - upTime;

                if (mTime.Minutes > 30)
                    sLogInterval = 4;  // 1 min
                else if (mTime.Hours > 1)
                    sLogInterval = 12; // 3 min
                else if (mTime.Hours > 5)
                    sLogInterval = 20; // 5 min
                else if (mTime.Days > 1)
                    sLogInterval = 40;  // 10 min

            }
            catch (Exception ex)
            {
                string msg = "Start2 exception: " + ex.Message + "---" + ex.StackTrace;
                sLogCount++;
                m_State = STATE.PENDING;
                return;
            }
            #endregion 

            string logFilename = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "-"
                + Constants.sDeploymentLogFilename;
            if (!logFilename.Equals(logger.GetLogPath()))   // if another day, create new log file
                logger = new Logger(System.IO.Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, logFilename));

            logger.LogMessageToFile("Start:" + sLogCount, sLogCount, sLogInterval);
            Log("===> searching for available build...");
            Console.WriteLine("===> searching for available build...");

            m_State = STATE.INPROGRESS;

            #region // Get All Buildnrs and DropLocationPaths From TFS depend on the setting from TFSConnection Setting: project, buildDefs and date filter
            if (manualBuildInfo == string.Empty)
            {
                List<BuildObject> allBuilds = null;
                try
                {
                    allBuilds = TfsUtilities.GetAllBuildObjects(mBuildDefs, mDateFilter, this, logger);
                }
                catch (Exception ex)
                {
                    string msg = "BuildUtilities.GetAllBuildObjects: " + ex.Message + "---" + ex.StackTrace;
                    logger.LogMessageToFile(msg, sLogCount, sLogInterval);
                    sLogCount++;
                    m_State = STATE.PENDING;
                    return;
                }

                if (allBuilds.Count == 0)
                {
                    Log("No Any build dircetory found:");
                    logger.LogMessageToFile("No Any build dircetory found:", sLogCount, sLogInterval);
                    //ClickUiScreenActionToAvoidScreenStandBy();
                    Thread.Sleep(18000);
                    sLogCount++;
                    m_State = STATE.PENDING;
                    return;
                }
                else
                {   // log all found builds
                    string allBuildsInfo = string.Empty;
                    logger.LogMessageToFile("(build dircetory found)allBuilds.Count :" + allBuilds.Count, sLogCount, sLogInterval);
                    IEnumerator EmpEnumerator = allBuilds.GetEnumerator(); //Getting the Enumerator
                    EmpEnumerator.Reset(); //Position at the Beginning
                    while (EmpEnumerator.MoveNext()) //Till not finished do print
                    {
                        BuildObject buildObject = (BuildObject)EmpEnumerator.Current;
                        Log(buildObject.BuildNr + "\t" + buildObject.DripLoc);
                        logger.LogMessageToFile(buildObject.BuildNr + "\t" + buildObject.DripLoc, sLogCount, sLogInterval);
                        allBuildsInfo = allBuildsInfo + "\n" + buildObject.BuildNr + "\t" + buildObject.FinishTime;
                    }
                    //System.Windows.Forms.MessageBox.Show("result:"+allBuildsInfo);
                }

                // Get Validated build , that can be tested by thisPC
                // X:\Nightly\Etricc 5\Etricc - Nightly_20100202.1
                //MessageBox.Show("m_Settings.BuildApplication" + m_Settings.BuildApplication);
                // get one application of this build if not yet tested
                try
                {   // m_TestPC will be added in testinfo.txt at first time
                    // ref mCurrentPlatform, ref mCurrentConfiguration, ref sInfoFileKey are read from TestDef file
                    m_ValidatedTestBuildObject = GetNewValidatedBuildDirectory(allBuilds, m_TestPC,
                          ref mCurrentPlatform, ref mCurrentConfiguration, ref sInfoFileKey);
                    //System.Windows.Forms.MessageBox.Show(" sInfoFileKey : " + sInfoFileKey, "xsddsff");
                }
                catch (Exception ex)
                {
                    string msg = "GetValidatedBuildDirectory: " + ex.Message + "---" + ex.StackTrace;
                    sLogCount++;
                    m_State = STATE.PENDING;
                    return;
                }

                if (m_ValidatedTestBuildObject == null)
                {
                    logger.LogMessageToFile("No new validated build found:", sLogCount, sLogInterval);
                    //ClickUiScreenActionToAvoidScreenStandBy();
                    m_State = STATE.PENDING;
                    sLogCount++;
                    return;
                }

                logger.LogMessageToFile("(validated build dircetory):", sLogCount, sLogInterval);

                m_BuildNumber = m_ValidatedTestBuildObject.BuildNr;
                //mTestApp = TfsUtilities.GetTestAppFromBuildNumber(m_BuildNumber);   
                // mTestApp is depend on BuildDef, normally BuildNr = BuildDef_yyyymmdd.X,but sometimes just number like 30903
                mTestApp = TfsUtilities.GetTestAppFromBuildDefinition(m_ValidatedTestBuildObject.BuildDef);
                mTeamProject = TfsUtilities.GetTeamProjectFromTestApp(mTestApp);
                mTargetPlatform = mCurrentPlatform;
                mTestDefFile = @"C:\EtriccTests\QA\TestDefinitions\" + TfsUtilities.GetTestDefNameFromTestApp(mTestApp);

                m_ValidatedBuildDropFolder = m_ValidatedTestBuildObject.DripLoc;

                // mCurrentPlatform: AnyCPU, x86, 
                sMsiRelativePath = "\\Installation\\" + mCurrentPlatform + "\\" + mCurrentConfiguration;
                if (mCurrentPlatform.Equals("AnyCPU"))
                    sMsiRelativePath = "\\Installation\\Any CPU\\" + mCurrentConfiguration;

                // disconnect map first, and reconnect again 
                int ret = Disconnect(ConstCommon.DRIVE_MAP_LETTER);
                if (ret == 0)
                    logger.LogMessageToFile(m_TestPC + "Disconnect MAP DRIVE OK:", sLogCount, sLogInterval);
                else if (ret == 2250)
                    logger.LogMessageToFile(m_TestPC
                        + "Disconnnet: MAP DRIVE The Network connection could not be found :" + ret,
                        sLogCount, sLogInterval);
                else
                    System.Windows.MessageBox.Show("Disconnect  DriveMap failed with error code:" + ret);

                Thread.Sleep(3000);

                // will be optimalised later
                sNetworkMap = m_ValidatedTestBuildObject.xMapString;
                logger.LogMessageToFile("sNetworkMap:" + sNetworkMap, sLogCount, sLogInterval);
                // @"\\Teamsystem.Teamsystems.egemin.be\Team Systems Builds"
                ret = OpenDriveMap(@sNetworkMap, ConstCommon.DRIVE_MAP_LETTER);
                while (!(ret == 0 || ret == 85))
                {
                    TestTools.MessageBoxEx.Show("OpenDriveMap failed with error code:" + ret
                        + "\nNetworkMap:" + sNetworkMap
                        + "\nDRIVE_MAP_LETTER:" + ConstCommon.DRIVE_MAP_LETTER
                        + "\nopen:" + m_ValidatedTestBuildObject.DripLoc,
                        "open:" + m_ValidatedTestBuildObject.DripLoc, 60000);
                    System.Threading.Thread.Sleep(60000);
                    logger.LogMessageToFile("OpenDriveMap failed with error code:" + ret + ":map is: " + sNetworkMap, sLogCount, sLogInterval);
                    ret = OpenDriveMap(@sNetworkMap, ConstCommon.DRIVE_MAP_LETTER);
                }

                // log drive map status
                if (ret == 0)
                    logger.LogMessageToFile(m_TestPC + "Open MAP DRIVE OK:", sLogCount, sLogInterval);
                else if (ret == 85)
                    logger.LogMessageToFile(m_TestPC + "Open MAP DRIVE not connected due to existing connection:", sLogCount, sLogInterval);

                m_ValidatedBuildDropFolder = ConstCommon.DRIVE_MAP_LETTER + "\\" + m_ValidatedTestBuildObject.RelativeDropLoc;
                //m_ValidatedBuildDropFolder = m_ValidatedTestBuildObject.DripLoc;
                if (sMsgDebug.StartsWith("true"))
                {
                    System.Windows.Forms.MessageBox.Show("m_ValidatedBuildDropFolder:" + m_ValidatedBuildDropFolder, "Start2:");
                }

                logger.LogMessageToFile(m_BuildNumber + "\t" + m_ValidatedBuildDropFolder, sLogCount, sLogInterval);

                //X:\CI\Etricc 5\Etricc - CI_20100301.1
                logger.LogMessageToFile("===== <This build will be tested>===== >" + m_ValidatedBuildDropFolder, 0, 0);
                // Tested buildnr Etricc - Nightly_20100202.1

                //m_BuildNumber = BuildUtilities.getBuildnr(m_ValidatedBuildDropFolder);
                Log("testing  build nr: " + m_BuildNumber);
                logger.LogMessageToFile("===== <testing  build nr > m_BuildNumber=: " + m_BuildNumber, 0, 0);

                Log("testing  application : " + mTestApp);
                logger.LogMessageToFile("===== <testing  application >  mTestApp =: " + mTestApp, 0, 0);

                Log("testing  platform : " + mCurrentPlatform);
                logger.LogMessageToFile("===== <testing  application >mCurrentPlatform =: " + mCurrentPlatform, 0, 0);

                // Tested build type. Nightly CI Version
                mTestDef = TfsUtilities.GetTestDefinition(m_ValidatedBuildDropFolder);
                Log("testing definition : " + mTestDef);
                logger.LogMessageToFile("===== <testing  Definition > mTestDef=: " + mTestDef, 0, 0);

                // Current test v. Debug,Release,Protected
                Log("testing  Configuration : " + mCurrentConfiguration);
                logger.LogMessageToFile("===== <testing  Configuration > mCurrentConfiguration=: " + mCurrentConfiguration, 0, 0);

                if (sMsgDebug.StartsWith("true"))
                    System.Windows.Forms.MessageBox.Show("ValidatedBuildDirectory:" + m_ValidatedBuildDropFolder + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);

                m_TestAutoMode = true;
            }
            else
            {
                Log(" manual test starting : " + manualBuildInfo);
                Log(" manual test starting mTeamProject: " + mTeamProject);
                Log(" manual test starting mTargetPlatform: " + mTargetPlatform);
                Log(" manual test starting mTestApp: " + mTestApp);
                m_ValidatedBuildDropFolder = manualBuildInfo;
                #region
                //MessageBox.Show("BuildPath:" + BuildPath);
                m_BuildNumber = manualBuildInfo;
                MessageBox.Show("m_BuildNumber:" + m_BuildNumber);

                    // get build number --> droploc
                    //Uri serverUri = new Uri(Tfs.ServerUrl);
                    //System.Net.ICredentials tfsCredentials
                    //= new System.Net.NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);
                    //TfsTeamProjectCollection tfsProjectCollection = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                    //tfsProjectCollection.EnsureAuthenticated();
                    //IBuildServer buildServer = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));

                    //Uri buildUri = null;
                    // create spec instance and apply the filter like build name, date time
                    //IBuildDetailSpec spec = buildServer.CreateBuildDetailSpec(mTeamProject);  // mTeamProject = Epia 4
                    //    spec.BuildNumber = manualBuildInfo; //Example – “Daily_20110502.4″;
                    //IBuildQueryResult buildDetails = buildServer.QueryBuilds(spec);
                    /*if (buildDetails != null)
                        buildUri = (buildDetails.Builds[0]).Uri;

                    string dropLocation = (buildDetails.Builds[0]).DropLocation;
                    m_ValidatedBuildDropFolder = ConstCommon.DRIVE_MAP_LETTER + "\\" + dropLocation.Substring(55);
                    MessageBox.Show("m_ValidatedBuildDropFolder:" + m_ValidatedBuildDropFolder);

                    mTestDef = TfsUtilities.GetTestDefinition(m_ValidatedBuildDropFolder);
                    MessageBox.Show("mTestDef:" + mTestDef);

                    string platform = "Any CPU";
                    if (mTargetPlatform.Equals("x86"))
                        platform = "x86";
                    else if (mTargetPlatform.Equals("AnyCPU"))
                        platform = "Any CPU";
                    else
                    {
                        System.Windows.Forms.MessageBox.Show(mTargetPlatform + " platform is not allowed for manual testing, please select other platform");
                        return;
                    }
                    */
                    //m_installScriptDir = m_ValidatedBuildDropFolder + "\\\\Installation\\\\" + platform + "\\\\Debug\\\\";
                    m_installScriptDir = m_ValidatedBuildDropFolder;
                MessageBox.Show("m_installScriptDir:" + m_installScriptDir);
                #endregion // end region manual testing
                Log(" manual testing : " + m_BuildNumber);
                logger.LogMessageToFile(" manual testing : " + m_BuildNumber, 0, 0);

                m_TestAutoMode = false;
                TFSConnected = false;
            }

            if (TFSConnected)
            {
                if (sMsgDebug.StartsWith("true"))
                {
                    MessageBox.Show("mTestApp=" + mTestApp);
                    MessageBox.Show("BuildUtilities.GetProjectName(mTestApp)=" + TfsUtilities.GetProjectName(mTestApp));
                }
                //m_Uri = m_buildStore.GetBuildUri(ProjectName(m_testApp), m_BuildNumber);
                //MessageBox.Show("m_BuildSvc" + m_BuildSvc.GetBuildDefinition(GetProjectName(testApp),"").Id);
                try
                {
                    m_Uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, TfsUtilities.GetProjectName(mTestApp), m_BuildNumber);
                }
                catch (Exception ex)
                {
                    string message = "get uri of Build:" + m_BuildNumber + " and TestApp" + mTestApp + " exception :"
                        + ex.Message + "  ----  " + ex.StackTrace;

                    TestTools.MessageBoxEx.Show(message, "get uri exception", 5 * 60000);

                    //throw new Exception(message);
                    m_State = STATE.PENDING;
                    return;
                }

                if (sMsgDebug.StartsWith("true"))
                    MessageBox.Show("m_Uri" + m_Uri);
            }
            #endregion

            sTestResultFolder = m_ValidatedBuildDropFolder + "\\TestResults";

            // Prepare deployment
            //MessageBox.Show("m_testApp:" + m_testApp, "prepare deployment 430");
            #region // Prepare deployment
            sCurrentBuildInTesting = "Prepare deployment: " + m_BuildNumber;
            // We have BuildNumber now, now Deploy application by check m_Settings.Application
            // get m_installScriptDir and then check msi file exist, if not exist 
            // --> log to infofile as GUI Tests Passed and update build qualityand return
            if (mTestApp.Equals(Constants.ETRICC5))
            {
                if (m_TestAutoMode)
                    m_installScriptDir = m_ValidatedBuildDropFolder + sMsiRelativePath;
                ProcessUtilities.CloseProcess(EgeminApplication.EPIA_LAUNCHER);
                ProcessUtilities.CloseProcess(EgeminApplication.EPIA_EXPLORER);
            }
            else if (mTestApp.Equals(TestApp.EPIA4) || mTestApp.Equals(TestApp.EPIANET45)  
                || mTestApp.Equals(Constants.ETRICCUI) || mTestApp.Equals(Constants.ETRICCSTATISTICS))
            {
                if (m_TestAutoMode)
                    m_installScriptDir = m_ValidatedBuildDropFolder + sMsiRelativePath;

                if (sMsgDebug.StartsWith("true"))
                    System.Windows.MessageBox.Show("m_installScriptDir...   " + m_installScriptDir);

                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);
                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EWCS_SERVER);
                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EWCS_TOOLS_DATABASE_FILLER);
                #region check epia.msi exist?
                string EpiaMsiFile = System.IO.Path.Combine(m_installScriptDir, Constants.EPIA_MSI);
                logger.LogMessageToFile("====   check epia.msi exist? : " + EpiaMsiFile, 0, 0);
                if (!System.IO.File.Exists(EpiaMsiFile))
                {
                    if (m_TestAutoMode)
                    {
                        TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Passed", "No epia msi file found", sInfoFileKey);
                        logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Started: No epia msi file found : m_testApp " + mTestApp, 0, 0);

                        #region    // Update build quality   "GUI Tests Started" to "GUI Tests Passed" if needed

                        string msgXx = "update build quality Before GUI Tests Started";
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgXx);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgXx + "\nWill try to reconnect the Server ...",
                                   "update build quality Deployment Started", (uint)Tfs.ReconnectDelay);
                            System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                            TFSConnected = TfsUtilities.CheckTFSConnection(ref msgXx);
                        }

                        if (TFSConnected)
                        {
                            //Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                            //   TestTools.TfsUtilities.GetProjectName(mTestApp), m_BuildNumber);
                            // m_Uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, BuildUtilities.GetProjectName(mTestApp), m_BuildNumber);
                            string quality = m_BuildSvc.GetMinimalBuildDetails(m_Uri).Quality;

                            sErrorMessage = string.Empty;
                            mTestDefFile = @"C:\EtriccTests\QA\TestDefinitions\" + TfsUtilities.GetTestDefNameFromTestApp(mTestApp);
                            logger.LogMessageToFile("  mTestDefFile :" + mTestDefFile, 0, 0);
                            string[] mTestDefinitionTypes = System.IO.File.ReadAllLines(mTestDefFile);

                            for (int i = 0; i < mTestDefinitionTypes.Length; i++)
                            {
                                Console.WriteLine(i + " testdefinition : " + mTestDefinitionTypes[i]);
                                logger.LogMessageToFile(i + " testdefinition : " + mTestDefinitionTypes[i], 0, 0);
                            }
                            logger.LogMessageToFile("GUI Tests Passed:" + mTestApp, 0, 0);

                            Console.WriteLine(" Update build quality:  quality: " + quality);
                            if (quality.Equals("GUI Tests Failed"))
                            {
                                Console.WriteLine("Quality is GUI Tests Failed : No Update Needed ");
                                logger.LogMessageToFile(m_BuildNumber + " has failed quality, no update needed :" + quality, 0, 0);
                            }
                            else
                            {
                                Console.WriteLine(" mTestDefinitionTypes[0]: " + mTestDefinitionTypes[0]);
                                if (TestListUtilities.IsAllTestDefinitionsTested(mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage) == false)
                                {
                                    logger.LogMessageToFile("NOT All Test definitions tested " + sErrorMessage, 0, 0);
                                    Console.WriteLine("NOT All Test definitions tested " + sErrorMessage);
                                }
                                else
                                {
                                    if (TestListUtilities.IsAllTestStatusPassed(mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage) == true)
                                    {
                                        Console.WriteLine("update quality to GUI Tests Passed -----  ");
                                        string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TfsUtilities.GetProjectName(mTestApp),
                                        "GUI Tests Passed", 
                                        //m_BuildSvc, sDemonstration);
                                            m_BuildSvc, "false");// only for demo

                                        Log(updateResult);
                                        if (updateResult.StartsWith("Error"))
                                        {
                                            System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                            throw new Exception(updateResult);
                                        }
                                        Thread.Sleep(1000);
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    m_State = STATE.PENDING;
                    return;
                }
                #endregion
            }
            else
            {
                System.Windows.MessageBox.Show("Unknown Application, try other application again...   " + mTestApp);
                m_State = STATE.PENDING;
                return;
            }
            #endregion

            #region    // Update build quality   "GUI Tests Started"
            string msgX = "update build quality GUI Tests Started";
            /*TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
            while (TFSConnected == false)
            {
                TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                       "update build quality GUI Tests Started", (uint)Tfs.ReconnectDelay);
                System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
            }*/
            if (TFSConnected)
            {
                /*string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TfsUtilities.GetProjectName(mTestApp),
                    "GUI Tests Started", 
                    // m_BuildSvc, sDemonstration);
                    m_BuildSvc, "false");// only for demo

                Log(updateResult);
                if (updateResult.StartsWith("Error"))
                {
                    System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                    throw new Exception(updateResult);
                }*/

                //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                if (m_TestAutoMode)
                {
                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Started", "Start Deployment", sInfoFileKey);
                    logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Started: Start Deployment : m_testApp " + mTestApp, 0, 0);

                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.Forms.MessageBox.Show("Deployment Started:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);

                }
            }
            #endregion
            Thread.Sleep(3000);
            string installPath = m_CurrentDrive;   // only Etricc 5 depende on Drive, all other apps are on C:

            // Start Deployment   .......
            #region// Start Deployment   .......

            try
            {
                sCurrentBuildInTesting = "Start Deployment: " + m_installScriptDir;
                sEpia4InstallerName = Constants.sEpia4InstallerName;  // epia.msi

                if (mTestApp.Equals(Constants.ETRICC5))
                {
                    #region // Etricc 5
                    //MessageBox.Show("Start depmoyment:" + m_BuildNumber);
                    logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
                    System.Threading.Thread.Sleep(2000);

                    //  Install setup
                    //Install new setup and recompile Worker
                    mEpiaPath = m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Etricc Server" + "\\";
                    //MessageBox.Show(" Install path :" + mEpiaPath);
                    Log(" Install path :" + mEpiaPath);
                    logger.LogMessageToFile(" Install path :" + mEpiaPath, 0, 0);

                    // Remove Old SetUP in silent mode
                    RemoveSetupSilentMode(Constants.ETRICC5);

                    //Move the current Setup files to a backup location
                    if (!CopySetup(mCurrentSetupPathEtricc5, mPreviousSetupPathEtricc5))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }

                    if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtricc5, sEtriccMsiName))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }

                    // check if current folder is empty
                    FileInfo[] FilesToCopy;
                    DirectoryInfo DirInfo = new DirectoryInfo(mCurrentSetupPathEtricc5);
                    FilesToCopy = DirInfo.GetFiles(sEtriccMsiName);
                    if (FilesToCopy.Length == 0)
                    {
                        #region    // Update build quality   "Deployment Failed"
                        msgX = "update build quality Deployment Failed";
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                  "update build quality Deployment Failed", (uint)Tfs.ReconnectDelay);
                            System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                            TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TfsUtilities.GetProjectName(mTestApp), "Deployment Failed", 
                                //m_BuildSvc, sDemonstration);
                                    m_BuildSvc, "false");// only for demo

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed", "Deployment Failed: no 'Etricc ?.msi' file found-->" + mTestApp, sInfoFileKey);
                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : mTestApp " + mTestApp, 0, 0);
                            }
                        }
                        #endregion
                        m_State = STATE.PENDING;
                        return;
                    }

                    // remove E'tricc Service in case this service is still exist
                    System.Diagnostics.Process procRemoveService = new System.Diagnostics.Process();
                    procRemoveService.EnableRaisingEvents = false;
                    procRemoveService.StartInfo.FileName = "sc";
                    procRemoveService.StartInfo.Arguments = "delete " + '"' + "E'tricc Service" + '"';
                    procRemoveService.StartInfo.WorkingDirectory = m_TestWorkingDirectory;
                    procRemoveService.Start();
                    procRemoveService.WaitForExit();

                    // Install Current Setup
                    logger.LogMessageToFile(" *** start EtriccCore installing :" + m_BuildNumber, 0, 0);
                    mInstallEtriccLauncher = m_TestConfigSettings.Element.InstallEtriccLauncher;
                    mInstallOldEtricc5Service = m_TestConfigSettings.Element.InstallOldEtriccService;
                    //if (InstallEtricc5Setup2(mCurrentSetupPathEtricc5, ""))
                    if (ProjAppInstall.InstallEtriccCoreSetup(mCurrentSetupPathEtricc5, sEtriccMsiName, ref sErrorMessage, logger, mInstallEtriccLauncher, mInstallOldEtricc5Service))
                    {
                        //save the name and the path in the registry to remove the setup
                        WriteInstallationToReg(Constants.ETRICC5, mCurrentSetupPathEtricc5, sEtriccMsiName);
                        logger.LogMessageToFile(" **************** ( END Etricc ?.msi Deployment )************************** ", 0, 0);
                        logger.LogMessageToFile(" EtriccCore installed :" + m_BuildNumber, 0, 0);

                        string cscOut = RecompileTestRuns();
                        logger.LogMessageToFile("------ReCOMPILE TestRUNS output -------:" + cscOut, 0, 0);
                    }

                    #endregion
                }
                else if (mTestApp.Equals(TestApp.EPIA4))
                {
                    #region // Epia4
                    logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
                    System.Threading.Thread.Sleep(2000);
                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server", true);
                    }

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell", true);
                    }

                    //  Install Epia setup
                    //Install new setup
                    mEpiaPath = m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Epia Server";
                    // Clean up epia server folder
                    if (System.IO.Directory.Exists(mEpiaPath))
                    {
                        logger.LogMessageToFile(" Clean up Server folder :", 0, 0);
                        if (System.IO.Directory.Exists(mEpiaPath + "\\Data"))
                            System.IO.Directory.Delete(mEpiaPath + "\\Data", true);
                        if (System.IO.Directory.Exists(mEpiaPath + "\\Log"))
                            System.IO.Directory.Delete(mEpiaPath + "\\Log", true);
                    }

                    //MessageBox.Show(" Install path :" + mEpiaPath);
                    //MessageBox.Show(" mCurrentSetupPath :" + mCurrentSetupPathEpia);
                    Log(" Install path :" + mEpiaPath);
                    logger.LogMessageToFile(" Install path :" + mEpiaPath, 0, 0);

                    // Remove Old Epia SetUP in silent mode 
                    RemoveSetupSilentMode(EgeminApplication.EPIA);

                    //Move the current Setup files to a backup location
                    if (!CopySetup(mCurrentSetupPathEpia, mPreviousSetupPathEpia))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }

                    //MessageBox.Show(" m_installScriptDir :" + m_installScriptDir);
                    //if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, "*" + sEpia4InstallerName))
                    if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, "Epia*.msi"))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }

                    // check if current folder is empty
                    FileInfo[] FilesToCopy;
                    DirectoryInfo DirInfo = new DirectoryInfo(mCurrentSetupPathEpia);
                    FilesToCopy = DirInfo.GetFiles("*" + sEpia4InstallerName);
                    if (FilesToCopy.Length == 0)
                    {
                        #region    // Update build quality   "GUI Tests Started" -->  "GUI Tests Failed"
                        msgX = "update epia build quality Deployment Failed";
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                    "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay);
                            System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                            TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            if (m_TestAutoMode)
                            {
                                string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri,
                                    TfsUtilities.GetProjectName(mTestApp), "GUI Tests Failed", 
                                    //m_BuildSvc, sDemonstration);
                                        m_BuildSvc, "false");// only for demo

                                Log(updateResult);
                                if (updateResult.StartsWith("Error"))
                                {
                                    System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                    throw new Exception(updateResult);
                                }

                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed", "Deployment Failed", sInfoFileKey);
                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed: Deployment Failed : m_testApp " + mTestApp, 0, 0);

                                if (sMsgDebug.StartsWith("true"))
                                    System.Windows.Forms.MessageBox.Show("Deployment Started:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);

                            }
                        }
                        #endregion
                        m_State = STATE.PENDING;
                        return;
                    }

                    //-----------------------------------
                    // Install Current Epia Setup
                    //-----------------------------------
                    logger.LogMessageToFile(" *** start Epia installing :" + m_BuildNumber, 0, 0);
                    if (ProjAppInstall.InstallApplication(mCurrentSetupPathEpia, EgeminApplication.EPIA, EgeminApplication.SetupType.Default, ref sErrorMessage, logger, Convert.ToBoolean(sDemonstration)) )
                    {
                        //save the name and the path in the registry to remove the setup
                        WriteInstallationToReg(TestApp.EPIA4, mCurrentSetupPathEpia, "Epia.msi");

                        logger.LogMessageToFile(" **************** ( END Epia Epia.msi Deployment )************************** ", 0, 0);
                        logger.LogMessageToFile(" Epia installed :" + m_BuildNumber, 0, 0);
                    }
                    else
                    {
                        logger.LogMessageToFile(" *** Epia installing  failed:" + sErrorMessage, 0, 0);
                        throw new Exception(" *** Epia installing  failed:" + sErrorMessage);
                    }
                    #endregion
                }
                else if (mTestApp.Equals(Constants.ETRICCUI))
                {
                    #region // EtriccUI
                    logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
                    System.Threading.Thread.Sleep(2000);

                    // Remove Old SetUP first
                    RemoveSetupSilentMode(Constants.ETRICCUI);   // because the register is EtriccInstallation
                    RemoveSetupSilentMode(Constants.ETRICC5);
                    RemoveSetupSilentMode(EgeminApplication.EPIA);

                    //Move the current Etricc shell, hosttest and playback msi files to a backup location
                    if (!CopySetup(mCurrentSetupPathEtriccShell, mPreviousSetupPathEtriccShell))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }

                    //string s = " m_installScriptDir :" + m_installScriptDir + "\n\n mCurrentSetupPathEtricc :" + mCurrentSetupPathEtricc;
                    //TestTools.MessageBoxEx.Show(s, 30);
                    // copy Etricc Shell HostTest PlayBack msi files to current folder
                    if (System.IO.File.Exists(System.IO.Path.Combine(m_installScriptDir, "Etricc Shell.msi")))
                    {
                        if (!CopyEtriccSetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccShell, "*Shell.msi"))
                        {
                            m_State = STATE.PENDING;
                            return;
                        }

                        if (!CopyEtriccSetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccShell, "*HostTest.msi"))
                        {
                            m_State = STATE.PENDING;
                            return;
                        }

                        if (!CopyEtriccSetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccShell, "*Playback.msi"))
                        {
                            m_State = STATE.PENDING;
                            return;
                        }
                    }
                    else
                    {
                        #region    // Update build quality   "GUI Tests Started" -->  "GUI Tests Failed"
                        msgX = "update epia build quality Deployment Failed";
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                    "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay);
                            System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                            TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            if (m_TestAutoMode)
                            {
                                string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri,
                                    TfsUtilities.GetProjectName(mTestApp), "GUI Tests Failed", 
                                    //m_BuildSvc, sDemonstration);
                                        m_BuildSvc, "false");// only for demo

                                Log(updateResult);
                                if (updateResult.StartsWith("Error"))
                                {
                                    System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                    throw new Exception(updateResult);
                                }

                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed", "Deployment Failed", sInfoFileKey);

                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed: Deployment Failed : m_testApp " + mTestApp, 0, 0);

                                if (sMsgDebug.StartsWith("true"))
                                    System.Windows.Forms.MessageBox.Show("Deployment Started:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);

                            }
                        }
                        #endregion
                        m_State = STATE.PENDING;
                        return;
                    }

                    //---------------------------------------------------------------
                    //Move the current Epia msi files to a backup location
                    if (!CopySetup(mCurrentSetupPathEpia, mPreviousSetupPathEpia))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }

                    if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, Constants.EPIA_MSI))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }

                    //---------------------------------------------------------------
                    //Move the current Etricc5 msi files to a backup location
                    if (!CopySetup(mCurrentSetupPathEtricc5, mPreviousSetupPathEtricc5))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }

                    bool msiExist = true;
                    sEtriccMsiName = "Etricc.msi";
                    string EtriccCoreMsiFile = System.IO.Path.Combine(m_installScriptDir, sEtriccMsiName);
                    if (!System.IO.File.Exists(EtriccCoreMsiFile))
                    {
                        sEtriccMsiName = "Etricc 6.msi";
                        EtriccCoreMsiFile = System.IO.Path.Combine(m_installScriptDir, sEtriccMsiName);
                        if (!System.IO.File.Exists(EtriccCoreMsiFile))
                        {
                            sEtriccMsiName = "Etricc 5.msi";
                            EtriccCoreMsiFile = System.IO.Path.Combine(m_installScriptDir, sEtriccMsiName);
                            if (!System.IO.File.Exists(EtriccCoreMsiFile))
                            {
                                msiExist = false;
                            }
                        }
                    }

                    if (msiExist == true)
                    {
                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtricc5, sEtriccMsiName))
                        {
                            m_State = STATE.PENDING;
                            return;
                        }
                    }
                    else
                    {
                        #region    // Update build quality   "GUI Tests Started" -->  "GUI Tests Failed"
                        msgX = "update epia build quality Deployment Failed";
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                    "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay);
                            System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                            TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            if (m_TestAutoMode)
                            {
                                string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri,
                                    TfsUtilities.GetProjectName(mTestApp), "GUI Tests Failed", 
                                    //m_BuildSvc, sDemonstration);
                                        m_BuildSvc, "false");// only for demo

                                Log(updateResult);
                                if (updateResult.StartsWith("Error"))
                                {
                                    System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                    throw new Exception(updateResult);
                                }

                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed", "Deployment Failed", sInfoFileKey);

                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed: Deployment Failed : m_testApp " + mTestApp, 0, 0);

                                if (sMsgDebug.StartsWith("true"))
                                    System.Windows.Forms.MessageBox.Show("Deployment Started:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);

                            }
                        }
                        #endregion
                        m_State = STATE.PENDING;
                        return;
                    }

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server", true);
                    }

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell", true);
                    }

                    //-----------------------------------------------------------------
                    // remove E'tricc Service in case this service is still exist
                    System.Diagnostics.Process procRemoveService = new System.Diagnostics.Process();
                    procRemoveService.EnableRaisingEvents = false;
                    procRemoveService.StartInfo.FileName = "sc";
                    procRemoveService.StartInfo.Arguments = "delete " + '"' + "E'tricc Service" + '"';
                    procRemoveService.StartInfo.WorkingDirectory = m_TestWorkingDirectory;
                    procRemoveService.Start();
                    procRemoveService.WaitForExit();
                    //-----------------------------------
                    // Install Current Epia Setup 
                    //-----------------------------------
                    logger.LogMessageToFile(" *** start Epia installing :" + m_BuildNumber, 0, 0);
                    if (ProjAppInstall.InstallApplication(mCurrentSetupPathEpia, EgeminApplication.EPIA, EgeminApplication.SetupType.Default, ref sErrorMessage, logger))
                    {
                        //save the name and the path in the registry to remove the setup
                        WriteInstallationToReg(TestApp.EPIA4, mCurrentSetupPathEpia, "Epia.msi");
                        logger.LogMessageToFile(" **************** ( END Epia Epia.msi Deployment )************************** ", 0, 0);
                        logger.LogMessageToFile(" Epia installed :" + m_BuildNumber, 0, 0);
                    }
                    else
                    {
                        logger.LogMessageToFile(" *** Epia installing  failed:" + sErrorMessage, 0, 0);
                        throw new Exception(" *** Epia installing  failed:" + sErrorMessage);
                    }

                    //----------------------------------
                    // Install Current Etricc Core Setup 
                    //----------------------------------
                    logger.LogMessageToFile(" *** start EtriccCore installing :" + m_BuildNumber, 0, 0);
                    mInstallEtriccLauncher = m_TestConfigSettings.Element.InstallEtriccLauncher;
                    mInstallOldEtricc5Service = m_TestConfigSettings.Element.InstallOldEtriccService;
                    if (ProjAppInstall.InstallEtriccCoreSetup(mCurrentSetupPathEtricc5, sEtriccMsiName, ref sErrorMessage, logger, mInstallEtriccLauncher, mInstallOldEtricc5Service))
                    {
                        //save the name and the path in the registry to remove the setup
                        WriteInstallationToReg(Constants.ETRICC5, mCurrentSetupPathEtricc5, sEtriccMsiName);
                        logger.LogMessageToFile(" **************** ( END Etricc ?.msi Deployment )************************** ", 0, 0);
                        logger.LogMessageToFile(" EtriccCore installed :" + m_BuildNumber, 0, 0);
                    }
                    else
                    {
                        logger.LogMessageToFile(" *** Etricc Core installing  failed:" + sErrorMessage, 0, 0);
                        throw new Exception(" *** Etricc Core installing  failed:" + sErrorMessage);
                    }
                    //-----------------------------------
                    // Install Current Etricc Shell Setup 
                    //-----------------------------------
                    logger.LogMessageToFile(" *** start Etricc shell installing :" + m_BuildNumber, 0, 0);
                    if (ProjAppInstall.InstallApplication(mCurrentSetupPathEtriccShell, EgeminApplication.ETRICC_SHELL, 
                        EgeminApplication.SetupType.Default, ref sErrorMessage, logger))
                    {
                        //save the name and the path in the registry to remove the setup
                        WriteInstallationToReg(Constants.ETRICCUI, mCurrentSetupPathEtriccShell, "Etricc Shell.msi");

                        logger.LogMessageToFile(" **************** ( END Etricc Shell.msi Deployment )************************** ", 0, 0);
                        logger.LogMessageToFile(" Etricc Shell installed :" + m_BuildNumber, 0, 0);
                    }
                    else
                    {
                        logger.LogMessageToFile(" *** Etricc shell installing  failed:" + sErrorMessage, 0, 0);
                        throw new Exception(" *** Etricc shell installing  failed:" + sErrorMessage);
                    }

                    //--------------------------------------------------------------------------------
                    // Check Project File     Current default is DEMO.xml 
                    //--------------------------------------------------------------------------------
                    sEtricc5InstallationFolder = OSVersionInfoClass.ProgramFilesx86() + "\\Egemin\\Etricc Server";
                    Log(" Install path :" + mEpiaPath);
                    logger.LogMessageToFile("--- Install path :" + mEpiaPath, 0, 0);
                    //logger.LogMessageToFile("--- m_Settings.SelectedProjectFile :" + m_Settings.SelectedProjectFile, 0, 0);
                    logger.LogMessageToFile("--- m_Settings.SelectedProjectFile :" + m_TestConfigSettings.Element.SelectedProjectFile, 0, 0);
                    #region //   Check Project file
                    if (mTestApp.Equals(Constants.ETRICCUI))
                    {
                        //string projectXml = m_Settings.SelectedProjectFile;
                        string projectXml = m_TestConfigSettings.Element.SelectedProjectFile;
                        //string xmlPath = mEpiaPath + "\\Data\\Etricc\\Demo.xml";
                        string xmlPath = sEtricc5InstallationFolder + "\\Data\\Etricc\\Demo.xml";
                        logger.LogMessageToFile("---- Check Project file:" + xmlPath, 0, 0);
                        // Etricc deployment take long time than Epia, It copy demo.xml, check if copied
                        while (!System.IO.File.Exists(xmlPath))
                        {
                            logger.LogMessageToFile("---- Not copied yet, wait:" + xmlPath, 0, 0);
                            Thread.Sleep(3000);
                        }

                        logger.LogMessageToFile("--- projectXml :" + projectXml, 0, 0);

                        if (projectXml.Equals("Demo.xml"))
                            sProjectFile = "Demo.xml";
                        else if (projectXml.IndexOf("Eurobaltic") >= 0)
                        {
                            string zipFile = Path.Combine(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "Eurobaltic.zip");
                            FastZip fz = new FastZip();
                            fz.ExtractZip(zipFile, ConstCommon.ETRICC_TESTS_DIRECTORY, "");
                            Thread.Sleep(10000);
                            string Xml = Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "Eurobaltic.xml");
                            string xmlPathBackup = mEpiaPath + "\\Data\\Etricc\\DemoBackup.xml";
                            Thread.Sleep(2000);
                            System.IO.File.Delete(xmlPath);
                            Thread.Sleep(2000);
                            System.IO.File.Copy(Xml, xmlPath, true);
                            sProjectFile = System.IO.Path.GetFileName(projectXml);
                        }
                        else if (projectXml.IndexOf("TestProject") >= 0)
                        {
                            string zipFile = projectXml;
                            FastZip fz = new FastZip();
                            fz.ExtractZip(zipFile, m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "");
                            Thread.Sleep(30000);
                            string Xml = Path.Combine(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "TestProject.xml");
                            string xmlPathBackup = mEpiaPath + "\\Data\\Etricc\\DemoBackup.xml";
                            System.IO.File.Delete(xmlPath);
                            System.IO.File.Copy(Xml, xmlPath, true);
                            sProjectFile = System.IO.Path.GetFileName(projectXml);
                        }
                        else
                        {
                            // do later

                        }

                        if (m_TestConfigSettings.Element.FunctionalTesting)
                        {
                            string cscOut = RecompileTestRuns();
                            logger.LogMessageToFile("------ReCOMPILE TestRUNS output -------:" + cscOut, 0, 0);
                            System.Windows.Forms.MessageBox.Show( cscOut, "------ReCOMPILE TestRUNS output -------:" );
                            return;
                        }
                    }
                    #endregion

                    #endregion
                }
                else if (mTestApp.Equals(Constants.ETRICCSTATISTICS))
                {
                    #region // EtriccStatistics
                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.MessageBox.Show("Start depmoyment:" + m_BuildNumber);
                    logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
                    System.Threading.Thread.Sleep(2000);

                    // Remove Old SetUP first
                    RemoveSetupSilentMode(Constants.ETRICCSTATISTICS_UI);
                    RemoveSetupSilentMode(EgeminApplication.EPIA);
                    RemoveSetupSilentMode(Constants.ETRICCSTATISTICS_PARSERCONFIGURATOR);
                    RemoveSetupSilentMode(Constants.ETRICCSTATISTICS_PARSER_SETUP);

                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.MessageBox.Show("Check removed applications:");

                    //(1) Move the current Etricc Statistics UI Setup files to a backup location
                    #region //---------------------------------------------------------------
                    if (!CopySetup(mCurrentSetupPathEtriccStatisticsUI, mPreviousSetupPathEtriccStatisticsUI))
                        return;
                    else
                        logger.LogMessageToFile("CopySetup(mCurrentSetupPathEtriccStatisticsUI, mPreviousSetupPathEtriccStatisticsUI):OK", 0, 0);

                    if (sMsgDebug.StartsWith("true"))
                        MessageBox.Show(" m_installScriptDir :" + m_installScriptDir);

                    string EtriccStatisticsUIMsi = "Etricc.Statistics.UI.msi";
                    string EtriccStatisticsUIMsiFile = System.IO.Path.Combine(m_installScriptDir, EtriccStatisticsUIMsi);
                    if (System.IO.File.Exists(EtriccStatisticsUIMsiFile))
                    {
                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccStatisticsUI, EtriccStatisticsUIMsi))
                        {
                            m_State = STATE.PENDING;
                            return;
                        }
                        else
                            logger.LogMessageToFile("CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccStatisticsUI, EtriccStatisticsUIMsi):OK", 0, 0);
                    }
                    else
                    {
                        #region    // Update build quality   "GUI Tests Started" -->  "GUI Tests Failed"
                        msgX = "update statistics build quality Deployment Failed";
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                    "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay);
                            System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                            TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            if (m_TestAutoMode)
                            {
                                string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri,
                                    TfsUtilities.GetProjectName(mTestApp), "GUI Tests Failed", 
                                    //m_BuildSvc, sDemonstration);
                                        m_BuildSvc, "false");// only for demo

                                Log(updateResult);
                                if (updateResult.StartsWith("Error"))
                                {
                                    System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                    throw new Exception(updateResult);
                                }

                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed", "EtriccStatisticsUIMsi not exist", sInfoFileKey);

                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed: EtriccStatisticsUIMsi not exist:: m_testApp " + mTestApp, 0, 0);

                                if (sMsgDebug.StartsWith("true"))
                                    System.Windows.Forms.MessageBox.Show("Deployment Started:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);

                            }
                        }
                        #endregion
                        m_State = STATE.PENDING;
                        return;
                    }
                    #endregion

                    //(2) Move the current Etricc Statistics Parser Setup files to a backup location
                    #region //---------------------------------------------------------------
                    if (!CopySetup(mCurrentSetupPathEtriccStatisticsParser, mPreviousSetupPathEtriccStatisticsParser))
                        return;
                    else
                        logger.LogMessageToFile("CopySetup(mCurrentSetupPathEtriccStatisticsParser, mPreviousSetupPathEtriccStatisticsParser):OK", 0, 0);

                    if (sMsgDebug.StartsWith("true"))
                        MessageBox.Show(" m_installScriptDir :" + m_installScriptDir);

                    string EtriccStatisticsParserMsi = "Etricc.Statistics.Parser.msi";
                    string EtriccStatisticsParserMsiFile = System.IO.Path.Combine(m_installScriptDir, EtriccStatisticsParserMsi);
                    if (System.IO.File.Exists(EtriccStatisticsParserMsiFile))
                    {
                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccStatisticsParser, EtriccStatisticsParserMsi))
                        {
                            m_State = STATE.PENDING;
                            return;
                        }
                        else
                            logger.LogMessageToFile("CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccStatisticsParser, EtriccStatisticsParserMsi):OK", 0, 0);
                    }
                    else
                    {
                        #region    // Update build quality   "GUI Tests Started" -->  "GUI Tests Failed"
                        msgX = "update statistics build quality Deployment Failed";
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                    "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay);
                            System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                            TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            if (m_TestAutoMode)
                            {
                                string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri,
                                    TfsUtilities.GetProjectName(mTestApp), "GUI Tests Failed", 
                                    //m_BuildSvc, sDemonstration);
                                        m_BuildSvc, "false");// only for demo

                                Log(updateResult);
                                if (updateResult.StartsWith("Error"))
                                {
                                    System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                    throw new Exception(updateResult);
                                }

                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed", "EtriccStatisticsParserMsi not exist", sInfoFileKey);

                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed: EtriccStatisticsParserMsi not exist:: m_testApp " + mTestApp, 0, 0);

                                if (sMsgDebug.StartsWith("true"))
                                    System.Windows.Forms.MessageBox.Show("Deployment Started:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);
                            }
                        }
                        #endregion
                        m_State = STATE.PENDING;
                        return;
                    }
                    #endregion

                    //(3) Move the current Etricc Statistics ParserConfigurator Setup files to a backup location
                    #region //---------------------------------------------------------------
                    if (!CopySetup(mCurrentSetupPathEtriccStatisticsParserConfigurator, mPreviousSetupPathEtriccStatisticsParserConfigurator))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }
                    else
                        logger.LogMessageToFile("CopySetup(mCurrentSetupPathEtriccStatisticsParserConfigurator, mPreviousSetupPathEtriccStatisticsParserConfigurator):OK", 0, 0);

                    if (sMsgDebug.StartsWith("true"))
                        MessageBox.Show(" m_installScriptDir :" + m_installScriptDir);

                    string EtriccStatisticsParserConfiguratorMsi = "Etricc.Statistics.ParserConfigurator.msi";
                    string EtriccStatisticsParserConfiguratorMsiFile = System.IO.Path.Combine(m_installScriptDir, EtriccStatisticsParserConfiguratorMsi);
                    if (System.IO.File.Exists(EtriccStatisticsParserConfiguratorMsiFile))
                    {
                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccStatisticsParserConfigurator, EtriccStatisticsParserConfiguratorMsi))
                        {
                            m_State = STATE.PENDING;
                            return;
                        }
                        else
                            logger.LogMessageToFile("CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccStatisticsParserConfigurator, EtriccStatisticsParserConfiguratorMsi):OK", 0, 0);


                    }
                    else
                    {
                        #region    // Update build quality   "GUI Tests Started" -->  "GUI Tests Failed"
                        msgX = "update statistics build quality Deployment Failed";
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                    "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay);
                            System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                            TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            if (m_TestAutoMode)
                            {
                                string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri,
                                    TfsUtilities.GetProjectName(mTestApp), "GUI Tests Failed", 
                                    //m_BuildSvc, sDemonstration);
                                        m_BuildSvc, "false");// only for demo

                                Log(updateResult);
                                if (updateResult.StartsWith("Error"))
                                {
                                    System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                    throw new Exception(updateResult);
                                }

                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed", "EtriccStatisticsParserConfiguratorMsi not exist", sInfoFileKey);

                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed: EtriccStatisticsParserConfiguratorMsi not exist:: m_testApp " + mTestApp, 0, 0);

                                if (sMsgDebug.StartsWith("true"))
                                    System.Windows.Forms.MessageBox.Show("Deployment Started:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);

                            }
                        }
                        #endregion
                        m_State = STATE.PENDING;
                        return;
                    }
                    #endregion

                    //(4) Move the current Epia Setup files to a backup location
                    #region //---------------------------------------------------------------
                    if (!CopySetup(mCurrentSetupPathEpia, mPreviousSetupPathEpia))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }
                    else
                        logger.LogMessageToFile("CopySetup(mCurrentSetupPathEpia, mPreviousSetupPathEpia):OK", 0, 0);

                    //string EtriccStatisticsParserConfiguratorMsi = "Etricc.Statistics.ParserConfigurator.msi";
                    string Epia4MsiFile = System.IO.Path.Combine(m_installScriptDir, sEpia4InstallerName);
                    if (System.IO.File.Exists(Epia4MsiFile))
                    {
                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, "*" + sEpia4InstallerName))
                        {
                            m_State = STATE.PENDING;
                            return;
                        }
                        else
                            logger.LogMessageToFile("CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, ' * '" + sEpia4InstallerName + ":OK", 0, 0);

                    }
                    else
                    {
                        #region    // Update build quality   "GUI Tests Started" -->  "GUI Tests Failed"
                        msgX = "update statistics build quality Deployment Failed";
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                    "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay);
                            System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                            TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            if (m_TestAutoMode)
                            {
                                string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri,
                                    TfsUtilities.GetProjectName(mTestApp), "GUI Tests Failed", 
                                    //m_BuildSvc, sDemonstration);
                                        m_BuildSvc, "false");// only for demo

                                Log(updateResult);
                                if (updateResult.StartsWith("Error"))
                                {
                                    System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                    throw new Exception(updateResult);
                                }

                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed", sEpia4InstallerName + " not exist", sInfoFileKey);

                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed: " + sEpia4InstallerName + " not exist:: m_testApp " + mTestApp, 0, 0);

                                if (sMsgDebug.StartsWith("true"))
                                    System.Windows.Forms.MessageBox.Show("Deployment Started:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);

                            }
                        }
                        #endregion
                        m_State = STATE.PENDING;
                        return;
                    }
                    #endregion

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server", true);
                    }

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Etricc Statistics Parser"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Etricc Statistics Parser", true);
                    }

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Etricc Statistics ParserConfigurator"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Etricc Statistics ParserConfigurator", true);
                    }

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell", true);
                    }

                    //-----------------------------------------------------------------
                    // remove Egemin E'tricc Statistics Parser Service in case this service is still exist
                    System.Diagnostics.Process procRemoveService = new System.Diagnostics.Process();
                    procRemoveService.EnableRaisingEvents = false;
                    procRemoveService.StartInfo.FileName = "sc";
                    procRemoveService.StartInfo.Arguments = "delete " + '"' + "Egemin E'tricc Statistics Parser" + '"';
                    procRemoveService.StartInfo.WorkingDirectory = m_TestWorkingDirectory;
                    procRemoveService.Start();
                    procRemoveService.WaitForExit();

                    //-----------------------------------
                    // Install Current EtriccStatisticsParser Setup 
                    //-----------------------------------
                    logger.LogMessageToFile(" *** start EtriccStatisticsParser installing :" + m_BuildNumber, 0, 0);
                    if (ProjAppInstall.InstallApplication(mCurrentSetupPathEtriccStatisticsParser, EgeminApplication.ETRICC_STATISTICS_PARSER, 
                        EgeminApplication.SetupType.Default, ref sErrorMessage, logger))
                    {
                        //save the name and the path in the registry to remove the setup
                        WriteInstallationToReg(Constants.ETRICCSTATISTICS_PARSER_SETUP, mCurrentSetupPathEtriccStatisticsParser, EgeminApplication.ETRICC_STATISTICS_PARSER + ".msi");

                        logger.LogMessageToFile(" **************** ( END EtriccStatisticsParser " + EgeminApplication.ETRICC_STATISTICS_PARSER + ".msi"
                            + "Deployment )************************** ", 0, 0);
                        logger.LogMessageToFile(" EtriccStatisticsParser installed :" + m_BuildNumber, 0, 0);
                        Log(" EtriccStatisticsParser installed :" + m_BuildNumber);
                    }
                    else
                    {
                        logger.LogMessageToFile(" *** EtriccStatisticsParser installing  failed:" + sErrorMessage, 0, 0);
                        throw new Exception(" *** EtriccStatisticsParser installing  failed:" + sErrorMessage);
                    }

                    //-----------------------------------
                    // Install Current EtriccStatisticsParserConfigurator Setup 
                    //-----------------------------------
                    logger.LogMessageToFile(" *** start EtriccStatisticsParserConfigurator installing :" + m_BuildNumber, 0, 0);
                    if (ProjAppInstall.InstallApplication(mCurrentSetupPathEtriccStatisticsParserConfigurator, EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR,
                        EgeminApplication.SetupType.Default, ref sErrorMessage, logger))
                    {
                        //save the name and the path in the registry to remove the setup
                        WriteInstallationToReg(Constants.ETRICCSTATISTICS_PARSERCONFIGURATOR, mCurrentSetupPathEtriccStatisticsParserConfigurator,
                            EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR + ".msi");

                        logger.LogMessageToFile(" **************** ( END EtriccStatisticsParserConfigurator " + EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR + ".msi"
                            + "Deployment )************************** ", 0, 0);
                        logger.LogMessageToFile(" EtriccStatisticsParserConfigurator installed :" + m_BuildNumber, 0, 0);
                        Log(" EtriccStatisticsParserConfigurator installed :" + m_BuildNumber);
                    }
                    else
                    {
                        logger.LogMessageToFile(" *** EtriccStatisticsParserConfigurator installing  failed:" + sErrorMessage, 0, 0);
                        throw new Exception(" *** EtriccStatisticsParserConfigurator installing  failed:" + sErrorMessage);
                    }

                    //-----------------------------------
                    // Install Current Epia Setup 
                    //-----------------------------------
                    logger.LogMessageToFile(" *** start Epia installing :" + m_BuildNumber, 0, 0);
                    if (ProjAppInstall.InstallApplication(mCurrentSetupPathEpia, EgeminApplication.EPIA, EgeminApplication.SetupType.Default, ref sErrorMessage, logger))
                    {
                        //save the name and the path in the registry to remove the setup
                        WriteInstallationToReg(TestApp.EPIA4, mCurrentSetupPathEpia, "Epia.msi");

                        logger.LogMessageToFile(" **************** ( END Epia Epia.msi Deployment )************************** ", 0, 0);
                        logger.LogMessageToFile(" Epia installed :" + m_BuildNumber, 0, 0);
                        Log(" Epia installed :" + m_BuildNumber);
                    }
                    else
                    {
                        logger.LogMessageToFile(" *** Epia installing  failed:" + sErrorMessage, 0, 0);
                        throw new Exception(" *** Epia installing  failed:" + sErrorMessage);
                    }

                    //-----------------------------------
                    // Install Current EtriccStatisticsUI Setup 
                    //-----------------------------------
                    logger.LogMessageToFile(" *** start EtriccStatisticsUI installing :" + m_BuildNumber, 0, 0);
                    if (ProjAppInstall.InstallApplication(mCurrentSetupPathEtriccStatisticsUI, EgeminApplication.ETRICC_STATISTICS_UI,
                        EgeminApplication.SetupType.Default, ref sErrorMessage, logger))
                    {
                        //save the name and the path in the registry to remove the setup
                        WriteInstallationToReg(Constants.ETRICCSTATISTICS_UI, mCurrentSetupPathEtriccStatisticsUI, EgeminApplication.ETRICC_STATISTICS_UI + ".msi");

                        logger.LogMessageToFile(" **************** ( END EtriccStatisticsUI " + EgeminApplication.ETRICC_STATISTICS_UI + ".msi"
                            + "Deployment )************************** ", 0, 0);
                        logger.LogMessageToFile(" EtriccStatisticsUI installed :" + m_BuildNumber, 0, 0);
                        Log(" EtriccStatisticsUI installed :" + m_BuildNumber);
                    }
                    else
                    {
                        logger.LogMessageToFile(" *** EtriccStatisticsUI installing  failed:" + sErrorMessage, 0, 0);
                        throw new Exception(" *** EtriccStatisticsUI installing  failed:" + sErrorMessage);
                    }
                    #endregion
                }
                else if (mTestApp.Equals(TestApp.EPIANET45))
                {
                    #region // Epia4
                    logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
                    // Remove Old Epia SetUP in silent mode 
                    RemoveSetupSilentMode(EgeminApplication.EPIA);

                    System.Threading.Thread.Sleep(2000);
                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server", true);
                    }

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell", true);
                    }

                    //  Install Epia45 setup
                    //Install new setup
                    mEpiaPath = m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Dematic\Epia Server";
                    // Clean up epia server folder
                    if (System.IO.Directory.Exists(mEpiaPath))
                    {
                        logger.LogMessageToFile(" Clean up Server folder :", 0, 0);
                        if (System.IO.Directory.Exists(mEpiaPath + "\\Data"))
                            System.IO.Directory.Delete(mEpiaPath + "\\Data", true);
                        if (System.IO.Directory.Exists(mEpiaPath + "\\Log"))
                            System.IO.Directory.Delete(mEpiaPath + "\\Log", true);
                    }

                    //MessageBox.Show(" Install path :" + mEpiaPath);
                    //MessageBox.Show(" mCurrentSetupPath :" + mCurrentSetupPathEpia);
                    Log(" Install path :" + mEpiaPath);
                    logger.LogMessageToFile(" Install path :" + mEpiaPath, 0, 0);

                  
                    //Move the current Setup files to a backup location
                    if (!CopySetup(mCurrentSetupPathEpia, mPreviousSetupPathEpia))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }

                    //MessageBox.Show(" m_installScriptDir :" + m_installScriptDir);
                    //if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, "*" + sEpia4InstallerName))
                    logger.LogMessageToFile("CopySetupFilesWithWildcards --->" + " m_installScriptDir " + m_installScriptDir, 0, 0);

                    if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, "Epia*.msi"))
                    {
                        m_State = STATE.PENDING;
                        return;
                    }

                    // check if current folder is empty
                    FileInfo[] FilesToCopy;
                    DirectoryInfo DirInfo = new DirectoryInfo(mCurrentSetupPathEpia);
                    FilesToCopy = DirInfo.GetFiles("*" + sEpia4InstallerName);
                    if (FilesToCopy.Length == 0)
                    {
                        #region    // Update build quality   "GUI Tests Started" -->  "GUI Tests Failed"
                        msgX = "update epia build quality Deployment Failed";
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                    "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay);
                            System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                            TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            if (m_TestAutoMode)
                            {
                                string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri,
                                    TfsUtilities.GetProjectName(mTestApp), "GUI Tests Failed", 
                                    //m_BuildSvc, sDemonstration);
                                        m_BuildSvc, "false");// only for demo

                                Log(updateResult);
                                if (updateResult.StartsWith("Error"))
                                {
                                    System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                    throw new Exception(updateResult);
                                }

                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed", "Deployment Failed", sInfoFileKey);
                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed: Deployment Failed : m_testApp " + mTestApp, 0, 0);

                                if (sMsgDebug.StartsWith("true"))
                                    System.Windows.Forms.MessageBox.Show("Deployment Started:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);

                            }
                        }
                        #endregion
                        m_State = STATE.PENDING;
                        return;
                    }

                    //-----------------------------------
                    // Install Current Epia45 Setup
                    //-----------------------------------
                    logger.LogMessageToFile(" *** start Epia45 installing :" + m_BuildNumber, 0, 0);
                    if (ProjAppInstall.InstallApplicationNet45(mCurrentSetupPathEpia, EgeminApplication.EPIA, EgeminApplication.SetupType.Default, 
                        ref sErrorMessage, logger))
                    {
                        //save the name and the path in the registry to remove the setup
                        WriteInstallationToReg(TestApp.EPIA4, mCurrentSetupPathEpia, "Epia.msi");

                        logger.LogMessageToFile(" **************** ( END Epia45 Epia.msi Deployment )************************** ", 0, 0);
                        logger.LogMessageToFile(" Epia installed :" + m_BuildNumber, 0, 0);
                    }
                    else
                    {
                        logger.LogMessageToFile(" *** Epia installing  failed:" + sErrorMessage, 0, 0);
                        throw new Exception(" *** Epia installing  failed:" + sErrorMessage);
                    }
                    #endregion
                }
                else
                {
                    System.Windows.MessageBox.Show("Unknown Application, try other application again..." + mTestApp);
                    m_State = STATE.PENDING;
                    return;
                }

                sCurrentBuildInTesting = "Complete deployment: " + m_installScriptDir;
                System.Configuration.Configuration config =
                    ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                // Add an Application Setting.
                config.AppSettings.Settings.Add("LastDeploymentedBuild", m_installScriptDir);
                // Save the changes in App.config file.
                config.Save(ConfigurationSaveMode.Modified);
                // Force a reload of a changed section.
                System.Configuration.ConfigurationManager.RefreshSection("appSettings");

                #region // update test info file to "Deployment Completed"
                if (m_TestAutoMode)
                {
                    // some time this build is already deleted, then do nothing, check base folder exist X:\\CI or X:\\Nightly
                    string baseFolder = m_ValidatedTestBuildObject.xMapString + "\\" + mTestDef;
                    try
                    {
                        TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Completed", "Deployment Completed", sInfoFileKey);
                        logger.LogMessageToFile(m_ValidatedBuildDropFolder + "---   Deployment Completed : testApp " + mTestApp, 0, 0);
                    }
                    catch (System.IO.FileNotFoundException fileException)
                    {
                        if (Directory.Exists(baseFolder))// this mean no network problem, this build is deleted, do nothing
                        {
                            logger.LogMessageToFile(baseFolder + "--- folder is exist, " + mTestApp, 0, 0);
                            logger.LogMessageToFile(m_ValidatedBuildDropFolder + " not found, this mean no networkproblem, this build is deleted, do nothing", 0, 0);
                            logger.LogMessageToFile(fileException.Message + "---   " + fileException.StackTrace, 0, 0);
                        }
                        else
                        {
                            logger.LogMessageToFile(baseFolder + "--- folder not exist, " + mTestApp, 0, 0);
                            System.Windows.Forms.MessageBox.Show(baseFolder + "--- folder not exist " + mTestApp);
                            throw new Exception(baseFolder + "--- folder not exist, ");
                        }
                    }

                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.Forms.MessageBox.Show("Deployment Completed:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);
                }
                #endregion

                sDeploymentEndTime = DateTime.Now;
                sIsDeployed = true;
                // end deployment 

                // start testing
                // // clear test directory first : remove  xls file
                string deletePathXls = ConstCommon.ETRICC_TESTS_DIRECTORY + @"\*.xls";
                string msg = "Before testing start, first clear test directory-->delete files:" + deletePathXls;
                logger.LogMessageToFile(msg, 0, 0);
                if (!FileManipulation.DeleteFilesWithWildcards(deletePathXls, ref msg))
                    throw new Exception(msg);

                // 
                //m_TestWorkingDirectory = System.IO.Directory.GetCurrentDirectory();
                m_TestWorkingDirectory = TestTools.OSVersionInfoClass.ProgramFilesx86() + @"\Dematic\AutomaticTesting";
                logger.LogMessageToFile("<----->  starting testing  mTestApp:" + mTestApp, 0, 0);

                if (mTestApp.Equals(TestApp.EPIA4) ||
                    mTestApp.Equals(TestApp.EPIANET45) ||
                    mTestApp.Equals(Constants.ETRICCUI))
                {
                    /*mExcelVisible = m_Settings.ExcelVisible;
                    string AllowFunctionalTesting = m_Settings.FunctionalTesting.ToString().ToLower();
                    mServerRunAs = m_Settings.ServerRunAs;
                    mMail = m_Settings.Mail.ToString().ToLower();*/

                    mExcelVisible = m_TestConfigSettings.Element.Excel.ToString();
                    string AllowFunctionalTesting = m_TestConfigSettings.Element.FunctionalTesting.ToString().ToLower();
                    mServerRunAs = m_TestConfigSettings.Element.ServerRunAs;
                    mMail = m_TestConfigSettings.Element.Mail.ToString().ToLower();

                    string EpiaArgs = '"' + m_installScriptDir + '"'            //  0
                        + " " + '"' + m_ValidatedBuildDropFolder + '"'           //  1
                        + " " + '"' + m_BuildNumber + '"'                       //  2  
                        + " " + '"' + mTeamProject + '"'                        //  3
                        + " " + '"' + mTestApp + '"'                            //  4
                        + " " + '"' + mTargetPlatform + '"'                     //  5
                        + " " + '"' + mCurrentPlatform + '"'                    //  6
                        + " " + '"' + mTestDef + '"'                           //  7
                        + " " + '"' + mCalledProgram + '"'                      //  8
                        + " " + '"' + TESTTOOL_VERSION + '"'                    //  9
                        + " " + '"' + m_TestAutoMode.ToString().ToLower() + '"' //  10
                        + " " + '"' + Tfs.ServerUrl + '"'                       //  11
                        + " " + '"' + mServerRunAs + '"'                        //  12
                        + " " + '"' + mExcelVisible + '"'                       //  13
                        + " " + '"' + sDemonstration.ToString().ToLower() + '"' //  14    // from config file
                        + " " + '"' + mMail + '"'                               //  15
                        + " " + '"' + mTestDefFile + '"'                        //  16 TestDefinitionFilename="Epia4TestDefinition.txt
                        + " " + '"' + sInfoFileKey + '"'                        //  17 info file key: Windows7.64.AnyCPU.Protected.EPIAAUTOTEST1
                        + " " + '"' + sNetworkMap + '"'                         //  18 \\FileCluster2.Ecorp.Int\TFSDROPFOLDER\Builds
                        +" " + '"' + sDemoCaseCount + '"';                      //  19
                    
                    string EtriccArgs = string.Empty;
                    //mTestAppDirectory = m_TestWorkingDirectory;
                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.Forms.MessageBox.Show("" + m_TestWorkingDirectory, "TRY TO find exe");

                    string filename = string.Empty;
                    if (mTestApp.Equals(TestApp.EPIA4))
                    {
                        #region
                        filename = "Egemin.Epia.Testing.QATestEpiaUI.exe";
                        if (m_installScriptDir.IndexOf("Protected") > 0)
                        {
                            filename = "Egemin.Epia.Testing.QATestEpiaProtected.exe";
                            if (sMsgDebug.StartsWith("true"))
                                System.Windows.Forms.MessageBox.Show("" + m_installScriptDir, "Test");
                        }
                        #endregion
                    }
                    else if (mTestApp.Equals(Constants.ETRICCUI))
                    {
                        #region
                        filename = "Egemin.Epia.Testing.QATestEtriccUI.exe";
                        //MessageBox.Show("1: "+dir);
                        // at TFS test application is at same location as deployment application
                        //dir = System.IO.Directory.GetCurrentDirectory();
                        //string testpath = System.IO.Path.Combine(mTestAppDirectory, filename); // 
                        string testpath = System.IO.Path.Combine(m_TestWorkingDirectory, filename); // 
                        logger.LogMessageToFile("testpath:" + testpath, 0, 0);

                        string x = "@" + '"' + sEtricc5InstallationFolder + '"';

                        sEtricc5InstallationFolder = System.IO.Path.GetDirectoryName(sEtricc5InstallationFolder);

                        EtriccArgs = '"' + m_installScriptDir + '"'         //  0
                        + " " + '"' + m_ValidatedBuildDropFolder + '"'           //  1
                        + " " + '"' + m_BuildNumber + '"'                       //  2  
                        + " " + '"' + mTeamProject + '"'                        //  3
                        + " " + '"' + mTestApp + '"'                            //  4
                        + " " + '"' + mTargetPlatform + '"'                     //  5
                        + " " + '"' + mCurrentPlatform + '"'                    //  6
                        + " " + '"' + mTestDef + '"'                           //  7
                        + " " + '"' + mCalledProgram + '"'                      //  8
                        + " " + '"' + TESTTOOL_VERSION + '"'                    //  9
                        + " " + '"' + m_TestAutoMode.ToString().ToLower() + '"' //  10
                        + " " + '"' + Tfs.ServerUrl + '"'              //  11
                        + " " + '"' + mServerRunAs + '"'                         //  12
                        + " " + '"' + mExcelVisible + '"'                        //  13
                        + " " + '"' + sDemonstration.ToString().ToLower() + '"' //  14    // from config file
                        + " " + '"' + mMail + '"'                               //  15                        
                        + " " + '"' + AllowFunctionalTesting + '"'    //  16
                        + " " + '"' + sProjectFile + '"'             //  17
                        + " " + '"' + sEtricc5InstallationFolder + '"'  //  18
                        + " " + '"' + mTestDefFile + '"'                         //  19 TestDefinitionFilename="EtriccTestDefinition.txt
                        + " " + '"' + sInfoFileKey + '"'                       //  20 info file key: Windows7.64.AnyCPU.Debug.ETRICCAUTOTEST1
                        + " " + '"' + sNetworkMap + '"';                         //  21 \\FileCluster2.Ecorp.Int\TFSDROPFOLDER\Builds
                        #endregion
                    }
                    else if (mTestApp.Equals(TestApp.EPIANET45))
                    {
                        #region  //Start EPIANET45
                        //filename = "Egemin.Epia.Testing.QATestEpiaNet45UI.exe";
                        filename = "Egemin.Epia.Testing.QATestEpiaUI.exe";
                        EpiaArgs = '"' + m_installScriptDir + '"'           //  0
                        + " " + '"' + m_ValidatedBuildDropFolder + '"'           //  1
                        + " " + '"' + m_BuildNumber + '"'                       //  2  
                        + " " + '"' + mTeamProject + '"'                        //  3
                        + " " + '"' + mTestApp + '"'                            //  4
                        + " " + '"' + mTargetPlatform + '"'                     //  5
                        + " " + '"' + mCurrentPlatform + '"'                    //  6
                        + " " + '"' + mTestDef + '"'                           //  7
                        + " " + '"' + mCalledProgram + '"'                      //  8
                        + " " + '"' + TESTTOOL_VERSION + '"'                    //  9
                        + " " + '"' + m_TestAutoMode.ToString().ToLower() + '"' //  10
                        + " " + '"' + Tfs.ServerUrl + '"'                       //  11
                        + " " + '"' + mServerRunAs + '"'                        //  12
                        + " " + '"' + mExcelVisible + '"'                       //  13
                        + " " + '"' + sDemonstration.ToString().ToLower() + '"' //  14    // from config file
                        + " " + '"' + mMail + '"'                               //  15
                        + " " + '"' + mTestDefFile + '"'                        //  16 TestDefinitionFilename="Epia4TestDefinition.txt
                        + " " + '"' + sInfoFileKey + '"'                        //  17 info file key: Windows7.64.AnyCPU.Protected.EPIAAUTOTEST1
                        + " " + '"' + sNetworkMap + '"';                         //  18 \\FileCluster2.Ecorp.Int\TFSDROPFOLDER\Builds
                        #endregion
                    }

                    string arg = EpiaArgs;
                    if (mTestApp.StartsWith("Etricc"))
                        arg = EtriccArgs;

                    logger.LogMessageToFile(" TestApplication args :" + arg, 0, 0);
                    Log(" TestApplication args:" + arg);
                    Log(" filename:" + filename);
                    Log(System.Environment.NewLine + " < ----------------->TestApplicationDir:" + m_TestWorkingDirectory);
                    // UI Test started
                    logger.LogMessageToFile(System.Environment.NewLine + "<----------------->  TestApplicationDir:" + m_TestWorkingDirectory, 0, 0);

                    try
                    {
                        System.Diagnostics.Process proc5 = new System.Diagnostics.Process();
                        proc5.EnableRaisingEvents = false;
                        proc5.StartInfo.FileName = filename;
                        proc5.StartInfo.Arguments = arg;
                        proc5.StartInfo.WorkingDirectory = m_TestWorkingDirectory;
                        proc5.Start();
                    }
                    catch (Exception ex)
                    {
                            Log(" ex.Message:" + ex.Message );
                            Log(" 1:" + System.Environment.NewLine);
                            Log(" ex.StackTrace:" + ex.StackTrace);
                            Log(" 2:" + System.Environment.NewLine);
                            throw new Exception("Start tesing App exception: appname=" + filename + "--- WOrkingDir:" + m_TestWorkingDirectory
                            + "--- arg:" + arg,
                            ex);
                    }


                    m_TestAutoMode = false;
                }
                else if (mTestApp.Equals(Constants.ETRICC5))
                {
                    #region  //Start Etricc 5 Testing
                    //MessageBox.Show("Start Etricc 5 Test:" + m_BuildNumber);
                    logger.LogMessageToFile("Start Etricc 5 Test:" + m_BuildNumber, 0, 0);

                    // unzip project file, only test Eurobaltic Project
                    //string projectXml = m_Settings.SelectedProjectFile;
                    string projectXml = m_TestConfigSettings.Element.SelectedProjectFile;
                    if (projectXml.Equals("Demo.xml") || projectXml.IndexOf("Eurobaltic") >= 0)
                    {
                        //sProjectFile = "Demo.xml";
                        sProjectFile = "EurobalticWorker.xml";

                        if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))
                            projectXml = Path.Combine(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "EurobalticWorker.zip");
                        else
                            projectXml = m_CurrentDrive + @"Team Systems\Epia 3\Testing\Automatic\AutomaticTests\TestData\Etricc5\EurobalticWorker.zip";
                        string zipFile = projectXml;
                        FastZip fz = new FastZip();
                        fz.ExtractZip(zipFile, ConstCommon.ETRICC_TESTS_DIRECTORY, "");
                        Thread.Sleep(10000);

                        //string Xml = @"C:\Testing\Eurobaltic.xml";
                        //string xmlPath = installPath + "\\Etricc\\server\\Data\\Etricc\\Demo.xml";
                        //string xmlPathBackup = installPath + "\\Etricc\\server\\Data\\Etricc\\DemoBackup.xml";
                        //System.IO.File.Delete(xmlPath);
                        //System.IO.File.Copy(Xml, xmlPath, true);
                        //sProjectFile = System.IO.Path.GetFileName(projectXml);

                    }
                    else  // 
                    {   // TestProject File
                        sProjectFile = "TestProjectWorker.xml";
                        if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))
                            projectXml = Path.Combine(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "TestProjectWorker.zip");
                        else
                            projectXml = m_CurrentDrive + @"Team Systems\Epia 3\Testing\Automatic\AutomaticTests\TestData\Etricc5\TestProjectWorker.zip";
                        string zipFile = projectXml;
                        FastZip fz = new FastZip();
                        fz.ExtractZip(zipFile, ConstCommon.ETRICC_TESTS_DIRECTORY, "");

                        string Xml = Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "TestProjectWorker.xml");
                        while (!System.IO.File.Exists(Xml))
                        {
                            logger.LogMessageToFile("---- TestProject.XML Not copied yet, wait:" + Xml, 0, 0);
                            Thread.Sleep(3000);
                        }
                        //string xmlPathBackup = installPath + "\\Etricc\\Server\\Data\\Etricc\\DemoBackup.xml";
                        //System.IO.File.Delete(xmlPath);
                        //System.IO.File.Copy(Xml, xmlPath, true);
                        //sProjectFile = System.IO.Path.GetFileName(projectXml);
                        Thread.Sleep(1000);
                    }

                    #region  // update build quality to "GUI Tests Started"
                    /*msgX = "update build quality GUI Tests Started";
                    TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                    while (TFSConnected == false)
                    {
                        TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                               "update build quality GUI Tests Started", (uint)Tfs.ReconnectDelay);
                        System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                    }*/
                    if (TFSConnected)
                    {
                        Log(TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TfsUtilities.GetProjectName(mTestApp), "GUI Tests Started",
                        //    m_BuildSvc, sDemonstration));
                        m_BuildSvc, "false"));   // only for demo
                        //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "GUI Tests Started", m_BuildSvc);
                        string quality2 = m_BuildSvc.GetMinimalBuildDetails(m_Uri).Quality;
                        logger.LogMessageToFile(")))))))))))) start test worker:::::::::::::::::::quality:" + quality2, 0, 0);
                    }

                    if (m_TestAutoMode)
                    {
                        if (mTestApp.Equals(TestApp.EPIA4) || mTestApp.Equals(Constants.ETRICCUI))
                        {
                            if (m_installScriptDir.IndexOf("Protected") > 0)
                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", "Tests Started", sInfoFileKey);
                            else
                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", "Tests Started", sInfoFileKey);
                        }
                        else
                            TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", "Tests Started", sInfoFileKey);
                        logger.LogMessageToFile(m_ValidatedBuildDropFolder + "GUI Tests Started : m_testApp " + mTestApp, 0, 0);
                    }

                    #endregion

                    //   TestWorker starting... 
                    logger.LogMessageToFile(")))))))))))) start test worker:::::::::::::::::::cnt:", 0, 0);
                    logger.LogMessageToFile("sEtricc5InstallationFolder:" + sEtricc5InstallationFolder, 0, 0);
                    logger.LogMessageToFile("sProjectFile:" + sProjectFile, 0, 0);
                    //
                    StartTestWorker(sEtricc5InstallationFolder, ConstCommon.ETRICC_TESTS_DIRECTORY, sProjectFile);
                    //
                    int cnt = 1;
                    string exceptionMsg = string.Empty;
                    DateTime StartTime = DateTime.Now;
                    TimeSpan Time = DateTime.Now - StartTime;
                    // start time 
                    while (cnt > 0)
                    {
                        cnt = 0;   // start with cnt = 0 --> assume test is finished, check status
                        Thread.Sleep(15000);
                        // check output has record
                        StreamReader reader = File.OpenText(Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "TestAutoDeploymentOutputLog.txt"));
                        while (reader.ReadLine() != null)
                            cnt++;  // cnt == 1 ----> test is still running
                        reader.Close();
                        // if time span < 15 min
                        // Try to catch Activat Error
                        //MessageBox.Show("cnt="+cnt+"  --- msgLength="+exceptionMsg.Length+" -- time.min="+
                        //    Time.Minutes);
                        while (cnt == 1 && exceptionMsg.Length == 0 && Time.Minutes <= 10)
                        {
                            if (CatchExceptionMessage(ref exceptionMsg))
                            {
                                cnt = 0;
                                logger.LogMessageToFile(")))))))))))) exception message found:" + exceptionMsg, 0, 0);
                            }
                            else
                            {
                                logger.LogMessageToFile(")))))))))))) exception message not found :::::::::::::::::::cnt:" + cnt, 0, 0);
                                break;
                            }

                            Time = DateTime.Now - StartTime;
                        }
                        Thread.Sleep(15000);
                        //ClickUiScreenActionToAvoidScreenStandBy();
                    }

                    logger.LogMessageToFile(")))))))))))) end test worker:::::::::::::::::::cnt:" + cnt, 0, 0);

                    //----------------------------------
                    string GUIstatus = "GUI Tests Passed";
                    StreamReader readerWorker = File.OpenText(Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "TestResults.txt"));
                    string results = readerWorker.ReadLine();
                    readerWorker.Close();
                    logger.LogMessageToFile(")))))))))))) update quility:::result:" + results, 0, 0);

                    if (exceptionMsg.Length > 10)
                    {
                        GUIstatus = "GUI Tests Failed";
                    }
                    else
                    {
                        if (results.Trim().ToLower().StartsWith("ok"))
                        {
                            logger.LogMessageToFile(")))))))))))) result:ok" + results, 0, 0);
                            GUIstatus = "GUI Tests Passed";
                        }
                        else
                            GUIstatus = "GUI Tests Failed";
                    }

                    #region  // update build quality to "GUI Tests"
                    msgX = "update build quality GUI Status";
                    TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                    while (TFSConnected == false)
                    {
                        TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                               "update build quality GUI Status", (uint)Tfs.ReconnectDelay);
                        System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                    }
                    if (TFSConnected)
                    {
                        Log(TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TfsUtilities.GetProjectName(mTestApp), GUIstatus,
                        //    m_BuildSvc, sDemonstration));
                         m_BuildSvc, "false"));// only for demo
                        //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), GUIstatus, m_BuildSvc);
                    }

                    if (m_TestAutoMode)
                    {
                        if (exceptionMsg.Length > 10)
                        {
                            if (mTestApp.Equals(TestApp.EPIA4) || mTestApp.Equals(Constants.ETRICCUI))
                            {
                                if (m_installScriptDir.IndexOf("Protected") > 0)
                                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Exception", exceptionMsg, sInfoFileKey);
                                else
                                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Exception", exceptionMsg, sInfoFileKey);
                            }
                            else
                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Exception", exceptionMsg, sInfoFileKey);
                            logger.LogMessageToFile(m_ValidatedBuildDropFolder + "exceptionMsg : mTestApp " + mTestApp, 0, 0);
                        }
                        else
                        {
                            if (mTestApp.Equals(TestApp.EPIA4) || mTestApp.Equals(Constants.ETRICCUI))
                            {
                                if (m_installScriptDir.IndexOf("Protected") > 0)
                                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, GUIstatus, mTestApp, sInfoFileKey);
                                else
                                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, GUIstatus, mTestApp, sInfoFileKey);
                            }
                            else
                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, GUIstatus, mTestApp, sInfoFileKey);
                            logger.LogMessageToFile(m_ValidatedBuildDropFolder + "GUIstatus : m_testApp " + mTestApp, 0, 0);
                        }
                    }
                    #endregion

                    ProcessUtilities.CloseProcess("EPIA.Launcher");
                    ProcessUtilities.CloseProcess("EPIA.Explorer");

                    if (m_TestAutoMode)
                        FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");

                    logger.LogMessageToFile(" **************** ( END ETRICC 5 TESTING )************************** ", 0, 0);
                    #endregion
                }
                else if (mTestApp.Equals(Constants.ETRICCSTATISTICS))
                {
                    #region  //Start Etricc Statistics Testing
                    //mExcelVisible = m_Settings.ExcelVisible;
                    mExcelVisible = m_TestConfigSettings.Element.Excel.ToString();
                    //string AllowFunctionalTesting = m_Settings.FunctionalTesting.ToString().ToLower();
                    //mServerRunAs = m_Settings.ServerRunAs;
                    mServerRunAs = m_TestConfigSettings.Element.ServerRunAs;
                    //mMail = m_Settings.Mail.ToString().ToLower();
                    mMail = m_TestConfigSettings.Element.Mail.ToString().ToLower();
                    string EpiaArgs = '"' + m_installScriptDir + '"'            //  0
                        + " " + '"' + m_ValidatedBuildDropFolder + '"'           //  1
                        + " " + '"' + m_BuildNumber + '"'                       //  2  
                        + " " + '"' + mTeamProject + '"'                        //  3
                        + " " + '"' + mTestApp + '"'                            //  4
                        + " " + '"' + mTargetPlatform + '"'                     //  5
                        + " " + '"' + mCurrentPlatform + '"'                    //  6
                        + " " + '"' + mTestDef + '"'                           //  7
                        + " " + '"' + mCalledProgram + '"'                      //  8
                        + " " + '"' + TESTTOOL_VERSION + '"'                    //  9
                        + " " + '"' + m_TestAutoMode.ToString().ToLower() + '"' //  10
                        + " " + '"' + Tfs.ServerUrl + '"'              //  11
                        + " " + '"' + mServerRunAs + '"'                         //  12
                        + " " + '"' + mExcelVisible + '"'                        //  13
                        + " " + '"' + sDemonstration.ToString().ToLower() + '"' //  14    // from config file
                        + " " + '"' + mMail + '"'                              //  15
                        + " " + '"' + mTestDefFile + '"'                         //  16 TestDefinitionFilename="Epia4TestDefinition.txt
                        + " " + '"' + sInfoFileKey + '"'                       //  17 info file key: Windows7.64.AnyCPU.Protected.EPIAAUTOTEST1
                        + " " + '"' + sNetworkMap + '"'                          //  18 \\FileCluster2.Ecorp.Int\TFSDROPFOLDER\Builds
                        + " " + '"' + m_ValidatedTestBuildObject.BuildDef + '"'; //  19 \\ build definition

                    //mTestAppDirectory = m_TestWorkingDirectory;
                    string filename = "Egemin.Epia.Testing.QATestEtriccStatistics.exe";
                    //MessageBox.Show("1: "+dir);
                    // at TFS test application is at same location as deployment application
                    //dir = System.IO.Directory.GetCurrentDirectory();
                    string testpath = System.IO.Path.Combine(m_TestWorkingDirectory, filename);

                    string arg = EpiaArgs;

                    logger.LogMessageToFile(" TestApplication args :" + arg, 0, 0);
                    // UI Test started
                    logger.LogMessageToFile("TestWorkingDirectoryDir:" + m_TestWorkingDirectory, 0, 0);

                    System.Diagnostics.Process proc5 = new System.Diagnostics.Process();
                    proc5.EnableRaisingEvents = false;
                    proc5.StartInfo.FileName = filename;
                    proc5.StartInfo.Arguments = arg;
                    proc5.StartInfo.WorkingDirectory = m_TestWorkingDirectory;
                    proc5.Start();

                    m_TestAutoMode = false;

                    #endregion
                }

                // Update testinfo file to status GUI Tests Started
                logger.LogMessageToFile(" UpdateStatusInTestInfoFile GUI Tests Started Test started:" + sTestResultFolder, 0, 0);
                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", "Test started", sInfoFileKey);
            }
            catch (Exception ex)
            {
                // Your error handler here
                sErrorMessage = ex.Message + System.Environment.NewLine + ex.StackTrace
                    + "-- m_installScriptDir" + m_installScriptDir;
                //System.Windows.Forms.MessageBox.Show(sErrorMessage, "Started test application dir:" + mTestAppDirectory);

                Log(ex.Message + System.Environment.NewLine + ex.StackTrace);

                // Update build quality
                msgX = "update build quality Deployment Failed6";
                TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                while (TFSConnected == false)
                {
                    TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                           "update build quality Deployment Failed6", (uint)Tfs.ReconnectDelay);
                    System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                    TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
                }

                if (TFSConnected)
                {
                    TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Exception", sErrorMessage, sInfoFileKey);

                    /*Console.WriteLine(" Update build quality:  quality: GUI Tests Failed");
                    string updateResult = TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TfsUtilities.GetProjectName(mTestApp),
                    "GUI Tests Failed", 
                    //m_BuildSvc, sDemonstration);
                      m_BuildSvc, "false");// only for demo

                    Log(updateResult);
                    if (updateResult.StartsWith("Error"))
                    {
                        System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                        throw new Exception(updateResult);
                    }
                    */
                    logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Exception:  " + sErrorMessage, 0, 0);

                    //Log(TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(mTestApp), "GUI Tests Failed",
                    //    m_BuildSvc, sDemonstration));

                    if (m_TestAutoMode)
                    {
                        if (ex.Message.IndexOf("RecompileTestRuns") >= 0)  // willbecheck later
                        {
                            string msg = "Recompile TestRuns Exception:" + ex.Message + "---" + ex.StackTrace;
                            TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Completed", msg, sInfoFileKey);
                            logger.LogMessageToFile(m_ValidatedBuildDropFolder + "Recompile TestRuns Exception: : mTestApp " + mTestApp, 0, 0);
                        }
                        else
                        {
                            string msg = "Deployment Exception:" + ex.Message + "---" + ex.StackTrace;
                            if (mTestApp.Equals(TestApp.EPIA4) || mTestApp.Equals(Constants.ETRICCUI) || mTestApp.Equals(Constants.ETRICCSTATISTICS))
                            {
                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Exception", msg, sInfoFileKey);
                            }
                            else
                            {
                                TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Exception", msg, sInfoFileKey);
                            }

                            logger.LogMessageToFile(m_ValidatedBuildDropFolder + "Deployment Exception: : mTestApp " + mTestApp + "+" + mCurrentPlatform, 0, 0);
                        }
                    }
                }
            }
            finally
            {
                m_State = STATE.PENDING;
            }
            #endregion

            }
            catch (Exception ex )
            {
                Log(" ---  Start2 exception:" + ex.Message + "--- " + ex.StackTrace);
                //logger.LogMessageToFile(" Start2 exception:" + ex.Message +"--- " + ex.StackTrace, 0, 0);
            }
            finally
            {
                m_State = STATE.PENDING;
            }
        }

        public BuildObject GetNewValidatedBuildDirectory(List<BuildObject> allBuildsInfo, string testPC, /*string testApp,*/ /*string platform, */
             ref string currentPlatform, ref string configuration, ref string infoKey)   // platform : AnyCPU, x86    // configuration : Debug, Release, Protected
        {
            string testApp = "DefaultTestApp";
            string currentPlatformApp = "AnyCPU.Debug";
            Dictionary<string, string> validBuild = new Dictionary<string, string>();
            BuildObject currenBuildObject = null;
            string currentBuildNr = string.Empty;
            string currentBuildLocation = string.Empty;

            IEnumerator EmpEnumerator = allBuildsInfo.GetEnumerator(); //
            EmpEnumerator.Reset(); //Position at the Beginning
            while (EmpEnumerator.MoveNext()) //
            {
                currenBuildObject = (BuildObject)EmpEnumerator.Current;
                currentBuildNr = currenBuildObject.BuildNr;

                //testApp = TfsUtilities.GetTestAppFromBuildNumber(currentBuildNr);   // get all info from buildnr
                testApp = TfsUtilities.GetTestAppFromBuildDefinition(currenBuildObject.BuildDef);   // get all info from buildnr
                logger.LogMessageToFile(" testApp:" + testApp, 0, 0);

                currentBuildLocation = currenBuildObject.DripLoc;
                if (sMsgDebug.StartsWith("true"))
                {
                    System.Windows.Forms.MessageBox.Show(" currentBuildNr:" + currentBuildNr, "current build");
                    System.Windows.Forms.MessageBox.Show(" currentBuildLocation:" + currentBuildLocation, "current build");
                }

                sTestResultFolder = currentBuildLocation + "\\" + Constants.sTestResultFolderName;
                string testInfoTxtFile = Path.Combine(sTestResultFolder, ConstCommon.TESTINFO_FILENAME);
                logger.LogMessageToFile(" testInfoTxtFile:" + testInfoTxtFile, 0, 0);

                try
                {
                    //System.IO.Directory.CreateDirectory(sTestResultFolder);
                    if (!System.IO.Directory.Exists(sTestResultFolder))
                        System.IO.Directory.CreateDirectory(sTestResultFolder);

                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.MessageBox.Show("testResultDirectory:" + sTestResultFolder);

                    // check TestInfo.txt header
                    //string testWorkingFile = Path.Combine(sTestResultFolder, ConstCommon.TESTWORKING_FILENAME);
                    // TestInfo file exist
                    if (File.Exists(testInfoTxtFile))
                    {
                        Thread.Sleep(1000);
                    }
                    else // create initial testinfo log file which is copy from testdefinition file depended on test projectr 
                    {
                        string msgX = "Get QA Definition Files";
                        bool TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX, ref tfsProjectCollection);
                        while (TFSConnected == false)
                        {
                            System.Windows.Forms.DialogResult dr = System.Windows.Forms.MessageBox.Show("Are you want continue to connect again?",
                                "TFS connection failed",
                                System.Windows.Forms.MessageBoxButtons.YesNo);
                            if (dr == System.Windows.Forms.DialogResult.Yes)
                            {
                                TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                        "Get QA Definition Files", (uint)Tfs.ReconnectDelay);
                                System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                                TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX, ref tfsProjectCollection);
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (TFSConnected)
                        {
                            string sTestDefinitionFilesPath = System.Configuration.ConfigurationManager.AppSettings.Get("TestDefinitionFilePath");
                            string sQAFolder = sTestDefinitionFilesPath.Substring(0, sTestDefinitionFilesPath.LastIndexOf('\\'));
                            if (TfsUtilities.GetTestProjectQA(tfsProjectCollection, sQAFolder, ref msgX) == false)
                            {
                                //System.Windows.Forms.MessageBox.Show("" + msgX, "Get test definition files from TFS failed");
                                Console.WriteLine(msgX);
                            }
                        }

                        logger.LogMessageToFile(" create empty testinfo file: TestInfo:" + testInfoTxtFile, 0, 0);
                        DeployUtilities.CreateInitialTestInfoFile(testApp, testInfoTxtFile);
                    }
                }
                catch (Exception ex)
                {   // during testing build maybe deleted, just catch exception and check next build
                    logger.LogMessageToFile(" create TestResultFolder or empty testinfo file exception:" + ex.Message + "---" + ex.StackTrace, 0, 0);
                    continue;
                }

                // get test platform array : AnyCPU.Debug, AnyCPU.Release, AnyCPU.Protected,x86.Debug etc
                logger.LogMessageToFile(" get test platform array : AnyCPU.Debug, AnyCPU.Release, AnyCPU.Protected,x86.Debug etc:testApp:" + testApp, 0, 0);
                string key = DeployUtilities.getThisPCOS();
                logger.LogMessageToFile(" get test platform array : AnyCPU.Debug, AnyCPU.Release, AnyCPU.Protected,x86.Debug etc:key:" + key, 0, 0);
                string[] TestPlatformArray = DeployUtilities.GetTestPlatformArray(testApp, testInfoTxtFile);
                logger.LogMessageToFile(" TestPlatformArray.Length:" + TestPlatformArray.Length, 0, 0);

                if (testApp.Equals(TestApp.EPIA4) || testApp.Equals(Constants.ETRICCSTATISTICS) || testApp.Equals(Constants.ETRICC5)
                    || testApp.Equals(Constants.ETRICCUI))
                {
                    #region // check is tested by current OS
                    string caseApp = string.Empty;
                    bool foundTestVersion = false;
                    for (int i = 0; i < TestPlatformArray.Length; i++)
                    {
                        foundTestVersion = false;
                        //if (IsThisPCTested(testInfoTxtFile, testPC, testApp, TestPlatformArray[i]) == false)
                        if (DeployUtilities.IsThisOSPlatformTested(testInfoTxtFile, logger, testApp, TestPlatformArray[i]) == false)
                        {
                            currentPlatform = TestPlatformArray[i].Substring(0, TestPlatformArray[i].IndexOf('.'));   // AnyCPU
                            currentPlatformApp = TestPlatformArray[i];                                                  // AnyCPU.Debug
                            configuration = TestPlatformArray[i].Substring(TestPlatformArray[i].IndexOf('.') + 1);    // Debug
                            if (configuration.IndexOf('.') >= 0)
                                configuration = configuration.Substring(0, configuration.IndexOf('.'));
                            foundTestVersion = true;
                            break;
                        }
                    }

                    if (foundTestVersion == false)
                    {
                        logger.LogMessageToFile("current build ..." + currentBuildNr +
                                    "is already tested,  ", sLogCount, sLogInterval);
                        currenBuildObject = null;
                        continue;
                    }
                    #endregion
                }
                else if (testApp.Equals(TestApp.EPIANET45)) //  
                {
                    #region // check is tested by current OS
                    string caseApp = string.Empty;
                    bool foundTestVersion = false;
                    for (int i = 0; i < TestPlatformArray.Length; i++)
                    {
                        //logger.LogMessageToFile(" TestPlatformArray["+i+"]:" + TestPlatformArray[i], 0, 0);
                        foundTestVersion = false;
                        //if (IsThisPCTested(testInfoTxtFile, testPC, testApp, TestPlatformArray[i]) == false)
                        if (DeployUtilities.IsThisOSPlatformTested(testInfoTxtFile, logger, testApp, TestPlatformArray[i]) == false)
                        {
                            currentPlatform = TestPlatformArray[i].Substring(0, TestPlatformArray[i].IndexOf('.'));   // AnyCPU
                            currentPlatformApp = TestPlatformArray[i];                                                  // AnyCPU.Debug
                            configuration = TestPlatformArray[i].Substring(TestPlatformArray[i].IndexOf('.') + 1);    // Debug
                            if (configuration.IndexOf('.') >= 0)
                                configuration = configuration.Substring(0, configuration.IndexOf('.'));
                            foundTestVersion = true;
                            break;
                        }
                    }

                    if (foundTestVersion == false)
                    {
                        logger.LogMessageToFile("current build ..." + currentBuildNr +
                                    "is already tested,  ", sLogCount, sLogInterval);
                        currenBuildObject = null;
                        continue;
                    }
                    #endregion
                }
                else if (testApp.Equals(Constants.EWMS)) //  current_testApp will be decided next
                {
                    string ewmsBuildLogFilePath = string.Empty;
                    //buildlogfile = currentBuildLocation + ewmsBuildLogFilePath;
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Wrong testApp:" + testApp, "xyz");
                    testApp = "Wrong testApp:";
                }

                #region add this test info to info file
                infoKey = DeployUtilities.getThisPCOS() + "[" + currentPlatform + "." + configuration + "]" + testPC;
                //System.Windows.Forms.MessageBox.Show("infoKey:" + infoKey, "zzzsedf");

                if (File.Exists(testInfoTxtFile))
                {
                    #region  // Add test info                                                                                                                                                                                                                                                       #region
                    StreamReader readerInfo = File.OpenText(testInfoTxtFile);
                    string info = readerInfo.ReadToEnd();
                    readerInfo.Close();

                    logger.LogMessageToFile(" read testinfo file: content is " + info, 0, 0);
                    if (testApp.Equals(TestApp.EPIA4) || testApp.Equals(Constants.ETRICCUI) || testApp.Equals(Constants.ETRICCSTATISTICS)
                        || testApp.Equals(TestApp.EPIANET45))
                    {
                        info = info + infoKey + "-" + "Deployment Starting:" + ":" + "Starting";
                    }

                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.Forms.MessageBox.Show("Logline:" + info);

                    Log(testPC + " Added into InfoFile of build:" + currentBuildNr);
                    logger.LogMessageToFile(testPC + " Added into Info file:" + currentBuildNr, sLogCount, sLogInterval);

                    // ------------  if write infotext file failure, not do anything , continue go to iteration 
                    try
                    {
                        string pKey = DeployUtilities.getThisPCOS() + "[" + currentPlatform + "." + configuration + "]";
                        DeployUtilities.AddUpdateStatusInTestInfoFile(testInfoTxtFile, "Deployment started", "Starting", pKey, infoKey, logger);

                        //StreamWriter writeInfo = File.CreateText(testInfoTxtFile);
                        //logger.LogMessageToFile(" write testinfo file: content is " + info, 0, 0);
                        //writeInfo.WriteLine(info);
                        //writeInfo.Close();
                    }
                    catch (Exception ex)
                    {
                        string msg = testPC + " Exception Add " + infoKey + " to InfoText:" + currentBuildLocation + "=====" + ex.Message + "" + ex.StackTrace;
                        Log(msg);
                        logger.LogMessageToFile(msg, sLogCount, sLogInterval);
                        continue;
                    }

                    #endregion
                    validBuild.Add(currentBuildNr, currentBuildLocation);  // will change to found validated build

                }
                #endregion

                if (validBuild.Count == 1)
                {
                    break;
                }
                else
                    currenBuildObject = null;
            }

            return currenBuildObject;
        }

        /// <summary>
        ///  For valid build quality, Check if it has tested by current PC
        /// </summary>
        /// <param name="testInfoTxtFile"></param>
        /// <param name="testPC"></param>
        /// <param name="BuildDefinition: CI, Nightly, Weekly or Version"></param>
        /// <returns></returns>
        public bool IsThisPCTested(string testInfoTxtFile, string testPC, string testApp, string platform)
        {
            string testKey = string.Empty;
            try
            {
                string thisPCOS = DeployUtilities.getThisPCOS();
                StreamReader readerInfo = File.OpenText(testInfoTxtFile);
                string info = readerInfo.ReadToEnd();
                readerInfo.Close();
                testKey = thisPCOS + "[" + platform + "]" + testPC;
                Console.WriteLine("IsThisPCTested: testKey is-->" + testKey);
                if (info.IndexOf(testKey) >= 0)   // Windows7.32.x86Debug.EPIAAUTOTEST1 
                {
                    //Log("but "+testPC + " is already in test info file");
                    //logger.LogMessageToFile(testPC + " is already in test info file", sLogCount, sLogInterval);
                    return true;
                }
                else
                {
                    Log(testKey + " ---------->> is not in test info file");
                    logger.LogMessageToFile(testKey + " ----------> is not in test info file",
                        sLogCount, sLogInterval);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log("IsTestWorking exception:" + testKey + " - message:" + ex.Message + " --- " + ex.StackTrace);
                logger.LogMessageToFile("IsTestWorking exception:" + testKey + "+" + testApp + "+" + platform + "=" + " - message:" + ex.Message + " --- " + ex.StackTrace,
                    sLogCount, sLogInterval);
                return true;
            }
        }

        private void RemoveSetupSilentMode(string AppName)
        {
            logger.LogMessageToFile("------------Start Removed Setup in Silent Mode: AppName: " + AppName, 0, 0);
            string InstallPath = string.Empty;
            string InstallName = string.Empty;
            //find installed msi in reg
            try
            {
                // private const string REGKEY = "Software\\Egemin\\Automatic testing\\";
                RegistryKey key = Registry.CurrentUser.OpenSubKey(REGKEY);
                object keyvalue;

                keyvalue = key.GetValue(AppName + "InstallationPath");
                if (keyvalue != null)
                    InstallPath = keyvalue.ToString();

                keyvalue = key.GetValue(AppName + "InstallationName");
                if (keyvalue != null)
                    InstallName = keyvalue.ToString();

                logger.LogMessageToFile("(1)Removed Setup in Silent Mode from register key info:" + AppName + " -- InstallPath " + InstallPath + " and InstallName " + InstallName, 0, 0);

            }
            catch (System.NullReferenceException ex1)
            {
                //MessageBox.Show(ex.ToString() + System.Environment.NewLine + ex.StackTrace,
                //    "DeployTestLogic.Tester  RemoveSetup: find register Key");
                logger.LogMessageToFile(AppName + " REGKEY not exist:" + ex1.ToString() + System.Environment.NewLine + ex1.StackTrace, 0, 0);

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString() + System.Environment.NewLine + ex.StackTrace,
                    "DeployTestLogic.Tester  RemoveSetup in Silent Mode: find register Key");
                logger.LogMessageToFile(ex.ToString() + System.Environment.NewLine + ex.StackTrace, 0, 0);
            }

            logger.LogMessageToFile("Removed Setup in Silent Mode " + InstallName + " at " + InstallPath, 0, 0);
            string unattendedXmlFilePath = m_CurrentDrive + @"Epia 3\Testing\Automatic\AutomaticTests\TestData\Etricc5";
            AutomationEventHandler UIUninstallAppEventHandler = new AutomationEventHandler(ProjBasicEvent.OnUninstallAppEvent);
            try
            {

                if ((InstallPath == string.Empty) || (InstallName == string.Empty))
                    logger.LogMessageToFile("(1)RegistryKey Empty, cannot removed in silent mode", 0, 0);
                else
                {

                    if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))
                        unattendedXmlFilePath = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY;

                    //remove the setup in silent mode
                    string xmlFile = System.IO.Path.Combine(unattendedXmlFilePath, "UnattendedXmlFile.xml");
                    string uninstallParm = Path.Combine(InstallPath, InstallName);
                    string args = "/passive /x " + '"' + uninstallParm + '"' + " " + "UnattendedXmlFile=" + '"' + xmlFile + '"';

                    // Add Open window Event Handler
                    Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                        AutomationElement.RootElement, TreeScope.Descendants, UIUninstallAppEventHandler);
                    Thread.Sleep(5000);

                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.EnableRaisingEvents = false;
                    proc.StartInfo.FileName = "MsiExec.exe";
                    proc.StartInfo.Arguments = args;
                    proc.StartInfo.WorkingDirectory = InstallPath;
                    proc.Start();
                    proc.WaitForExit();

                    Thread.Sleep(10000);
                    logger.LogMessageToFile("(1)Removed in Silent Mode MsiExec.exe :" + args + " -- WorkingDirectory " + InstallPath, 0, 0);
                }


                //----------------------------------------------------------
                // also remove by UI again to make sure this app is uninstalled 
                //-----------------------------------------------------------
                logger.LogMessageToFile("(2)also remove by UI again tomake sure this app is uninstalled AppName:" + AppName, 0, 0);
                if (AppName.Equals(EgeminApplication.EPIA))
                {
                    logger.LogMessageToFile("remove the setup in interactive mode :" + EgeminApplication.EPIA, 0, 0);
                    if (DeployUtilities.getThisPCOS().StartsWith("WindowsXP.32"))
                    {
                        Console.WriteLine("ProjAppInstall.UninstallApplicationXP");
                        Thread.Sleep(5000);
                        ProjAppInstall.UninstallApplicationXP(EgeminApplication.EPIA, ref sErrorMessage);
                    }
                    else
                    {
                        Console.WriteLine(" ===================== ProjAppInstall.UninstallApplication");
                        Thread.Sleep(5000);
                        ProjAppInstall.UninstallApplication(EgeminApplication.EPIA, ref sErrorMessage);
                    }
                }
                else if (AppName.Equals(Constants.ETRICCUI))
                {
                    logger.LogMessageToFile("remove the setup in interactive mode :" + EgeminApplication.ETRICC_SHELL, 0, 0);
                    ProjAppInstall.UninstallApplication(EgeminApplication.ETRICC_SHELL, ref sErrorMessage);
                }
                else if (AppName.Equals(Constants.ETRICC5))
                {
                    logger.LogMessageToFile("remove the setup in interactive mode :" + EgeminApplication.ETRICC, 0, 0);
                    ProjAppInstall.UninstallApplication(EgeminApplication.ETRICC, ref sErrorMessage);
                }
                else if (AppName.Equals(Constants.ETRICCSTATISTICS_UI))
                {
                    logger.LogMessageToFile("remove the setup in interactive mode :" + EgeminApplication.ETRICC_STATISTICS_UI, 0, 0);
                    ProjAppInstall.UninstallApplication(EgeminApplication.ETRICC_STATISTICS_UI, ref sErrorMessage);
                }
                else if (AppName.Equals(Constants.ETRICCSTATISTICS_PARSERCONFIGURATOR))
                {
                    logger.LogMessageToFile("remove the setup in interactive mode :" + EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR, 0, 0);
                    ProjAppInstall.UninstallApplication(EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR, ref sErrorMessage);
                }
                else if (AppName.Equals(Constants.ETRICCSTATISTICS_PARSER_SETUP))
                {
                    logger.LogMessageToFile("remove the setup in interactive mode :" + EgeminApplication.ETRICC_STATISTICS_PARSER, 0, 0);
                    ProjAppInstall.UninstallApplication(EgeminApplication.ETRICC_STATISTICS_PARSER, ref sErrorMessage);
                }
                else
                    System.Windows.Forms.MessageBox.Show("AppName: " + AppName, "Constants.ETRICCUI :" + Constants.ETRICCUI);

                //Remove the install information from the registry
                //string RegName = AppName;
                //if (AppName.EndsWith(TestTools.ConstCommon.ETRICC_UI))
                //    RegName = Constants.ETRICCUI;
                //WriteInstallationToReg(RegName, string.Empty, string.Empty);
            }
            catch (Exception ex)
            {
                logger.LogMessageToFile("remove the setup in silent mode exception :" + ex.Message + " --- " + ex.StackTrace, 0, 0);
            }
            finally
            {
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent, AutomationElement.RootElement, UIUninstallAppEventHandler);
            }
        }

        public string RecompileTestRuns()
        {
            string output = string.Empty;
            try
            {
                #region // Recompile TestRuns
                //
                // Recompile TestRuns
                //
                if (!Directory.Exists(mTestRunsDirectory))
                {
                    Directory.CreateDirectory(mTestRunsDirectory);
                }

                string deletePathDll = ConstCommon.ETRICC_TESTS_DIRECTORY + @"\TestRuns\bin\debug\Egemin*.dll";
                string msg = "Recompile TestRuns:first remove old files-->delete files:" + deletePathDll;
                if (!FileManipulation.DeleteFilesWithWildcards(deletePathDll, ref msg))
                    throw new Exception(msg);

                string deletePathPdb = ConstCommon.ETRICC_TESTS_DIRECTORY + @"\TestRuns\bin\debug\Egemin*.pdb";
                string msg1 = "Recompile TestRuns:first remove old files-->delete files:" + deletePathPdb;
                if (!FileManipulation.DeleteFilesWithWildcards(deletePathPdb, ref msg1))
                    throw new Exception(msg1);

                //string deleteInteropDll = Constants.ETRICC_TESTS_DIRECTORY + @"\TestRuns\bin\debug\Interop*.dll";
                //DeleteFilesWithWildcards(deleteInteropDll);

                Thread.Sleep(3000);

                sEtricc5InstallationFolder = OSVersionInfoClass.ProgramFilesx86() + "\\Egemin\\Etricc Server";
                string origPath = sEtricc5InstallationFolder + @"\Egemin*.dll";
                string destPath = ConstCommon.ETRICC_TESTS_DIRECTORY + @"\TestRuns\bin\debug\";

                m_TestWorkingDirectory = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY;
                //System.Windows.Forms.MessageBox.Show(origPath, "origPath");
                //System.Windows.Forms.MessageBox.Show(destPath, "destPath");
                //System.Windows.Forms.MessageBox.Show(m_TestWorkingDirectory, "m_TestWorkingDirectory");
                
                string msg2 = "Recompile TestRuns:";
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

                //dllPath = m_SystemDrive + @"Program Files\\Egemin\\Epia " + mTestedVersion + "\\";
                dllPath = sEtricc5InstallationFolder;
                // DOTNET Version is defined in config file
                if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))  //  Deployed Testing
                {
                    //System.Windows.Forms.MessageBox.Show(m_TestWorkingDirectory, "m_TestWorkingDirectory");
                    arg = "/debug /target:library /out:" + Qmark + ConstCommon.ETRICC_TESTS_DIRECTORY + @"\TestRuns\bin\debug\TestRuns.dll" + Qmark;
                    arg = arg + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestWorker.cs" + Qmark
                         + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestProjectWorker.cs" + Qmark
                         + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestConstants.cs" + Qmark
                         + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\Logger.cs" + Qmark
                         + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestData.cs" + Qmark
                         + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestUtility.cs" + Qmark;
                    arg = arg + space + "/reference:";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.Design.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.Interfaces.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.Security.dll" + '"' + ";";
                    arg = arg + '"' + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\Egemin.Epia.Testing.TestTools.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.Security.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.UI.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Definitions.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.WCS.dll" + '"' + ";";
                    //arg = arg + '"' + Path.Combine(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "Interop.VBIDE.dll") + '"' + ";";
                    arg = arg + '"' + Path.Combine(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "Microsoft.Office.Interop.Excel.dll") + '"' + ";";

                    if (!File.Exists(Path.Combine(destPath, "Microsoft.Office.Interop.Excel.dll")))
                        File.Copy(Path.Combine(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "Microsoft.Office.Interop.Excel.dll"),
                            Path.Combine(destPath, "Microsoft.Office.Interop.Excel.dll"), true);

                    //if (!File.Exists(Path.Combine(destPath, "Interop.VBIDE.dll")))
                    //	File.Copy(Path.Combine(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY, "Interop.VBIDE.dll"),
                    //		Path.Combine(destPath, "Interop.VBIDE.dll"), true);
                }
                else    //  test in development environment 
                {
                    arg = "/debug /target:library /out:" + Qmark + ConstCommon.ETRICC_TESTS_DIRECTORY + @"\TestRuns\bin\debug\TestRuns.dll" + Qmark;
                    arg = arg + space + Qmark + getRootPath() + @"TestRuns\TestWorker.cs" + Qmark
                         + space + Qmark + getRootPath() + @"TestRuns\TestProjectWorker.cs" + Qmark
                         + space + Qmark + getRootPath() + @"TestRuns\TestConstants.cs" + Qmark
                         + space + Qmark + getRootPath() + @"TestRuns\Logger.cs" + Qmark
                         + space + Qmark + getRootPath() + @"TestRuns\TestData.cs" + Qmark
                         + space + Qmark + getRootPath() + @"TestRuns\TestUtility.cs" + Qmark;
                    arg = arg + space + "/reference:";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.Design.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.Interfaces.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.Security.dll" + '"' + ";";
                    //arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.Security.SSPI.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.Security.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Common.UI.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.Definitions.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + @"\Egemin.EPIA.WCS.dll" + '"' + ";";
                    //arg = arg + '"' + m_CurrentDrive+ @"Epia 3\Testing\Automatic\AutomaticTests\Source\TestRuns\bin\Debug\Interop.Excel.dll" + '"' + ";";
                    arg = arg + '"' + m_CurrentDrive + @"Team Systems\Epia 3\Testing\Automatic\AutomaticTests\OEM\Microsoft\Microsoft.Office.Interop.Excel.dll" + '"' + ";";

                    if (!File.Exists(Path.Combine(destPath, "Microsoft.Office.Interop.Excel.dll")))
                        File.Copy(m_CurrentDrive + @"Team Systems\Epia 3\Testing\Automatic\AutomaticTests\OEM\Microsoft\Microsoft.Office.Interop.Excel.dll",
                            Path.Combine(destPath, "Microsoft.Office.Interop.Excel.dll"), true);

                    //if (!File.Exists(Path.Combine(destPath, "Interop.VBIDE.dll")))
                    //	File.Copy(m_CurrentDrive + @"Epia 3\Testing\Automatic\AutomaticTests\OEM\Microsoft\Interop.VBIDE.dll",
                    //		Path.Combine(destPath, "Interop.VBIDE.dll"), true);
                }

                Log(" arg :" + arg);
                logger.LogMessageToFile("  arg :" + arg, 0, 0);
                //System.Windows.Forms.MessageBox.Show("  arg :" + arg, "RecompileTestRuns");

                string DotnetVersionPath = m_SystemDrive + @"WINDOWS\Microsoft.NET\Framework\" + Constants.sRecompileDotnetVersion;
                string exePath = Path.Combine(DotnetVersionPath, "csc.exe");

                //System.Windows.Forms.MessageBox.Show("  exePath :" + exePath, "RecompileTestRuns");
                logger.LogMessageToFile("  exePath :" + exePath, 0, 0);

                // Run recompile Process
                //System.Windows.Forms.MessageBox.Show(arg, " arg");
                output = ProcessUtilities.RunProcessAndGetOutput(exePath, arg);
                if (output.IndexOf("error") >= 0)
                    throw new Exception(output);

                //System.Windows.Forms.MessageBox.Show("  output :" + output, "RecompileTestRuns");
                logger.LogMessageToFile("  output :" + output, 0, 0);
                Log("TestRun Recompiled ");

                Thread.Sleep(2000);
                #endregion
            }
            catch (Exception exRecomp)
            {
                string msg = exRecomp.ToString() + System.Environment.NewLine + exRecomp.StackTrace;
                System.Windows.Forms.MessageBox.Show(msg, "RecompileTestRuns");
                throw new Exception(msg + "  during <RecompileTestRuns>");
            }

            return output;
        }

        public DateTime GetDeploymentEndTime()
        {
            return sDeploymentEndTime;
        }

        public bool GetDeploymentStatus()
        {
            return sIsDeployed;
        }

        public void SetDeploymentStatus(bool status)
        {
            sIsDeployed = status;
        }

        private static void WriteInstallationToReg(string AppName, string FilePath, string MsiName)
        {
            logger.LogMessageToFile("------------WriteInstallationToReg: " + AppName + "InstallationPath"
                + System.Environment.NewLine + "FilePath: " + FilePath + " & " + "MsiName: " + MsiName, 0, 0);
            RegistryKey key = Registry.CurrentUser.CreateSubKey(REGKEY);
            key.SetValue(AppName + "InstallationPath", FilePath);
            key.SetValue(AppName + "InstallationName", MsiName);
        }

        static public void InstallCurrentApplicationSetup(string currentSetupPath, string EgeminApplicationApp, string buildnr, ref string errorMsg, Logger logger)
        {
            logger.LogMessageToFile(" *** start " + EgeminApplicationApp + " installing :" + buildnr, 0, 0);
            if (ProjAppInstall.InstallApplication(currentSetupPath, EgeminApplicationApp, EgeminApplication.SetupType.Default, ref errorMsg, logger))
            {
                //save the name and the path in the registry to remove the setup
                WriteInstallationToReg(Constants.ETRICCSTATISTICS_PARSER_SETUP, currentSetupPath, EgeminApplicationApp + ".msi");
                logger.LogMessageToFile(" **************** ( END " + " " + EgeminApplicationApp + ".msi" + "Deployment )*****************", 0, 0);
                logger.LogMessageToFile(EgeminApplicationApp + " installed :" + buildnr, 0, 0);
            }
            else
            {
                logger.LogMessageToFile(" *** " + EgeminApplicationApp + " installing  failed:" + errorMsg, 0, 0);
                throw new Exception(" *** " + EgeminApplicationApp + " installing  failed:" + errorMsg);
            }
        }

        public string getLogPath()
        {
            return sLogFilename;
        }

        internal void Log(string Message)
        {
            Message = System.String.Format("{0:G}: {1}.", System.DateTime.Now, " - " + Message);
            // set max log line is 100
            while (m_Logging.Count > 100)
            {
                m_Logging.RemoveAt(0);
            }

            m_Logging.Add(Message);
            if (OnLoggingChanged != null)
                OnLoggingChanged(this, new EventArgs());
        }

        public bool CatchExceptionMessage(ref string msg)
        {
            bool hasException = false;
            AutomationElement aeProjectForm = null;
            AutomationElement aeExceptionBox = null;
            AutomationElement aeExceptionTxt = null;
            string projectFormID = "frmMain";
            projectFormID = "frmScript";
            string exceptionBoxID = "ExceptionBoxForm";
            string exceptionTxtID = "txbException";
            DateTime StartTime = DateTime.Now;
            TimeSpan Time = DateTime.Now - StartTime;
            int wt = 0;
            while (aeProjectForm == null && wt < 5)
            {
                aeProjectForm = AUIUtilities.FindElementByID(projectFormID, AutomationElement.RootElement);
                Thread.Sleep(1000);
                Time = DateTime.Now - StartTime;
                wt = wt + 2;
            }

            if (aeProjectForm != null)
            {
                logger.LogMessageToFile("project window opened:" + aeProjectForm.Current.Name,
                    0, 0);

                wt = 0;
                aeExceptionBox = null;
                while (aeExceptionBox == null && wt < 5)
                {
                    aeExceptionBox = AUIUtilities.FindElementByID(exceptionBoxID, aeProjectForm);
                    Thread.Sleep(1000);
                    Time = DateTime.Now - StartTime;
                    wt = wt + 2;
                }

                if (aeExceptionBox != null)
                {
                    logger.LogMessageToFile("Exception Box found:" + aeProjectForm.Current.Name,
                        0, 0);

                    wt = 0;
                    aeExceptionTxt = null;
                    while (aeExceptionTxt == null && wt < 5)
                    {
                        aeExceptionTxt = AUIUtilities.FindElementByID(exceptionTxtID, aeExceptionBox);
                        Thread.Sleep(1000);
                        Time = DateTime.Now - StartTime;
                        wt = wt + 2;
                    }

                    if (aeExceptionTxt != null)
                    {
                        logger.LogMessageToFile("Exception Txt found:" + aeProjectForm.Current.Name,
                        0, 0);

                        TextPattern tp = (TextPattern)aeExceptionTxt.GetCurrentPattern(TextPattern.Pattern);
                        Thread.Sleep(1000);
                        msg = tp.DocumentRange.GetText(-1).Trim();
                        hasException = true;
                    }
                }
            }
            return hasException;
        }

        public void StartTestWorker(string setupPath, string projectPath, string projectName)
        {
            // startup Launcher
            //if (Configuration.Launcher)
            //{
            //MessageBox.Show( setupPath);
            DirectoryInfo workingpath = new DirectoryInfo(setupPath);
            //MessageBox.Show("path:" + workingpath.FullName);

            //MessageBox.Show(setupPath);
            System.Diagnostics.Process procLauncher = new System.Diagnostics.Process();
            procLauncher.EnableRaisingEvents = false;
            procLauncher.StartInfo.FileName = "Epia.Launcher.exe";
            procLauncher.StartInfo.Arguments = "/objecttype Egemin.EPIA.WCS.Core.Project  /uri gtcp://localhost:50000/Project /Startup Overview";
            procLauncher.StartInfo.WorkingDirectory = workingpath.FullName;
            procLauncher.Start();
            //}

            //System.Windows.Forms.MessageBox.Show("Shell started:");
            // load xml to Explorer and start up
            //if (Configuration.Explorer)
            //{
            //projectNa;estring projectPath = "C:\\Epia 3\\Testing\\Automatic\\AutomaticTests\\Data\\Xml";
            //string projectPath = "C:\\Epia 3\\Testing\\Automatic\\AutomaticTests\\TestData\\Etricc5"; //project

            string EtriccXmlPath = m_CurrentDrive + @"Team Systems\Epia 3\Testing\Automatic\AutomaticTests\TestData\Etricc5"; //script path
            if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))
                EtriccXmlPath = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY;

            //string EtriccXmlFilename = "TestEurobaltic.xml";
            string xmlfile = Path.Combine(projectPath, projectName);

            // look for activate script
            string script = "activate.vbs";
            StringCollection tokens = new StringCollection();
            tokens.AddRange(projectName.Split(new char[] { '.' }));
            script = tokens[0].Trim() + "Activate.vbs";

            System.Diagnostics.Process procExplorer = new System.Diagnostics.Process();
            procExplorer.EnableRaisingEvents = false;
            procExplorer.StartInfo.FileName = "Epia.Explorer.exe";
            string explorerInput = "";
            string path = '"' + xmlfile + '"';
            string objectType = "Egemin.EPIA.WCS.Core.Project";
            string remote = "/r";
            //string activatepath = '"' + Path.Combine(getRootPath(), script) + '"';
            string activatepath = '"' + Path.Combine(EtriccXmlPath, script) + '"';
            string blank = " ";

            explorerInput = path + blank + objectType;
            //if (Configuration.Remoting)
            explorerInput = explorerInput + blank + remote;

            //if (Configuration.Activate)
            explorerInput = explorerInput + blank + activatepath;

            //MessageBox.Show(explorerInput, "StartTestWorker");
            procExplorer.StartInfo.Arguments = explorerInput;
            //procExplorer.StartInfo.Arguments =xmlfile+"  Egemin.EPIA.WCS.Core.Project  /r "+activatepath;
            procExplorer.StartInfo.WorkingDirectory = workingpath.FullName;
            //procExplorer.StartInfo.WorkingDirectory = @"C:\Program Files\Egemin\Epia 1.9.12";

            logger.LogMessageToFile(" explorer args :" + explorerInput, 0, 0);
            logger.LogMessageToFile(" explorer path :" + setupPath, 0, 0);
            try
            {
                procExplorer.Start();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("explor start exception:" + ex.Message + "-" + ex.StackTrace);
            }


            //}
        }

        /// <summary>
        /// Copies the new setup to the local machine
        /// </summary>
        /// <returns></returns>
        public bool CopySetup(string fromPath, string toPath)
        {
            if (fromPath.StartsWith(@"\\"))
            {
                //if the first action fails try to logon to the server
                if (CreateDriveMap(fromPath) != 0)
                {
                    System.Windows.MessageBox.Show("CreateDriveMap   failed:" + fromPath);
                    return false;
                }
            }

            if (sMsgDebug.StartsWith("true"))
            {
                System.Windows.Forms.MessageBox.Show("from:" + fromPath, "CopySetup setup");
                System.Windows.Forms.MessageBox.Show("to:" + toPath, "CopySetup setup");
            }

            if (!Directory.Exists(fromPath))
            {
                Directory.CreateDirectory(fromPath);
            }


            if (Directory.Exists(toPath))
            {
                DirectoryInfo DirInfo = new DirectoryInfo(toPath);
                FileInfo[] FilesToDelete = DirInfo.GetFiles();

                foreach (FileInfo file in FilesToDelete)
                {
                    try
                    {
                        FileAttributes attributes = FileAttributes.Normal;
                        File.SetAttributes(file.FullName, attributes);
                        file.Delete();
                    }
                    catch (Exception ex)
                    {
                        //if (m_Settings.EnableLog)
                        //{
                        //string logPath = Configuration.BuildInformationfilePath;
                        //string path = Path.Combine( logPath, Configuration.LogFilename );
                        //Logger logger = new Logger(path );
                        logger.LogMessageToFile("----------Setup Error  --------", 0, 0);
                        logger.LogMessageToFile("CopySetup Exception:" + ex.ToString(), 0, 0);
                        Log("CopySetup Exception:" + ex.ToString());
                        //}
                        System.Windows.MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
                        m_State = Tester.STATE.EXCEPTION;
                        return false;
                    }
                }
            }
            else
                Directory.CreateDirectory(toPath);

            bool copyStatus = false;
            while (copyStatus == false)
            {
                FileInfo[] FilesToCopy;
                try
                {
                    DirectoryInfo DirInfo = new DirectoryInfo(fromPath);
                    FilesToCopy = DirInfo.GetFiles();

                    foreach (FileInfo file in FilesToCopy)
                    {
                        file.CopyTo(Path.Combine(toPath, file.Name));
                    }
                    Log("Copied Setup from " + fromPath + " to " + toPath);
                    logger.LogMessageToFile("Copied Setup from " + fromPath + " to " + toPath, 0, 0);
                    copyStatus = true;
                }
                catch (Exception ex)
                {
                    copyStatus = false;
                    logger.LogMessageToFile("----------Setup Error --------", 0, 0);
                    logger.LogMessageToFile("CopySetup Exception:" + ex.ToString(), 0, 0);
                    Log("CopySetup Exception:" + ex.ToString());
                    MessageBoxEx.Show("FromPath=" + fromPath + "   " + ex.ToString() + "\r\n" + ex.StackTrace, 5 * 60000);
                }
            }
            return true;
        }

        /// <summary>
        /// Open a drive mapping to the destination
        /// </summary>
        /// <param name="Destination">Full drive path</param>
        private static int OpenDriveMap(string Destination, string driveLetter)
        {
            if ((Destination == null) || (Destination == ""))
                return -1;

            NETRESOURCEA[] netResource = new NETRESOURCEA[1];
            netResource[0] = new NETRESOURCEA();
            netResource[0].dwType = 1;
            netResource[0].lpLocalName = driveLetter;
            netResource[0].lpRemoteName = Destination;
            netResource[0].lpProvider = null;
            int dwFlags = 1; /*CONNECT_INTERACTIVE = 8|CONNECT_PROMPT = 16*/

            int result = WNetAddConnection2A(netResource, null, null, dwFlags);
            //int result = WNetAddConnection2A(netResource, null, @"teamsystems\tfstest", dwFlags);

            return result;
        }

        public bool CopyEtriccSetupFilesWithWildcards(string fromPath, string toPath, string filenameWithWildcards)
        {
            if (fromPath.StartsWith(@"\\"))
            {
                //if the first action fails try to logon to the server
                if (CreateDriveMap(fromPath) != 0)
                {
                    System.Windows.MessageBox.Show("CreateDriveMap2   failed:" + fromPath);
                    return false;
                }
            }

            if (!Directory.Exists(fromPath))
            {
                Directory.CreateDirectory(fromPath);
            }


            if (Directory.Exists(toPath))
            {
                DirectoryInfo DirInfo = new DirectoryInfo(toPath);
                FileInfo[] FilesToDelete = DirInfo.GetFiles();

                foreach (FileInfo file in FilesToDelete)
                {
                    try
                    {
                        FileAttributes attributes = FileAttributes.Normal;
                        File.SetAttributes(file.FullName, attributes);
                        if (file.FullName.IndexOf(filenameWithWildcards.Substring(2)) >= 0)
                            file.Delete();
                    }
                    catch (Exception ex)
                    {
                        //if (m_Settings.EnableLog)
                        //{
                        //string logPath = Configuration.BuildInformationfilePath;
                        //string path = Path.Combine( logPath, Configuration.LogFilename );
                        //Logger logger = new Logger(path );
                        logger.LogMessageToFile("----------Setup Error  --------", 0, 0);
                        logger.LogMessageToFile("CopySetup Exception:" + ex.ToString(), 0, 0);
                        Log("CopySetup Exception:" + ex.ToString());
                        //}
                        System.Windows.MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
                        m_State = Tester.STATE.EXCEPTION;
                        return false;
                    }
                }
            }
            else
                Directory.CreateDirectory(toPath);

            FileInfo[] FilesToCopy;
            try
            {
                DirectoryInfo DirInfo = new DirectoryInfo(fromPath);
                FilesToCopy = DirInfo.GetFiles(filenameWithWildcards);

                foreach (FileInfo file in FilesToCopy)
                {
                    file.CopyTo(Path.Combine(toPath, file.Name));
                }
                Log("Copied Setup from " + fromPath + " to " + toPath);
                try
                {
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    logger.LogMessageToFile("Copied Setup from " + fromPath + " to " + toPath, 0, 0);
                    //}
                }
                catch (Exception ex1)
                {
                    //if (m_Settings.EnableLog)
                    logger.LogMessageToFile("------ Test Exception : " + ex1.Message + "\r\n" + ex1.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                }
            }
            catch (Exception ex)
            {
                try
                {
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    logger.LogMessageToFile("----------Setup Error --------", 0, 0);
                    logger.LogMessageToFile("CopySetup Exception:" + ex.ToString(), 0, 0);
                    Log("CopySetup Exception:" + ex.ToString());
                    //}
                }
                catch (Exception ex2)
                {
                    //if (m_Settings.EnableLog)
                    logger.LogMessageToFile("------ Test Exception : " + ex2.Message + "\r\n" + ex2.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                }
                System.Windows.MessageBox.Show("FromPath=" + fromPath + "   " + ex.ToString() + "\r\n" + ex.StackTrace);
                m_State = Tester.STATE.EXCEPTION;
                return false;
            }
            return true;
        }

        public bool CopySetupFilesWithWildcards(string fromPath, string toPath, string filenameWithWildcards)
        {
            if (fromPath.StartsWith(@"\\"))
            {
                //if the first action fails try to logon to the server
                if (CreateDriveMap(fromPath) != 0)
                {
                    System.Windows.MessageBox.Show("CreateDriveMap2   failed:" + fromPath);
                    return false;
                }
            }

            if (!Directory.Exists(fromPath))
            {
                Directory.CreateDirectory(fromPath);
            }


            if (Directory.Exists(toPath))
            {
                DirectoryInfo DirInfo = new DirectoryInfo(toPath);
                FileInfo[] FilesToDelete = DirInfo.GetFiles();

                bool deleted = false;
                foreach (FileInfo file in FilesToDelete)
                {
                    // delete file first
                    deleted = false;
                    while (deleted == false)
                    {
                        try
                        {
                            FileAttributes attributes = FileAttributes.Normal;
                            File.SetAttributes(file.FullName, attributes);
                            file.Delete();
                            deleted = true;
                        }
                        catch (System.IO.IOException excep)
                        {
                            if (excep.ToString().IndexOf("is being used by another process") >= 0)
                            {
                                logger.LogMessageToFile(file.Name + " is being used by another process, cannot be deleted now, try again", 0, 0);
                                logger.LogMessageToFile("CopySetup Exception:" + excep.ToString(), 0, 0);
                                deleted = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.LogMessageToFile("----------Setup Error  --------", 0, 0);
                            logger.LogMessageToFile("CopySetup Exception:" + ex.ToString(), 0, 0);
                            Log("CopySetup Exception:" + ex.ToString());
                            //}
                            System.Windows.MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
                            m_State = Tester.STATE.EXCEPTION;
                            deleted = false;
                            return false;
                        }
                    } // end while
                }
            }
            else
                Directory.CreateDirectory(toPath);

            FileInfo[] FilesToCopy;
            try
            {
                DirectoryInfo DirInfo = new DirectoryInfo(fromPath);
                FilesToCopy = DirInfo.GetFiles(filenameWithWildcards);

                foreach (FileInfo file in FilesToCopy)
                {
                    file.CopyTo(Path.Combine(toPath, file.Name));
                }
                Log("Copied Setup from " + fromPath + " to " + toPath);
                try
                {
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    logger.LogMessageToFile("Copied Setup from " + fromPath + " to " + toPath, 0, 0);
                    //}
                }
                catch (Exception ex1)
                {
                    //if (m_Settings.EnableLog)
                    logger.LogMessageToFile("------ Test Exception : " + ex1.Message + "\r\n" + ex1.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                }
            }
            catch (Exception ex)
            {
                try
                {
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    logger.LogMessageToFile("----------Setup Error --------", 0, 0);
                    logger.LogMessageToFile("CopySetup Exception:" + ex.ToString(), 0, 0);
                    Log("CopySetup Exception:" + ex.ToString());
                    //}
                }
                catch (Exception ex2)
                {
                    //if (m_Settings.EnableLog)
                    logger.LogMessageToFile("------ Test Exception : " + ex2.Message + "\r\n" + ex2.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                }
                System.Windows.MessageBox.Show("FromPath=" + fromPath + "   " + ex.ToString() + "\r\n" + ex.StackTrace);
                m_State = Tester.STATE.EXCEPTION;
                return false;
            }
            return true;
        }

        /// <summary>
        /// Create a drive mapping to the destination
        /// </summary>
        /// <param name="Destination">Full drive path</param>
        private static int CreateDriveMap(string Destination)
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
            int result = WNetAddConnection2A(netResource, null, null, dwFlags);
            System.Windows.Forms.MessageBox.Show("Destination=" + Destination, "Create Map result:" + result);
            return result;
        }

        /// <summary>
        /// Cancel a drive mapping to the destination
        /// </summary>
        /// <param name="Destination">Full drive path</param>
        public static int Disconnect(string localpath)
        {
            int result = WNetCancelConnection2A(localpath, 1, 1);
            return result;
        }

        // only used in in Development to get Build base directory
        private string getRootPath()
        {
            string root = System.Windows.Forms.Application.StartupPath;

            StringCollection tokens = new StringCollection();
            tokens.AddRange(root.Split(new char[] { '\\' }));
            //Remove last three tokens and form a root path
            string rootPath = "";
            for (int i = 0; i < tokens.Count - 3; i++)
            {
                tokens[i] = tokens[i].Trim();
                rootPath = rootPath + tokens[i] + "\\";
            }
            return rootPath;
        }

        #region // —— UI Help Methods •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
        public void ClickUiScreenActionToAvoidScreenStandBy()
        {
            System.Windows.Point point = new Point(1, 1);
            TestTools.Input.MoveToAndRightClick(point);
            Thread.Sleep(2000);
            point.Y = point.Y + 300;
            TestTools.Input.MoveToAndClick(point);
            Thread.Sleep(2000);
        }
        #endregion

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Properties of Tester (2)
        public StringCollection Logging
        {
            get
            {
                return m_Logging;
            }
        }

        public STATE State
        {
            get
            {
                return m_State;
            }
            set
            {
                m_State = value;
            }
        }
        #endregion // —— Properties •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


        #region OnEvent ------------------------------------------------------------------------------------------------
        public AutomationElement WaitUntilMyButtonFoundInThisWindow(string WindowName, string ButtonName, int searchCnt)
        {
            AutomationElement aeMyButton = null;
            bool myButtonFound = false;
            DateTime foundStartTime = DateTime.Now;
            TimeSpan foundTime = DateTime.Now - foundStartTime;
            int foundCnt = 0;

            while (myButtonFound == false && foundCnt <= searchCnt)
            {
                //find install Window screen
                System.Windows.Automation.Condition c2p = new AndCondition(
                  new PropertyCondition(AutomationElement.NameProperty, WindowName),
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );

                AutomationElement aeWindow = null;
                DateTime StartTime = DateTime.Now;
                TimeSpan Time = DateTime.Now - StartTime;

                while (aeWindow == null && Time.TotalSeconds <= 600)
                {
                    aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c2p);
                    Thread.Sleep(2000);
                    Time = DateTime.Now - StartTime;
                    logger.LogMessageToFile("<-----> find Window: " + Time.Seconds, sLogCount, sLogInterval);
                }

                Thread.Sleep(3000);

                if (aeWindow == null)
                {
                    System.Windows.Forms.MessageBox.Show("find button in " + ButtonName + " <-----> Window not found: windows name is: " + WindowName, "WaitUntilMyButtonFoundInThisWindow:" + searchCnt);
                }
                else
                {
                    aeMyButton = FindSpecificButtonByName(aeWindow, ButtonName);

                    if (aeMyButton == null)
                    {
                        foundCnt++;
                        logger.LogMessageToFile(foundCnt + " time search <-----> My button not found and try again: " + ButtonName, sLogCount, sLogInterval);
                        continue;
                    }
                    else
                    {
                        logger.LogMessageToFile(foundCnt + " time search <====> My button found: " + ButtonName, sLogCount, sLogInterval);
                        myButtonFound = true;
                    }
                }
            }

            return aeMyButton;

        }

        public AutomationElement WaitUntilMyButtonFoundInThisWindowWithStatusEnable(string WindowName, string ButtonName, int searchCnt)
        {
            AutomationElement aeMyButton = null;
            bool myButtonFound = false;
            DateTime foundStartTime = DateTime.Now;
            TimeSpan foundTime = DateTime.Now - foundStartTime;
            int foundCnt = 0;

            while (myButtonFound == false && foundCnt <= searchCnt)
            {
                //find install Window screen
                System.Windows.Automation.Condition c2p = new AndCondition(
                  new PropertyCondition(AutomationElement.NameProperty, WindowName),
                  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                );

                AutomationElement aeWindow = null;
                DateTime StartTime = DateTime.Now;
                TimeSpan Time = DateTime.Now - StartTime;

                while (aeWindow == null && Time.Seconds <= 600)
                {
                    aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c2p);
                    Thread.Sleep(2000);
                    Time = DateTime.Now - StartTime;
                    logger.LogMessageToFile("<-----> find Window: " + Time.Seconds, sLogCount, sLogInterval);
                }

                Thread.Sleep(3000);

                if (aeWindow == null)
                {
                    System.Windows.Forms.MessageBox.Show("find button in " + ButtonName + " <-----> Window not found: windows name is: " + WindowName, "WaitUntilMyButtonFoundInThisWindowWithStatusEnable");
                }
                else
                {
                    aeMyButton = FindSpecificButtonByName(aeWindow, ButtonName);

                    if (aeMyButton == null)
                    {
                        foundCnt++;
                        logger.LogMessageToFile(foundCnt + " time search <-----> My button not found and try again: " + ButtonName, sLogCount, sLogInterval);
                        continue;
                    }
                    else
                    {
                        logger.LogMessageToFile(foundCnt + " time search <====> My button found: " + ButtonName, sLogCount, sLogInterval);
                        if (aeMyButton.Current.IsEnabled)
                        {
                            myButtonFound = true;
                        }
                        else
                        {
                            logger.LogMessageToFile(foundCnt + " time search <====> My button found , but not enable: " + ButtonName, sLogCount, sLogInterval);
                            continue;
                        }
                    }
                }
            }
            return aeMyButton;

        }

        public AutomationElement FindSpecificButtonByName(AutomationElement aeWindow, string buttonName)
        {
            logger.LogMessageToFile("<-----> FindSpecificButtonByName: " + buttonName, sLogCount, sLogInterval);
            // Set a property condition that will be used to find the control.
            System.Windows.Automation.Condition c = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Button);

            AutomationElementCollection aeAllButtons = aeWindow.FindAll(TreeScope.Element | TreeScope.Descendants, c);
            Thread.Sleep(1000);
            // get Next > button
            // get unexamined build quality
            AutomationElement aeButton = null;
            if (aeAllButtons == null || aeAllButtons.Count == 0)
            {
                aeButton = null;
                logger.LogMessageToFile("<-----> FindSpecificButtonByName: no buttons found ", sLogCount, sLogInterval);
            }
            else
            {
                foreach (AutomationElement s in aeAllButtons)
                {
                    logger.LogMessageToFile("<-----> list all buttons: " + s.Current.Name + " and isContentElement:" + s.Current.IsContentElement, sLogCount, sLogInterval);
                    if (s.Current.Name.Equals(buttonName) && s.Current.IsContentElement)
                    {
                        aeButton = s;
                    }
                }
            }

            if (aeButton == null)
                logger.LogMessageToFile("<-----> button not found: " + buttonName, sLogCount, sLogInterval);

            return aeButton;
        }

        public bool WindowHasOnlyThisButton(AutomationElement aeWindow, string buttonName)
        {
            bool status = false;
            logger.LogMessageToFile("<-----> HaveOnlyThisButton? : " + buttonName, sLogCount, sLogInterval);
            // Set a property condition that will be used to find the control.
            System.Windows.Automation.Condition c = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Button);

            AutomationElementCollection aeAllButtons = aeWindow.FindAll(TreeScope.Element | TreeScope.Descendants, c);
            if (aeAllButtons.Count == 1)
            {
                if (aeAllButtons[0].Current.Name.Equals(buttonName))
                {
                    status = true;
                }
            }
            return status;
        }
        //----------------------------------------------------------------------------------------------------------------------------
        private void StartAddRemoveProgramExecution()
        {
            System.Diagnostics.Process Proc = new System.Diagnostics.Process();
            Proc.StartInfo.FileName = @"C:\Windows\System32\appwiz.cpl";
            Proc.StartInfo.CreateNoWindow = false;
            Proc.Start();
        }

        public static bool IsApplicationInstalled(string ApplicationType, string uninstallWindowName)
        {
            bool applicationInstalled = false;

            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName);
            AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
            if (appElement != null)
            {   // (1) Programs and Features main window
                Console.WriteLine("Programs and Features main window opend");
                Thread.Sleep(1000);
                Console.WriteLine("Searching programs item button...");
                AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(appElement, "Folder View");
                if (aeGridView != null)
                    Console.WriteLine("Gridview found...");
                Thread.Sleep(1000);
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.DataItem);

                AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children, c);
                Console.WriteLine("Programs count ..." + aeProgram.Count);
                for (int i = 0; i < aeProgram.Count; i++)
                {
                    switch (ApplicationType)
                    {
                        case "Epia":
                            if (aeProgram[i].Current.Name.StartsWith("E'pia Fr"))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "EtriccCore":
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                    && aeProgram[i].Current.Name.IndexOf("Shell") < 0
                                 && aeProgram[i].Current.Name.IndexOf("Playback") < 0
                                   && aeProgram[i].Current.Name.IndexOf("Statistics") < 0
                                && aeProgram[i].Current.Name.IndexOf("HostTest") < 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "EtriccShell":
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                    && aeProgram[i].Current.Name.IndexOf("Shell") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "Etricc.Statistics.Parser":
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                    && aeProgram[i].Current.Name.IndexOf("Parser ") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "Etricc.Statistics.ParserConfigurator":
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                    && aeProgram[i].Current.Name.IndexOf("ParserConfigurator") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "Etricc.Statistics.UI":
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                    && aeProgram[i].Current.Name.IndexOf("UI") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "Ewcs":
                            if (aeProgram[i].Current.Name.StartsWith("E'wcs")
                                && aeProgram[i].Current.Name.IndexOf("TestPrograms") < 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "EwcsTestProgram":
                            if (aeProgram[i].Current.Name.StartsWith("E'wcs")
                                && aeProgram[i].Current.Name.IndexOf("TestPrograms") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "AutomaticTesting":
                            if (aeProgram[i].Current.Name.StartsWith("Automatic"))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                    }
                }
            }

            return applicationInstalled;
        }

        #endregion

        public string GetCurrentBuildInTesting()
        {
            return sCurrentBuildInTesting;
        }

        public void LoadTestConfigSectionSettings()
        {
            string sectionName = Constants.TestConfigSection;
            System.Configuration.Configuration config =
                  System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            m_TestConfigSettings = (TestsConfigSection)config.Sections[sectionName];
            if (m_TestConfigSettings == null)
                throw new Exception("Load Test Config Section Settings failed");

            Log("Test Config Section Settings Loaded");
        }

        public void LoadTfsSettingsSection()
        {
            string sectionName = Constants.TfsSettingsSection;
            System.Configuration.Configuration config =
                  System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            m_TFS_Settings = (TfsSettingsSection)config.Sections[sectionName];
            if (m_TFS_Settings == null)
                throw new Exception("Load TFS Settings Section failed");

            Log("TFS Settings Section Loaded");
        }

        public void SetTestAutoMode(bool value)
        {
            m_TestAutoMode = value;
        }

        public static void SetErrorMessage(string value)
        {
            sErrorMessage = value;
        }

        public static string GetErrorMessage()
        {
            return sErrorMessage;
        }

        public enum STATE
        {
            UNDEFINED,
            PENDING,
            INPROGRESS,
            EXCEPTION,
        }
    }
}
