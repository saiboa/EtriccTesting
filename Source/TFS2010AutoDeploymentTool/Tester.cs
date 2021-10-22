#region Using directives
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using ICSharpCode.SharpZipLib.Zip;

using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

using Microsoft.Win32;
using TestTools;
#endregion

namespace TFS2010AutoDeploymentTool
{
    public class Tester
    {
        #region fields
        internal static TestTools.Logger logger = null;
        internal static string sLogFilename = string.Empty;

        // GUI test params
        static string m_installScriptDir = string.Empty;                        // "0) Install msi file path: "
        static string m_ValidatedBuildDropFolder = string.Empty;                // "1) Build Drop folder: "
        static string m_BuildNumber = string.Empty;                             // "2) build nr: "
        static string mTeamProject = string.Empty;                              // "3) test Project: "
        static string mTestApp = string.Empty;                                  // "4) test Application: "
        static string mTargetPlatform = string.Empty;                           // "5) targeted platform: " --->  AnyCPU x86 X64    AnyCPU+X86
        static string mCurrentPlatform = string.Empty;                          // "6) current platform: "  --->  AnyCPU+X86   --> AnyCPU  than x86
        static string mTestDef = string.Empty;                                  // "7) test def: " -->  CI, Nightly, Weekly, Version
        const string mCalledProgram = "TFS2010AutoDeploymentTool";              // "8) Called by: "
        internal string TESTTOOL_VERSION = "3.10.08.05";                        // "9) TestTool version: "
        static internal bool m_TestAutoMode = true;                             // "10) Auto test: "
        //static string sTFSServerUrl = Constants.sTFSServerUrl;                // "11) TFSServerUrl: "
        string mServerRunAs = "Service";                                        // "12) Server Run As: "  --> read from configuration file
        string mExcelVisible = "Visible";                                       // "13) Excel Visible: "  --> read from configuration file
        static string sDemonstration = Constants.sDemonstration;                // "14) Demo test: "         --> read from App.config file
        static string mMail = "false";                                          // "15) Mail: "         --> read from configuration file

        static bool mInstallOldEtricc5Service = false; 
        static bool mInstallEtriccLauncher = false;
        static string sEtriccMsiName = "Etricc ?.msi";
        

        List<string> mBuildDefs = new List<string>();
        static string mDateFilter = "Today";
        static bool mProtected = true;
        // --- TEST PARAMS
        //private static Settings m_Settings;
        private static TestsConfigSection m_TestConfigSettings = new TestsConfigSection();
        private static TfsSettingsSection m_TFS_Settings = new TfsSettingsSection();
        private static string sCurrentBuildInTesting = "Searching ......";      
        static internal string m_TestPC = string.Empty;
        static string m_CurrentDrive = string.Empty;
        static string m_SystemDrive = string.Empty;
        public DateTime sTestStartUpTime;
        internal string m_TestWorkingDirectory = string.Empty;

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

        static string sTestResultFolder = string.Empty;      
        // tested build info
        BuildObject m_ValidatedTestBuildObject = new BuildObject();
        private Uri m_Uri = null;
       
        // --- Etricc 5
        static string mPreviousSetupPathEpia = string.Empty;
        static string mCurrentSetupPathEpia = string.Empty;
        static string mPreviousSetupPathEtricc = string.Empty;
        static string mCurrentSetupPathEtricc = string.Empty;
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
        static bool sEventEnd = false;

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

            // prepare test directory
            m_CurrentDrive = Path.GetPathRoot(Directory.GetCurrentDirectory());
            string windir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
            m_SystemDrive = Path.GetPathRoot(windir);

            // Epia
            mPreviousSetupPathEpia = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Epia\\Previous";
            mCurrentSetupPathEpia = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Epia\\Current";

            mPreviousSetupPathEtricc = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Etricc\\Previous";
            mCurrentSetupPathEtricc = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Etricc\\Current";

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



            mTestRunsDirectory = ConstCommon.ETRICC_TESTS_DIRECTORY + "\\TestRuns\\bin\\Debug";

            if (!Directory.Exists(mPreviousSetupPathEpia))
                Directory.CreateDirectory(mPreviousSetupPathEpia);

            if (!Directory.Exists(mCurrentSetupPathEpia))
                Directory.CreateDirectory(mCurrentSetupPathEpia);
            //-------------ETRICC SHELL-----------------------------
            if (!Directory.Exists(mPreviousSetupPathEtricc))
                Directory.CreateDirectory(mPreviousSetupPathEtricc);

            if (!Directory.Exists(mCurrentSetupPathEtricc))
                Directory.CreateDirectory(mCurrentSetupPathEtricc);
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

            // Get tfs server 
            try
            {
                sIsDeployed = false;
                sDeploymentEndTime = DateTime.Now;
                m_TestPC = System.Environment.MachineName;

                if (TFSConnected)
                {
                    Log("Connect to TFS");

                    Uri serverUri = new Uri(Constants.sTFSServerUrl);
                    System.Net.ICredentials tfsCredentials
                        = new System.Net.NetworkCredential(Constants.sTFSUsername, Constants.sTFSPassword, Constants.sTFSDomain);

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
		#endregion // —— Constructors/Destructors/Cleanup •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        /// <summary>
        /// Method will start new tests
        /// </summary>
        public void Start(ref DateTime StartUpTime)
        {
            Start2(string.Empty, ref StartUpTime );
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
            mTeamProject = m_TFS_Settings.Element.TestProject;
            mTestApp = m_TFS_Settings.Element.TestApp;
            mTargetPlatform = m_TFS_Settings.Element.TargetPlatform;

            mBuildDefs.Clear();;
            string buildDefs = m_TFS_Settings.Element.BuildDefinitions;
            string[] strArray = buildDefs.Split(';');
            for (int i = 0; i < strArray.Length; i++)
            {    //;Epia.Development.Dev02.CI; in this case buildDefs.count = 3, two of them are empty, should check length > 0
                if (strArray[i].Length > 0)
                    mBuildDefs.Insert(i, strArray[i]);
            }

            mDateFilter = m_TFS_Settings.Element.DateFilter;
            mProtected = m_TFS_Settings.Element.BuildProtected;

            TimeSpan mTime = DateTime.Now - upTime;

            if (mTime.Minutes > 30)
                sLogInterval = 4;  // 1 min
            else if (mTime.Hours > 1)
                sLogInterval = 12; // 3 min
            else if (mTime.Hours > 5)
                sLogInterval = 20; // 5 min
            else if (mTime.Days > 1)
                sLogInterval = 40;  // 10 min

            string logFilename = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "-"
                + Constants.sDeploymentLogFilename;
            if (!logFilename.Equals(logger.GetLogPath()))   // if another day, create new log file
                logger = new Logger(System.IO.Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, logFilename));

            logger.LogMessageToFile("Start:" + sLogCount, sLogCount, sLogInterval);
            Log("===> searching for available build...");

            m_State = STATE.INPROGRESS;

            #region // Get All Buildnrs and DropLocationPaths From TFS depend on the setting from TFSConnection Setting
            if (manualBuildInfo == string.Empty)
            {   // should check build quality
                // Get All Build Objects of this project and build definitions
                //System.Windows.Forms.MessageBox.Show("mBuildDefs.count:" + mBuildDefs.Count);
                //System.Windows.Forms.MessageBox.Show("mDateFilter:" + mDateFilter);
               
                List<BuildObject> allBuilds = BuildUtilities.GetAllBuildObjects(mBuildDefs, mTeamProject, mDateFilter);
                if (allBuilds.Count == 0)
                {
                    Log("No Any build dircetory found:");
                    logger.LogMessageToFile("No Any build dircetory found:", sLogCount, sLogInterval);
                    ClickUiScreenActionToAvoidScreenStandBy();
                    sLogCount++;
                    m_State = STATE.PENDING;
                    return;
                }
                else
                {
                    // log all found build
                    string br = string.Empty;
                    logger.LogMessageToFile("(build dircetory found):", sLogCount, sLogInterval);
                    IEnumerator EmpEnumerator = allBuilds.GetEnumerator(); //Getting the Enumerator
                    EmpEnumerator.Reset(); //Position at the Beginning
                    while (EmpEnumerator.MoveNext()) //Till not finished do print
                    {
                        BuildObject b = (BuildObject)EmpEnumerator.Current;
                        Log(b.BuildNr + "\t" + b.DripLoc);
                        logger.LogMessageToFile(b.BuildNr + "\t" + b.DripLoc, sLogCount, sLogInterval);
                        br = br + "\n" + b.BuildNr +"\t" + b.FinishTime;
                    }
                    //System.Windows.Forms.MessageBox.Show("result:"+br);
                }

                // Get Validated build , that can be tested by thisPC
                // X:\Nightly\Etricc 5\Etricc - Nightly_20100202.1
                //MessageBox.Show("m_Settings.BuildApplication" + m_Settings.BuildApplication);
                // get one application of this build if not yet tested
                m_ValidatedTestBuildObject = GetValidatedBuildDirectory(allBuilds, m_TestPC, mTestApp, mTargetPlatform, mProtected, ref mCurrentPlatform, ref sMsiRelativePath);
                if (m_ValidatedTestBuildObject == null)
                {
                    logger.LogMessageToFile("No new validated build found:", sLogCount, sLogInterval);
                    ClickUiScreenActionToAvoidScreenStandBy();
                    m_State = STATE.PENDING;
                    sLogCount++;
                    return;
                }

                logger.LogMessageToFile("(validated build dircetory):", sLogCount, sLogInterval);

                m_BuildNumber = m_ValidatedTestBuildObject.BuildNr;
               
                m_ValidatedBuildDropFolder = m_ValidatedTestBuildObject.DripLoc;

                int ret = Disconnect(ConstCommon.DRIVE_MAP_LETTER);
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

                // will be optimalised later
                string driveMap = m_ValidatedTestBuildObject.xMapString;
                // @"\\Teamsystem.Teamsystems.egemin.be\Team Systems Builds"
                ret = OpenDriveMap(@driveMap, ConstCommon.DRIVE_MAP_LETTER);
                if (ret == 0)
                {
                    logger.LogMessageToFile(m_TestPC + "Open MAP DRIVE OK:", sLogCount, sLogInterval);
                }
                else if (ret == 85)
                    logger.LogMessageToFile(m_TestPC + "Open MAP DRIVE not connected due to existing connection:", sLogCount, sLogInterval);
                else
                    System.Windows.MessageBox.Show("OpenDriveMap failed with error code:" + ret);
                
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
                
                // Tested build type. Nightlym CI Version
                mTestDef = BuildUtilities.getTestDefinition(m_ValidatedBuildDropFolder);
                Log("testing definition : " + mTestDef);
                logger.LogMessageToFile("===== <testing  Definition > mTestDef=: " + mTestDef, 0, 0);

                if (sMsgDebug.StartsWith("true"))
                    System.Windows.Forms.MessageBox.Show("ValidatedBuildDirectory:" + m_ValidatedBuildDropFolder + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);
                
                m_TestAutoMode = true;
            }
            else
            {
                Log(" manual test starting : " + manualBuildInfo);
                #region
                //MessageBox.Show("BuildPath:" + BuildPath);
                m_BuildNumber = manualBuildInfo;
                MessageBox.Show("m_BuildNumber:" + m_BuildNumber);

                // get build number --> droploc
                Uri serverUri = new Uri(Constants.sTFSServerUrl);
                System.Net.ICredentials tfsCredentials
                = new System.Net.NetworkCredential(Constants.sTFSUsername, Constants.sTFSPassword, Constants.sTFSDomain);
                TfsTeamProjectCollection tfsProjectCollection = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                tfsProjectCollection.EnsureAuthenticated();
                IBuildServer buildServer = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));

                Uri buildUri = null;
                // create spec instance and apply the filter like build name, date time
                IBuildDetailSpec spec = buildServer.CreateBuildDetailSpec(mTeamProject);
                spec.BuildNumber = manualBuildInfo; //Example – “Daily_20110502.4″;
                IBuildQueryResult buildDetails = buildServer.QueryBuilds(spec);
                if (buildDetails != null)
                    buildUri = (buildDetails.Builds[0]).Uri;

                string dropLocation = (buildDetails.Builds[0]).DropLocation;
                m_ValidatedBuildDropFolder = ConstCommon.DRIVE_MAP_LETTER + "\\" + dropLocation.Substring(55);
                MessageBox.Show("m_ValidatedBuildDropFolder:" + m_ValidatedBuildDropFolder);

                mTestDef = BuildUtilities.getTestDefinition(m_ValidatedBuildDropFolder);
                MessageBox.Show("mTestDef:" + mTestDef);

                string platform = "Any CPU";
                if (mTargetPlatform.Equals("x86"))
                    platform = "x86";
                else if (mTargetPlatform.Equals("AnyCPU"))
                    platform = "Any CPU";
                else
                {
                    System.Windows.Forms.MessageBox.Show(mTargetPlatform+" platform is not allowed for manual testing, please select other platform");
                    return;
                }

                m_installScriptDir = m_ValidatedBuildDropFolder + "\\\\Installation\\\\" + platform + "\\\\Debug\\\\";
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
                    MessageBox.Show("BuildUtilities.GetProjectName(mTestApp)=" + BuildUtilities.GetProjectName(mTestApp));
                }
                //m_Uri = m_buildStore.GetBuildUri(ProjectName(m_testApp), m_BuildNumber);
                //MessageBox.Show("m_BuildSvc" + m_BuildSvc.GetBuildDefinition(GetProjectName(testApp),"").Id);
                m_Uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, BuildUtilities.GetProjectName(mTestApp), m_BuildNumber);
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
            // get m_installScriptDir and then check msi file exist, if not exist --> just return and do nothing
            if (mTestApp.Equals(Constants.ETRICC5))
            {
                if (m_TestAutoMode)
                    m_installScriptDir = m_ValidatedBuildDropFolder + sMsiRelativePath;
                Utilities.CloseProcess("EPIA.Launcher");
                Utilities.CloseProcess("EPIA.Explorer");
            }
            else if (mTestApp.Equals(Constants.ETRICCUI))
            {
                if (m_TestAutoMode)
                    m_installScriptDir = m_ValidatedBuildDropFolder + sMsiRelativePath;

                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_TOOLS_DATABASE_FILLER);

                #region check epia.msi exist?
                string EpiaMsiFile = System.IO.Path.Combine(m_installScriptDir, Constants.EPIA_MSI);
                if (!System.IO.File.Exists(EpiaMsiFile))
                {
                    if (m_TestAutoMode)
                    {
                        if (mTestApp.Equals(Constants.EPIA4) || mTestApp.Equals(Constants.ETRICCUI))
                        {
                            if (m_installScriptDir.IndexOf("Protected") > 0)
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no epia msi file found", mTestApp + "+" + mCurrentPlatform + "Protected");
                            else
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no epia msi file found", mTestApp + "+" + mCurrentPlatform + "Normal");
                        }
                        else
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no epia msi file found", mTestApp + "+" + mCurrentPlatform);
                        logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : m_testApp " + mTestApp, 0, 0);

                        FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                    }
                    return;
                }
                #endregion


            }
            else if (mTestApp.Equals(Constants.EPIA4))
            {
                if (m_TestAutoMode)
                    m_installScriptDir = m_ValidatedBuildDropFolder + sMsiRelativePath;

                if (sMsgDebug.StartsWith("true"))
                    System.Windows.MessageBox.Show("m_installScriptDir...   " + m_installScriptDir);
                
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_TOOLS_DATABASE_FILLER);

                #region check epia.msi exist?
                string EpiaMsiFile = System.IO.Path.Combine(m_installScriptDir, Constants.EPIA_MSI);
                if (!System.IO.File.Exists(EpiaMsiFile))
                {
                    if (m_TestAutoMode)
                    {
                        if (mTestApp.Equals(Constants.EPIA4) || mTestApp.Equals(Constants.ETRICCUI))
                        {
                            if (m_installScriptDir.IndexOf("Protected") > 0)
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no epia msi file found", mTestApp + "+" + mCurrentPlatform + "Protected");
                            else
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no epia msi file found", mTestApp + "+" + mCurrentPlatform + "Normal");
                        }
                        else
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no epia msi file found", mTestApp + "+" + mCurrentPlatform);
                        logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : m_testApp " + mTestApp, 0, 0);

                        FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                    }
                    return;
                }
                #endregion
            }
            else if (mTestApp.Equals(Constants.ETRICCSTATISTICS))
            {
                if (m_TestAutoMode)
                    m_installScriptDir = m_ValidatedBuildDropFolder + sMsiRelativePath;

                if (sMsgDebug.StartsWith("true"))
                    System.Windows.MessageBox.Show("m_installScriptDir...   " + m_installScriptDir);

                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_STATISTICS_PARSER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_STATISTICS_PARSERCONFIGURATOR);
            }
            else
            {
                System.Windows.MessageBox.Show("Unknown Application, try other application again...   " + mTestApp);
                return;
            }
            #endregion

            #region    // Update build quality   "Deployment Started"
            string msgX = "update build quality Deployment Started";
            TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
            while (TFSConnected == false)
            {
                TestTools.MessageBoxEx.Show( msgX + "\nWill try to reconnect the Server after 10 minutes",
                       "update build quality Deployment Started", 10 * 60000);
                System.Threading.Thread.Sleep(10 * 60000);
                TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
            }


            if (TFSConnected)
            {
                string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, BuildUtilities.GetProjectName(mTestApp), "Deployment Started", m_BuildSvc, sDemonstration);

                Log(updateResult);
                if (updateResult.StartsWith("Error"))
                {
                    System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                    throw new Exception(updateResult);
                }

                //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                if (m_TestAutoMode)
                {
                    if (mTestApp.Equals(Constants.EPIA4)|| mTestApp.Equals(Constants.ETRICCUI))
                    {
                        if (m_installScriptDir.IndexOf("Protected") > 0)
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Started", mTestApp + "+" + mCurrentPlatform + "Protected");
                        else
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Started", mTestApp + "+" + mCurrentPlatform + "Normal");
                    }
                    else
                    {
                        TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Started", mTestApp + "+" + mCurrentPlatform);
                    }

                    logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Started : mTestApp " + mTestApp, 0, 0);

                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.Forms.MessageBox.Show("Deployment Started:" + m_installScriptDir + "--testApp-" + mTestApp + " --testDef-" + mTestDef, "m_BuildNumber:" + m_BuildNumber);
              
                }
            }
            #endregion
            Thread.Sleep(3000);
            string installPath = m_CurrentDrive;   // only Etricc 5 depende on Drive, all other apps are on C:
            string dir = string.Empty;
            // Start Deployment   .......
            #region// Start Deployment   .......

            try
            {
                sCurrentBuildInTesting = "Start Deployment: " + m_installScriptDir;   
                sEpia4InstallerName = Constants.sEpia4InstallerName;

                if (mTestApp.Equals(Constants.ETRICC5))
                {
                    #region // Etricc 5
                    //MessageBox.Show("Start depmoyment:" + m_BuildNumber);
                    logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
                    System.Threading.Thread.Sleep(2000);
                   
                    //  Install setup
                    //Install new setup and recompile Worker
                    mEpiaPath = m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName()+ @"\Egemin\Etricc Server" + "\\";
                    //MessageBox.Show(" Install path :" + mEpiaPath);
                    Log(" Install path :" + mEpiaPath);
                    logger.LogMessageToFile(" Install path :" + mEpiaPath, 0, 0);
                
                    // Remove Old SetUP
                    RemoveSetup(Constants.ETRICC5);

                    //Move the current Setup files to a backup location
                    if (!CopySetup(mCurrentSetupPathEtricc5, mPreviousSetupPathEtricc5))
                        return;

                    if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtricc5, sEtriccMsiName))
                        return;
                         
                    // check if current folder is empty
                    FileInfo[] FilesToCopy;
                    DirectoryInfo DirInfo = new DirectoryInfo(mCurrentSetupPathEtricc5);
                    FilesToCopy = DirInfo.GetFiles(sEtriccMsiName);
                    if (FilesToCopy.Length == 0)
                    {
                        #region    // Update build quality   "Deployment Failed"
                        msgX = "update build quality Deployment Failed";
                        TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                                  "update build quality Deployment Failed", 10 * 60000);
                            System.Threading.Thread.Sleep(10 * 60000);
                            TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, BuildUtilities.GetProjectName(mTestApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no 'Etricc ?.msi' file found", mTestApp + "+" + mCurrentPlatform);
                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : mTestApp " + mTestApp, 0, 0);
                            }
                        }
                        #endregion
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
                    string version = string.Empty;
                    if (InstallEtricc5Setup(mCurrentSetupPathEtricc5, version))
                    {
                        string cscOut = RecompileTestRuns();
                        logger.LogMessageToFile("------ReCOMPILE TestRUNS output -------:" + cscOut, 0, 0);
                    }
                    #endregion
                }
                else if (mTestApp.Equals(Constants.EPIA4))
                {
                    #region // Epia4
                    //MessageBox.Show("Start depmoyment:" + m_BuildNumber);
                    logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
                    System.Threading.Thread.Sleep(2000);

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName()+ "\\Egemin\\Epia Server"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server", true);
                    }

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell"))
                    {
                        System.IO.Directory.Delete(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell", true);
                    }

                    //  Install Epia setup
                    //Install new setup
                    mEpiaPath = m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Epia Server";
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

                    // Remove Old Epia SetUP
                    RemoveSetup(Constants.EPIA4);

                    //Move the current Setup files to a backup location
                    if (!CopySetup(mCurrentSetupPathEpia, mPreviousSetupPathEpia))
                        return;

                    //MessageBox.Show(" m_installScriptDir :" + m_installScriptDir);
                    if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, "*" + sEpia4InstallerName))
                        return;

                    // check if current folder is empty
                    FileInfo[] FilesToCopy;
                    DirectoryInfo DirInfo = new DirectoryInfo(mCurrentSetupPathEpia);
                    FilesToCopy = DirInfo.GetFiles("*" + sEpia4InstallerName);
                    if (FilesToCopy.Length == 0)
                    {
                        #region    // Update build quality   "Deployment Failed"
                        msgX = "update build quality Deployment Failed2";
                        TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                                    "update build quality Deployment Failed2", 10 * 60000);
                            System.Threading.Thread.Sleep(10 * 60000);
                            TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, BuildUtilities.GetProjectName(mTestApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                if (mTestApp.Equals(Constants.EPIA4) )
                                {
                                    if (m_installScriptDir.IndexOf("Protected") > 0)
                                        TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no msi file found", mTestApp + "+" + mCurrentPlatform + "Protected");
                                    else
                                        TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no msi file found", mTestApp + "+" + mCurrentPlatform + "Normal");
                                }
                                else
                                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no msi file found", mTestApp + "+" + mCurrentPlatform);
                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : m_testApp " + mTestApp, 0, 0);
                            }
                        }
                        #endregion
                        return;
                    }

                    // Install Current Setup
                    if (InstallEpiaSetup(mCurrentSetupPathEpia))
                    {
                        //MessageBox.Show(" Epia installed :" + mTestedVersion);
                        //string cscOut = RecompileTestRuns();
                        //logger.LogMessageToFile("------ReCOMPILE TestRUNS output -------:" + cscOut, 0, 0);
                    }
                    #endregion
                }
                else if (mTestApp.Equals(Constants.ETRICCUI))
                {
                    #region // EtriccUI
                    //MessageBox.Show("Start depmoyment:" + m_BuildNumber);
                    logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
                    System.Threading.Thread.Sleep(2000);
                   
                    //  Install setup
                    //Install new setup
                  
                    // Remove Old SetUP
                    RemoveSetup(Constants.EPIA4);
                    RemoveSetup(Constants.ETRICCUI);   // because the register is EtriccInstallation
                    RemoveSetup(Constants.ETRICC5);
                    
                    //Move the current Etricc Setup files to a backup location
                    //---------------------------------------------------------------
                    if (!CopySetup(mCurrentSetupPathEtricc, mPreviousSetupPathEtricc))
                        return;

                           
                    //string s = " m_installScriptDir :" + m_installScriptDir + "\n\n mCurrentSetupPathEtricc :" + mCurrentSetupPathEtricc;
                    //TestTools.MessageBoxEx.Show(s, 30);
                    string EtriccShellMsi = "Etricc Shell.msi";
                    string EtriccShellMsiFile = System.IO.Path.Combine(m_installScriptDir, EtriccShellMsi);
                    if (System.IO.File.Exists(EtriccShellMsiFile))
                    {
                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtricc, "*Shell.msi"))
                            return;
                    }
                    else
                    {
                        #region    // Update build quality   "Deployment Failed"
                        msgX = "update build quality Deployment Failed3";
                        TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                                    "update build quality Deployment Failed3", 10 * 60000);
                            System.Threading.Thread.Sleep(10 * 60000);
                            TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {   // build quality not update if no msi file exist, but log as failed in log file
                            /*string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(mTestApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }
                            */
                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                if ( mTestApp.Equals(Constants.ETRICCUI))
                                {
                                    if (m_installScriptDir.IndexOf("Protected") > 0)
                                        TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no '*Shell.msi' file found", mTestApp + "+" + mCurrentPlatform + "Protected");
                                    else
                                        TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no '*Shell.msi' file found", mTestApp + "+" + mCurrentPlatform + "Normal");
                                }
                                else
                                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no '*Shell.msi' file found", mTestApp + "+" + mCurrentPlatform);

                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : m_testApp " + mTestApp, 0, 0);

                                FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                            }
                        }
                        #endregion
                        return;
                    }
                  
                    //---------------------------------------------------------------
                    //Move the current Epia Setup files to a backup location
                    if (!CopySetup(mCurrentSetupPathEpia, mPreviousSetupPathEpia))
                        return;

                    if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, Constants.EPIA_MSI))
                        return;

                    //---------------------------------------------------------------
                    //Move the current Etricc5 Setup files to a backup location
                    if (!CopySetup(mCurrentSetupPathEtricc5, mPreviousSetupPathEtricc5))
                        return;

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
                        return;
                    }
                    else
                    {
                        #region    // Update build quality   "Deployment Failed"
                            msgX = "update build quality Deployment Failed5";
                            TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                            while (TFSConnected == false)
                            {
                                TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                                        "update build quality Deployment Failed5", 10 * 60000);
                                System.Threading.Thread.Sleep(10 * 60000);
                                TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                            }
                        if (TFSConnected)
                        {
                            // build quality not update if no msi file exist, but log as failed in log file
                            /*string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, BuildUtilities.GetProjectName(mTestApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }*/

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no 'Etricc ?.msi' file found", mTestApp + "+" + mCurrentPlatform);
                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : m_testApp " + mTestApp, 0, 0);
                                FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                            }
                        }
                        #endregion
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

                    // Install Current Epia Setup 
                    if (InstallEpiaSetup(mCurrentSetupPathEpia))
                    {
                        //MessageBox.Show(" Epia installed :" + mTestedVersion);
                        logger.LogMessageToFile(" Epia installed :" + m_BuildNumber, 0, 0);
                        Thread.Sleep(1000);
                    }

                    // Install Current Etricc Core Setup 
                    if (InstallEtricc5Setup(mCurrentSetupPathEtricc5, ""))
                    {
                        //MessageBox.Show(" Etricc5 installed :" + mTestedVersion);
                        logger.LogMessageToFile(" Etricc5 installed :" + m_BuildNumber, 0, 0);
                    }

                    if (InstallEtriccUISetup(mCurrentSetupPathEtricc, m_BuildNumber))
                    {
                        //MessageBox.Show(" Etricc installed :" + mTestedVersion);
                        logger.LogMessageToFile(" Etricc installed :" + m_BuildNumber, 0, 0);
                    }

                   
                    sEtricc5InstallationFolder = m_SystemDrive + Constants.sEtricc5InstallationFolder; ;
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

                    //  Install setup
                    //Install new setup
                    // Remove Old SetUP
                    RemoveSetup(Constants.EPIA4);
                    RemoveSetup(Constants.ETRICCSTATISTICS_UI);  
                    RemoveSetup(Constants.ETRICCSTATISTICS_PARSERCONFIGURATOR);
                    RemoveSetup(Constants.ETRICCSTATISTICS_PARSER_SETUP);

                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.MessageBox.Show("Check removed applications:");

                    //(1) Move the current Etricc Statistics UI Setup files to a backup location
                    #region //---------------------------------------------------------------
                    if (!CopySetup(mCurrentSetupPathEtriccStatisticsUI, mPreviousSetupPathEtriccStatisticsUI))
                        return;

                    if (sMsgDebug.StartsWith("true"))
                        MessageBox.Show(" m_installScriptDir :" + m_installScriptDir);

                    string EtriccStatisticsUIMsi = "Etricc.Statistics.UI.msi";
                    string EtriccStatisticsUIMsiFile = System.IO.Path.Combine(m_installScriptDir, EtriccStatisticsUIMsi);
                    if (System.IO.File.Exists(EtriccStatisticsUIMsiFile))
                    {
                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccStatisticsUI, EtriccStatisticsUIMsi))
                        return;
                    }
                    else
                    {
                        #region    // Update build quality   "Deployment Failed"
                        msgX = "update build quality Deployment Failed3";
                        TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                                    "update build quality Deployment Failed3", 10 * 60000);
                            System.Threading.Thread.Sleep(10 * 60000);
                            TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            /*string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(mTestApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }*/

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no 'Etricc.Statistics.UI.msi' file found", mTestApp + "+" + mCurrentPlatform);
                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : m_testApp " + mTestApp, 0, 0);

                                FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                            }
                        }
                        #endregion
                        return;
                    }
                    #endregion

                    //(2) Move the current Etricc Statistics Parser Setup files to a backup location
                    #region //---------------------------------------------------------------
                    if (!CopySetup(mCurrentSetupPathEtriccStatisticsParser, mPreviousSetupPathEtriccStatisticsParser))
                        return;

                    if (sMsgDebug.StartsWith("true"))
                        MessageBox.Show(" m_installScriptDir :" + m_installScriptDir);

                    string EtriccStatisticsParserMsi = "Etricc.Statistics.Parser.msi";
                    string EtriccStatisticsParserMsiFile = System.IO.Path.Combine(m_installScriptDir, EtriccStatisticsParserMsi);
                    if (System.IO.File.Exists(EtriccStatisticsParserMsiFile))
                    {
                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccStatisticsParser, EtriccStatisticsParserMsi))
                        return;
                    }
                    else
                    {
                        #region    // Update build quality   "Deployment Failed"
                        msgX = "update build quality Deployment Failed3";
                        TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                                    "update build quality Deployment Failed3", 10 * 60000);
                            System.Threading.Thread.Sleep(10 * 60000);
                            TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            /*string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(mTestApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }*/

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no 'Etricc.Statistics.Parser.msi' file found", mTestApp + "+" + mCurrentPlatform);
                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : m_testApp " + mTestApp, 0, 0);

                                FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                            }
                        }
                        #endregion
                        return;
                    }
                    #endregion

                    //(3) Move the current Etricc Statistics ParserConfigurator Setup files to a backup location
                    #region //---------------------------------------------------------------
                    if (!CopySetup(mCurrentSetupPathEtriccStatisticsParserConfigurator, mPreviousSetupPathEtriccStatisticsParserConfigurator))
                        return;

                    if (sMsgDebug.StartsWith("true"))
                        MessageBox.Show(" m_installScriptDir :" + m_installScriptDir);

                    string EtriccStatisticsParserConfiguratorMsi = "Etricc.Statistics.ParserConfigurator.msi";
                    string EtriccStatisticsParserConfiguratorMsiFile = System.IO.Path.Combine(m_installScriptDir, EtriccStatisticsParserConfiguratorMsi);
                    if (System.IO.File.Exists(EtriccStatisticsParserConfiguratorMsiFile))
                    {
                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccStatisticsParser, EtriccStatisticsParserConfiguratorMsi))
                        return;

                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtriccStatisticsParserConfigurator, EtriccStatisticsParserConfiguratorMsi))
                        return;

                    }
                    else
                    {
                        #region    // Update build quality   "Deployment Failed"
                        msgX = "update build quality Deployment Failed3";
                        TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                                    "update build quality Deployment Failed3", 10 * 60000);
                            System.Threading.Thread.Sleep(10 * 60000);
                            TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            /*string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(mTestApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }*/

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no 'Etricc.Statistics.ParserConfigurator.msi' file found", mTestApp + "+" + mCurrentPlatform);
                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : m_testApp " + mTestApp, 0, 0);

                                FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                            }
                        }
                        #endregion
                        return;
                    }
                    #endregion

                    //(4) Move the current Epia Setup files to a backup location
                    #region //---------------------------------------------------------------
                    if (!CopySetup(mCurrentSetupPathEpia, mPreviousSetupPathEpia))
                        return;

                    //string EtriccStatisticsParserConfiguratorMsi = "Etricc.Statistics.ParserConfigurator.msi";
                    string Epia4MsiFile = System.IO.Path.Combine(m_installScriptDir, sEpia4InstallerName);
                    if (System.IO.File.Exists(Epia4MsiFile))
                    {
                        if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, "*" + sEpia4InstallerName))
                        return;

                    }
                    else
                    {
                        #region    // Update build quality   "Deployment Failed"
                        msgX = "update build quality Deployment Failed4";
                        TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        while (TFSConnected == false)
                        {
                            TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                                    "update build quality Deployment Failed4", 10 * 60000);
                            System.Threading.Thread.Sleep(10 * 60000);
                            TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                        }
                        if (TFSConnected)
                        {
                            /*string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, BuildUtilities.GetProjectName(mTestApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }*/

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                if (mTestApp.Equals(Constants.EPIA4)|| mTestApp.Equals(Constants.ETRICCUI))
                                {
                                    if (m_installScriptDir.IndexOf("Protected") > 0)
                                        TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no epia msi file found", mTestApp + "+" + mCurrentPlatform + "Protected");
                                    else
                                        TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no epia msi file found", mTestApp + "+" + mCurrentPlatform + "Normal");
                                }
                                else
                                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no epia msi file found", mTestApp + "+" + mCurrentPlatform);
                                logger.LogMessageToFile(m_ValidatedBuildDropFolder + " Deployment Failed : m_testApp " + mTestApp, 0, 0);

                                FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                                
                            }
                        }
                        #endregion
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

                    if (System.IO.Directory.Exists(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName()+ "\\Egemin\\Epia Shell"))
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

                    // Install Current Setup 
                    if (InstallEtriccStatisticsParserSetup(mCurrentSetupPathEtriccStatisticsParser))
                    {
                        logger.LogMessageToFile(" EtriccStatisticsParser installed :" + m_BuildNumber, 0, 0);
                    }

                    if (InstallEtriccStatisticsParserConfiguratorSetup(mCurrentSetupPathEtriccStatisticsParserConfigurator))
                    {
                        logger.LogMessageToFile(" EtriccStatisticsParserConfigurator installed :" + m_BuildNumber, 0, 0);
                    }

                    if (InstallEpiaSetup(mCurrentSetupPathEpia))
                    {
                        logger.LogMessageToFile(" Epia installed :" + m_BuildNumber, 0, 0);
                        Thread.Sleep(1000);
                    }

                    if (InstallEtriccStatisticsUISetup(mCurrentSetupPathEtriccStatisticsUI))
                    {
                        logger.LogMessageToFile(" EtriccStatisticsUI installed :" + m_BuildNumber, 0, 0);
                    }
                    #endregion
                }
                else
                {
                    System.Windows.MessageBox.Show("Unknown Application, try other application again3..." + mTestApp);
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

                #region // update build quality to "Deployment Completed"
                // only if this is first time test and build quality not = "GUI Tests Failed"
                msgX = "update build quality Deployment Completed";
                TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                while (TFSConnected == false)
                {
                    TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                           "update build quality Deployment Completed", 10 * 60000);
                    System.Threading.Thread.Sleep(10 * 60000);
                    TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                }
                if (TFSConnected)
                {
                    Log(TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, BuildUtilities.GetProjectName(mTestApp), "Deployment Completed",
                        m_BuildSvc, sDemonstration));
                    //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Completed", m_BuildSvc);
                }

                if (m_TestAutoMode)
                {
                    if (mTestApp.Equals(Constants.EPIA4) || mTestApp.Equals(Constants.ETRICCUI ))
                    {
                        if (m_installScriptDir.IndexOf("Protected") > 0)
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Completed", mTestApp + "+" + mCurrentPlatform + "Protected");
                        else
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Completed", mTestApp + "+" + mCurrentPlatform + "Normal");
                    }
                    else
                        TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Completed", mTestApp + "+" + mCurrentPlatform);
                    logger.LogMessageToFile(m_ValidatedBuildDropFolder + "---   Deployment Completed : testApp " + mTestApp, 0, 0);
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
                if (mTestApp.Equals(Constants.EPIA4) ||
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
                        + " " + '"' + Constants.sTFSServerUrl + '"'              //  11
                        + " " + '"' + mServerRunAs + '"'                         //  12
                        + " " + '"' + mExcelVisible + '"'                        //  13
                        + " " + '"' + sDemonstration.ToString().ToLower() + '"' //  14    // from config file
                        + " " + '"' + mMail + '"';                               //  15

                    string EtriccArgs = string.Empty;
                    dir = m_TestWorkingDirectory;

                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.Forms.MessageBox.Show("" + m_TestWorkingDirectory, "TRY TO find exe");

                    string filename = string.Empty;
                    if (mTestApp.Equals(Constants.EPIA4))
                    {
                        #region
                        filename = "Egemin.Epia.Testing.UIAutoTest.exe";
                        if (m_installScriptDir.IndexOf("Protected") > 0)
                        {
                            filename = "Egemin.Epia.Testing.EPia4AppTestProtected.exe";
                            if (sMsgDebug.StartsWith("true"))
                                System.Windows.Forms.MessageBox.Show("" + m_installScriptDir, "Test");
                        }
                        //MessageBox.Show("1: "+dir);
                        // at TFS test application is at same location as deployment application
                        //dir = System.IO.Directory.GetCurrentDirectory();
                        string testpath = System.IO.Path.Combine(dir, filename);

                        if (!System.IO.File.Exists(testpath))
                        {
                            // for development only
                            int ib = dir.LastIndexOf("\\");
                            dir = dir.Substring(0, ib);
                            ib = dir.LastIndexOf("\\");
                            dir = dir.Substring(0, ib);
                            ib = dir.LastIndexOf("\\");
                            dir = dir.Substring(0, ib);
                            dir = dir + "\\GUIAutoTest\\bin\\Debug";
                        }
                        #endregion
                    }
                    else if (mTestApp.Equals(Constants.ETRICCUI))
                    {
                        #region
                        filename = "Egemin.Epia.Testing.EtriccUIAutoTest.exe";
                        //MessageBox.Show("1: "+dir);
                        // at TFS test application is at same location as deployment application
                        //dir = System.IO.Directory.GetCurrentDirectory();
                        string testpath = System.IO.Path.Combine(dir, filename); // 
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
                        + " " + '"' + Constants.sTFSServerUrl + '"'              //  11
                        + " " + '"' + mServerRunAs + '"'                         //  12
                        + " " + '"' + mExcelVisible + '"'                        //  13
                        + " " + '"' + sDemonstration.ToString().ToLower() + '"' //  14    // from config file
                        + " " + '"' + mMail + '"'                               //  15                        
                         + " " + '"' + AllowFunctionalTesting + '"'    //  16
                         + " " + '"' + sProjectFile + '"'             //  17
                         + " " + '"' + sEtricc5InstallationFolder + '"';  //  18
                       
                        
                       if (!System.IO.File.Exists(testpath))
                       {
                           // for development only
                           int ib = dir.LastIndexOf("\\");
                           dir = dir.Substring(0, ib);
                           ib = dir.LastIndexOf("\\");
                           dir = dir.Substring(0, ib);
                           ib = dir.LastIndexOf("\\");
                           dir = dir.Substring(0, ib);
                           dir = dir + "\\EtriccGUIAutoTest\\bin\\Debug";
                       }
                        #endregion
                    }

                    string arg = EpiaArgs;
                    if (mTestApp.StartsWith("Etricc"))
                        arg = EtriccArgs;

                    logger.LogMessageToFile(" TestApplication args :" + arg, 0, 0);
                    // UI Test started
                    logger.LogMessageToFile("dir:" + dir, 0, 0);
                   
                    System.Diagnostics.Process proc5 = new System.Diagnostics.Process();
                    proc5.EnableRaisingEvents = false;
                    proc5.StartInfo.FileName = filename;
                    proc5.StartInfo.Arguments = arg;
                    proc5.StartInfo.WorkingDirectory = dir;
                    proc5.Start();
                
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
                    msgX = "update build quality GUI Tests Started";
                    TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                    while (TFSConnected == false)
                    {
                        TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                               "update build quality GUI Tests Started", 10 * 60000);
                        System.Threading.Thread.Sleep(10 * 60000);
                        TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                    }
                    if (TFSConnected)
                    {
                        Log(TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(mTestApp), "GUI Tests Started",
                            m_BuildSvc, sDemonstration));
                        //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "GUI Tests Started", m_BuildSvc);
                        string quality2 = m_BuildSvc.GetMinimalBuildDetails(m_Uri).Quality;
                        logger.LogMessageToFile(")))))))))))) start test worker:::::::::::::::::::quality:" + quality2, 0, 0);
                    }

                    if (m_TestAutoMode)
                    {
                        if (mTestApp.Equals(Constants.EPIA4) || mTestApp.Equals(Constants.ETRICCUI) )
                        {
                            if (m_installScriptDir.IndexOf("Protected") > 0)
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", mTestApp + "+" + mCurrentPlatform + "Protected");
                            else
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", mTestApp + "+" + mCurrentPlatform + "Normal");
                        }
                        else
                            FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", mTestApp + "+" + mCurrentPlatform);
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
                        ClickUiScreenActionToAvoidScreenStandBy();
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
                    TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                    while (TFSConnected == false)
                    {
                        TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                               "update build quality GUI Status", 10 * 60000);
                        System.Threading.Thread.Sleep(10 * 60000);
                        TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                    }
                    if (TFSConnected)
                    {
                        Log(TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, BuildUtilities.GetProjectName(mTestApp), GUIstatus,
                            m_BuildSvc, sDemonstration));
                        //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), GUIstatus, m_BuildSvc);
                    }

                    if (m_TestAutoMode)
                    {
                        if (exceptionMsg.Length > 10)
                        {
                            if (mTestApp.Equals(Constants.EPIA4) || mTestApp.Equals(Constants.ETRICCUI) )
                            {
                                if (m_installScriptDir.IndexOf("Protected") > 0)
                                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, exceptionMsg, mTestApp + "+" + mCurrentPlatform + "Protected");
                                else
                                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, exceptionMsg, mTestApp + "+" + mCurrentPlatform + "Normal");
                            }
                            else
                                FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, exceptionMsg, mTestApp + "+" + mCurrentPlatform);
                            logger.LogMessageToFile(m_ValidatedBuildDropFolder + "exceptionMsg : mTestApp " + mTestApp, 0, 0);
                        }
                        else
                        {
                            if (mTestApp.Equals(Constants.EPIA4) || mTestApp.Equals(Constants.ETRICCUI) )
                            {
                                if (m_installScriptDir.IndexOf("Protected") > 0)
                                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, GUIstatus, mTestApp + "+" + mCurrentPlatform + "Protected");
                                else
                                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, GUIstatus, mTestApp + "+" + mCurrentPlatform + "Normal");
                            }
                            else
                                FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, GUIstatus, mTestApp + "+" + mCurrentPlatform);
                            logger.LogMessageToFile(m_ValidatedBuildDropFolder + "GUIstatus : m_testApp " + mTestApp, 0, 0);
                        }
                    }
                    #endregion

                    Utilities.CloseProcess("EPIA.Launcher");
                    Utilities.CloseProcess("EPIA.Explorer");

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
                        + " " + '"' + Constants.sTFSServerUrl + '"'              //  11
                        + " " + '"' + mServerRunAs + '"'                         //  12
                        + " " + '"' + mExcelVisible + '"'                        //  13
                        + " " + '"' + sDemonstration.ToString().ToLower() + '"' //  14    // from config file
                        + " " + '"' + mMail + '"';                               //  15

                    dir = m_TestWorkingDirectory;
                    string filename = "Egemin.Epia.Testing.EtriccStatisticsProgTest.exe";
                    //MessageBox.Show("1: "+dir);
                    // at TFS test application is at same location as deployment application
                    //dir = System.IO.Directory.GetCurrentDirectory();
                    string testpath = System.IO.Path.Combine(dir, filename);

                    string arg = EpiaArgs;

                    logger.LogMessageToFile(" TestApplication args :" + arg, 0, 0);
                    // UI Test started
                    logger.LogMessageToFile("dir:" + dir, 0, 0);

                    System.Diagnostics.Process proc5 = new System.Diagnostics.Process();
                    proc5.EnableRaisingEvents = false;
                    proc5.StartInfo.FileName = filename;
                    proc5.StartInfo.Arguments = arg;
                    proc5.StartInfo.WorkingDirectory = dir;
                    proc5.Start();

                    m_TestAutoMode = false;
                    
                    #endregion
                }
            }
            catch (Exception ex)
            {
                // Your error handler here
                System.Windows.Forms.MessageBox.Show(ex.Message + System.Environment.NewLine + ex.StackTrace
                    + "-- m_installScriptDir" + m_installScriptDir,
                    "Started dir:" + dir);
                Log(ex.Message + System.Environment.NewLine + ex.StackTrace);

                // Update build quality
                msgX = "update build quality Deployment Failed6";
                TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                while (TFSConnected == false)
                {
                    TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 10 minutes",
                           "update build quality Deployment Failed6", 10 * 60000);
                    System.Threading.Thread.Sleep(10 * 60000);
                    TFSConnected = BuildUtilities.CheckTFSConnection(ref msgX);
                }

                if (TFSConnected)
                {
                    Log(TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(mTestApp), "Deployment Failed",
                        m_BuildSvc, sDemonstration));

                    if (m_TestAutoMode)
                    {
                        if (ex.Message.IndexOf("RecompileTestRuns") >= 0)
                        {
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Recompile TestRuns Exception:"
                                + ex.Message + "---" + ex.StackTrace, mTestApp+"+"+mCurrentPlatform);

                            logger.LogMessageToFile(m_ValidatedBuildDropFolder + "Recompile TestRuns Exception: : mTestApp " + mTestApp, 0, 0);
                        }
                        else
                        {
                            if (mTestApp.Equals(Constants.EPIA4) || mTestApp.Equals(Constants.ETRICCUI))
                            {
                                if (m_installScriptDir.IndexOf("Protected") > 0)
                                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Exception:" + ex.Message + "---" + ex.StackTrace, mTestApp + "+" + mCurrentPlatform + "Protected");
                                else
                                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Exception:" + ex.Message + "---" + ex.StackTrace, mTestApp + "+" + mCurrentPlatform + "Normal");
                            }
                            else
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Exception:"
                                    + ex.Message + "---" + ex.StackTrace, mTestApp+"+"+mCurrentPlatform);
                            logger.LogMessageToFile(m_ValidatedBuildDropFolder + "Deployment Exception: : mTestApp " + mTestApp+"+"+mCurrentPlatform, 0, 0);
                        }
                    }
                }

                // Set TestWorking to false
                FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
            }
            finally
            {
                m_State = STATE.PENDING;
            }
            #endregion
        }
        
        public BuildObject GetValidatedBuildDirectory(List<BuildObject> allBuildsInfo, string testPC, string testApp, string platform, bool testProtected,
            ref string currentPlatform, ref string msiRelativePath )
        {
            // platform : AnyCPU, x86, AnyCPU+x86
            currentPlatform = platform;
            string currentPlatformApp = platform +"Normal";
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
                currentBuildLocation = currenBuildObject.DripLoc;
                if (sMsgDebug.StartsWith("true"))
                {
                    System.Windows.Forms.MessageBox.Show(" currentBuildNr:" + currentBuildNr, "current build");
                    System.Windows.Forms.MessageBox.Show(" currentBuildLocation:" + currentBuildLocation, "current build");
                }

                sTestResultFolder = currentBuildLocation + "\\"+Constants.sTestResultFolderName;
                //System.IO.Directory.CreateDirectory(sTestResultFolder);
                if (!System.IO.Directory.Exists(sTestResultFolder))
                    System.IO.Directory.CreateDirectory(sTestResultFolder);

                if (sMsgDebug.StartsWith("true"))
                    System.Windows.MessageBox.Show("testResultDirectory:" + sTestResultFolder);

                // check TestInfo.txt and TestWorking.txt files
                string testInfoTxtFile = Path.Combine(sTestResultFolder, ConstCommon.TESTINFO_FILENAME);
                string testWorkingFile = Path.Combine(sTestResultFolder, ConstCommon.TESTWORKING_FILENAME);

                // check build Succeeding by check Building.txt has text "0 Error(s)"
                string errorMsg = "CheckBuildSucceeding:";
                string buildlogfile = string.Empty;
                
                if (testApp.Equals(Constants.EPIA4))
                {
                    #region // Epia4
                    string caseApp = string.Empty;
                    if (platform.Equals("AnyCPU"))
                    { 
                        if (testProtected)
                            caseApp = "AnyCPU+Protected";
                        else
                            caseApp = "AnyCPU+Only";
                    }
                    else if (platform.Equals("x86"))
                    {
                        if (testProtected)
                            caseApp = "x86+Protected";
                        else
                            caseApp = "x86+Only";
                    }
                    else if (platform.Equals("AnyCPU+x86"))
                    {
                        if (testProtected)
                            caseApp = "AnyCPU+x86+Protected";
                        else
                            caseApp = "AnyCPU+x86+Only";
                    }


                    switch (caseApp )
                    {
                        case "AnyCPU+Only":
                            // get current platform
                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUNormal") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUNormal";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr + "is already tested by AnyCPU ONLY ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPU x86 not tested, first test AnyCPU
                            {
                                currentPlatform = "AnyCPU";
                                currentPlatformApp = "AnyCPUNormal";
                            }
                            break;
                        case "AnyCPU+Protected":
                            // get protected directory 
                            string protectedDirectory = currentBuildLocation + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaRelativeProtectedPath");
                            bool protectedDirectoryExist = false;
                            if (System.IO.Directory.Exists(protectedDirectory))
                                protectedDirectoryExist = true;

                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUNormal") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUNormal";
                                }
                                else if (protectedDirectoryExist == true && IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUProtected") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUProtected";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr +
                                        "is already tested by AnyCPUNormal and if protected directory exist AnyCPUProtected also tested ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPUNormal and AnyCPUProtected x86 tested, first test AnyCPUNormal
                            {
                                currentPlatform = "AnyCPU";
                                currentPlatformApp = "AnyCPUNormal";
                            }
                            break;
                        case "x86+Only":  // Only x86 no protected
                            // get current platform
                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Normal") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Normal";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr + "is already tested by x86 ONLY ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPU x86 not tested, first test AnyCPU
                            {
                                currentPlatform = "x86";
                                currentPlatformApp = "x86Normal";
                            }
                            break;
                        case "x86+Protected":
                            // get protected directory 
                            protectedDirectory = currentBuildLocation + System.Configuration.ConfigurationManager.AppSettings.Get("Epiax86RelativeProtectedPath");
                            protectedDirectoryExist = false;
                            if (System.IO.Directory.Exists(protectedDirectory))
                                protectedDirectoryExist = true;

                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Normal") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Normal";
                                }
                                else if (protectedDirectoryExist == true && IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Protected") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Protected";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr +
                                        "is already tested by x86Normal and if protected directory exist x86Protected also tested ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPUNormal and AnyCPUProtected x86 tested, first test AnyCPUNormal
                            {
                                currentPlatform = "x86";
                                currentPlatformApp = "x86Normal";
                            }
                            break;
                        case "AnyCPU+x86+Only":  // Only x86 no protected
                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUNormal") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUNormal";
                                }
                                else if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Normal") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Normal";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr + "is already tested by both AnyCPU and x86 ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPU x86 not tested, first test AnyCPU
                            {
                                currentPlatform = "AnyCPU";
                                currentPlatformApp = "AnyCPUNormal";
                            }
                            break;
                        case "AnyCPU+x86+Protected":  // Only x86 no protected
                            string AnyCPUprotectedDirectory = currentBuildLocation + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaRelativeProtectedPath");
                            bool AnyCPUprotectedDirectoryExist = false;
                            if (System.IO.Directory.Exists(AnyCPUprotectedDirectory))
                                AnyCPUprotectedDirectoryExist = true;

                            string x86protectedDirectory = currentBuildLocation + System.Configuration.ConfigurationManager.AppSettings.Get("Epiax86RelativeProtectedPath");
                            bool x86protectedDirectoryExist = false;
                            if (System.IO.Directory.Exists(x86protectedDirectory))
                                x86protectedDirectoryExist = true;

                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUNormal") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUNormal";
                                }
                                else if (AnyCPUprotectedDirectoryExist == true && IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUProtected") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUProtected";
                                }
                                else if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Normal") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Normal";
                                }
                                else if (x86protectedDirectoryExist == true && IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Protected") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Protected";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr + "is already tested by both AnyCPU and x86 ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPU x86 not tested, first test AnyCPU
                            {
                                currentPlatform = "AnyCPU";
                                currentPlatformApp = "AnyCPUNormal";
                            }
                            break;
                    }

                    if (currentPlatformApp.Equals("AnyCPUNormal"))
                    {
                        //epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaAnyCPUBuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaRelativePath");
                    }

                    if (currentPlatformApp.Equals("AnyCPUProtected"))
                    {
                        //epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaAnyCPUBuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaRelativeProtectedPath");
                    }

                    if (currentPlatformApp.Equals("x86Normal"))
                    {
                        //epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaAnyCPUBuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Epiax86RelativePath");
                    }

                    if (currentPlatformApp.Equals("x86Protected"))
                    {
                        //epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaAnyCPUBuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Epiax86RelativeProtectedPath");
                    }

                    /*
                    // get current platform
                    if (platform.IndexOf('+') == -1)    // AnyCPU or x86
                        currentPlatform = platform;
                    else                                // AnyCpu+x86
                    {
                        // get current platform
                        if (File.Exists(testInfoTxtFile))
                        {
                            if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPU") == false)
                                 currentPlatform = "AnyCPU";
                            else if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86") == false)
                                currentPlatform = "x86";
                            else
                            {
                                logger.LogMessageToFile("current build ..." + currentBuildNr + "is already tested by AnyCPU and x86 ", sLogCount, sLogInterval);
                                currenBuildObject = null;
                                continue;
                            }
                        }
                        else // both AnyCPU x86 not tested, first test AnyCPU
                            currentPlatform = "AnyCPU";
                    }
                    
                    string epiaBuildLogFilePath = string.Empty;
                    if (currentPlatform.Equals("AnyCPU"))
                    {
                        epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaAnyCPUBuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaRelativePath");
                    }
                    else if (currentPlatform.Equals("x86"))
                    {
                         epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("Epiax86BuildLogFile");
                         msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Epiax86RelativePath");
                    }
                    else if (currentPlatform.Equals("x64"))
                    {
                        epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("Epiax64BuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Epiax64RelativePath");
                    }

                    buildlogfile = currentBuildLocation + epiaBuildLogFilePath;
                    if (sMsgDebug.StartsWith("true"))
                        System.Windows.MessageBox.Show("buildlogfile:" + buildlogfile);
                    */
                    #endregion
                }
                else if (testApp.Equals(Constants.ETRICCUI) )
                {
                    #region // // EtriccUI
                    string caseApp = string.Empty;
                    if (platform.Equals("AnyCPU"))
                    {
                        //if (testProtected)
                        //    caseApp = "AnyCPU+Protected";
                        //else
                            caseApp = "AnyCPU+Only";
                    }
                    else if (platform.Equals("x86"))
                    {
                        //if (testProtected)
                        //    caseApp = "x86+Protected";
                        //else
                            caseApp = "x86+Only";
                    }
                    else if (platform.Equals("AnyCPU+x86"))
                    {
                        //if (testProtected)
                        //    caseApp = "AnyCPU+x86+Protected";
                        //else
                            caseApp = "AnyCPU+x86+Only";
                    }


                    switch (caseApp)
                    {
                        case "AnyCPU+Only":
                            // get current platform
                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUNormal") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUNormal";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr + "is already tested by AnyCPU ONLY ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPU x86 not tested, first test AnyCPU
                            {
                                currentPlatform = "AnyCPU";
                                currentPlatformApp = "AnyCPUNormal";
                            }
                            break;
                        case "AnyCPU+Protected":
                            // get protected directory 
                            string protectedDirectory = currentBuildLocation + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIRelativeProtectedPath");
                            bool protectedDirectoryExist = false;
                            if (System.IO.Directory.Exists(protectedDirectory))
                                protectedDirectoryExist = true;

                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUNormal") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUNormal";
                                }
                                else if (protectedDirectoryExist == true && IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUProtected") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUProtected";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr +
                                        "is already tested by AnyCPUNormal and if protected directory exist AnyCPUProtected also tested ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPUNormal and AnyCPUProtected x86 tested, first test AnyCPUNormal
                            {
                                currentPlatform = "AnyCPU";
                                currentPlatformApp = "AnyCPUNormal";
                            }
                            break;
                        case "x86+Only":  // Only x86 no protected
                            // get current platform
                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Normal") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Normal";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr + "is already tested by x86 ONLY ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPU x86 not tested, first test AnyCPU
                            {
                                currentPlatform = "x86";
                                currentPlatformApp = "x86Normal";
                            }
                            break;
                        case "x86+Protected":
                            // get protected directory 
                            protectedDirectory = currentBuildLocation + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx86RelativeProtectedPath");
                            protectedDirectoryExist = false;
                            if (System.IO.Directory.Exists(protectedDirectory))
                                protectedDirectoryExist = true;

                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Normal") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Normal";
                                }
                                else if (protectedDirectoryExist == true && IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Protected") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Protected";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr +
                                        "is already tested by x86Normal and if protected directory exist x86Protected also tested ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPUNormal and AnyCPUProtected x86 tested, first test AnyCPUNormal
                            {
                                currentPlatform = "x86";
                                currentPlatformApp = "x86Normal";
                            }
                            break;
                        case "AnyCPU+x86+Only":  // Only x86 no protected
                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUNormal") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUNormal";
                                }
                                else if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Normal") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Normal";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr + "is already tested by both AnyCPU and x86 ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPU x86 not tested, first test AnyCPU
                            {
                                currentPlatform = "AnyCPU";
                                currentPlatformApp = "AnyCPUNormal";
                            }
                            break;
                        case "AnyCPU+x86+Protected":  // Only x86 no protected
                            string AnyCPUprotectedDirectory = currentBuildLocation + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIRelativeProtectedPath");
                            bool AnyCPUprotectedDirectoryExist = false;
                            if (System.IO.Directory.Exists(AnyCPUprotectedDirectory))
                                AnyCPUprotectedDirectoryExist = true;

                            string x86protectedDirectory = currentBuildLocation + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx86RelativeProtectedPath");
                            bool x86protectedDirectoryExist = false;
                            if (System.IO.Directory.Exists(x86protectedDirectory))
                                x86protectedDirectoryExist = true;

                            if (File.Exists(testInfoTxtFile))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUNormal") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUNormal";
                                }
                                else if (AnyCPUprotectedDirectoryExist == true && IsThisPCTested(testInfoTxtFile, testPC, testApp, "AnyCPUProtected") == false)
                                {
                                    currentPlatform = "AnyCPU";
                                    currentPlatformApp = "AnyCPUProtected";
                                }
                                else if (IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Normal") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Normal";
                                }
                                else if (x86protectedDirectoryExist == true && IsThisPCTested(testInfoTxtFile, testPC, testApp, "x86Protected") == false)
                                {
                                    currentPlatform = "x86";
                                    currentPlatformApp = "x86Protected";
                                }
                                else
                                {
                                    logger.LogMessageToFile("current build ..." + currentBuildNr + "is already tested by both AnyCPU and x86 ", sLogCount, sLogInterval);
                                    currenBuildObject = null;
                                    continue;
                                }
                            }
                            else // both AnyCPU x86 not tested, first test AnyCPU
                            {
                                currentPlatform = "AnyCPU";
                                currentPlatformApp = "AnyCPUNormal";
                            }
                            break;
                    }

                    if (currentPlatformApp.Equals("AnyCPUNormal"))
                    {
                        //epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaAnyCPUBuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIRelativePath");
                    }

                    if (currentPlatformApp.Equals("AnyCPUProtected"))
                    {
                        //epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaAnyCPUBuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIRelativeProtectedPath");
                    }

                    if (currentPlatformApp.Equals("x86Normal"))
                    {
                        //epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaAnyCPUBuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx86RelativePath");
                    }

                    if (currentPlatformApp.Equals("x86Protected"))
                    {
                        //epiaBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaAnyCPUBuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx86RelativeProtectedPath");
                    }
                    #endregion
                    
                    #region // EtriccUI
                     /*
                    string etriccUiBuildLogFilePath = string.Empty;
                     if (currentPlatform.Equals("AnyCPU"))
                     {
                        etriccUiBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIAnyCPUBuildLogFile");
                         msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIRelativePath");
                     }
                    else if (currentPlatform.Equals("x86"))
                     {
                         etriccUiBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx86BuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx86RelativePath");
                    }
                    else if (currentPlatform.Equals("x64"))
                    {
                        etriccUiBuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx64BuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx64RelativePath");
                    }

                    buildlogfile = currentBuildLocation + etriccUiBuildLogFilePath;
                        */
                    #endregion
                }
                else if (testApp.Equals(Constants.ETRICC5))
                {
                    #region //Etricc5
                    string etricc5BuildLogFilePath = string.Empty;
                     if (currentPlatform.Equals("AnyCPU"))
                     {
                        etricc5BuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5AnyCPUBuildLogFile");
                         msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5RelativePath");
                     }
                    else if (currentPlatform.Equals("x86"))
                    {
                         etricc5BuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5x86BuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5x86RelativePath");
                    }
                    else if (currentPlatform.Equals("x64"))
                    {
                        etricc5BuildLogFilePath = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5x64BuildLogFile");
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5x64RelativePath");
                    }
                    buildlogfile = currentBuildLocation +  etricc5BuildLogFilePath;
                    #endregion
                }
                else if (testApp.Equals(Constants.ETRICCSTATISTICS)) //  etricc statistics
                {
                    #region //EtriccStatistics
                    if (currentPlatform.Equals("AnyCPU"))
                    {
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccStatisticsRelativePath");
                    }
                    else if (currentPlatform.Equals("x86"))
                    {
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccStatisticsx86RelativePath");
                    }
                    else if (currentPlatform.Equals("x64"))
                    {
                        msiRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccStatisticsx64RelativePath");
                    }
                    #endregion
                }
                else if (testApp.Equals(Constants.EWMS)) //  current_testApp will be decided next
                {
                    string ewmsBuildLogFilePath = string.Empty;
                    buildlogfile = currentBuildLocation + ewmsBuildLogFilePath;
                }
                else
                {
                    MessageBox.Show("Wrong testApp:" + testApp);
                    testApp = "Wrong testApp:";
                }

                logger.LogMessageToFile(testPC + " check buildlog file:" + buildlogfile, sLogCount, 0);
                #region check build log file (error )  --> test running ? -->  has tested by this PC  // not check log file anymore, canbe replaced by tfs build status
                //if (TestTools.FileManipulation.CheckSearchTextExistInFile(buildlogfile, "0 Error(s)", ref errorMsg))
                //{
                    // TestInfo file exist
                    if (File.Exists(testInfoTxtFile))
                    {
                        if (sMsgDebug.StartsWith("true"))
                            System.Windows.Forms.MessageBox.Show("is test working:" + sTestResultFolder);

                        if (IsTestWorking(sTestResultFolder) == false)
                        {
                            if (IsThisPCTested(testInfoTxtFile, testPC, testApp, currentPlatform) == false)
                            {
                                // Add test info
                                // Check test Working file
                                FileInfo workFile = new FileInfo(testWorkingFile);
                                File.SetAttributes(workFile.FullName, FileAttributes.Normal);

                                StreamReader readerInfo = File.OpenText(testInfoTxtFile);
                                string info = readerInfo.ReadToEnd();
                                readerInfo.Close();

                                if (sMsgDebug.StartsWith("true"))
                                    System.Windows.Forms.MessageBox.Show("EPIA4 Added Protected folder, now should see Normal or Protected App:testApp" + testApp);

                                if (testApp.Equals(Constants.EPIA4) || testApp.Equals(Constants.ETRICCUI) )
                                {   // EPIA4 Added Protected folder, now should see Normal or Protected App
                                    if (currentPlatformApp.Equals("AnyCPUNormal"))
                                        info = info + testPC + "+" + testApp + "+" + currentPlatform + "Normal" + "==Starting";          
                                    else  if (currentPlatformApp.Equals("AnyCPUProtected"))
                                        info = info + testPC + "+" + testApp + "+" + currentPlatform + "Protected" + "==Starting";
                                    else if (currentPlatformApp.Equals("x86Normal"))
                                        info = info + testPC + "+" + testApp + "+" + currentPlatform + "Normal" + "==Starting";
                                    else if (currentPlatformApp.Equals("x86Protected"))
                                        info = info + testPC + "+" + testApp + "+" + currentPlatform + "Protected" + "==Starting";
                                }
                                else
                                    info = info + testPC + "+" + testApp + "+" + currentPlatform + "==Starting";

                                if (sMsgDebug.StartsWith("true"))
                                    System.Windows.Forms.MessageBox.Show("Logline:" + info);


                                Log(testPC + " Added into InfoFile of build:" + currentBuildNr);
                                logger.LogMessageToFile(testPC + " Added into Info file:" + currentBuildNr, sLogCount, sLogInterval);

                                // ------------  if write infotext file failure, not do anything , continue go to iteration 
                                try
                                {
                                    StreamWriter writeInfo = File.CreateText(testInfoTxtFile);
                                    writeInfo.WriteLine(info);
                                    writeInfo.Close();
                                }
                                catch (Exception ex)
                                {
                                    string msg = testPC + " Exception Add pc to InfoText:" + currentBuildLocation + "=====" + ex.Message + "" + ex.StackTrace;
                                    Log(msg);
                                    logger.LogMessageToFile(msg, sLogCount, sLogInterval);
                                    continue;
                                }

                                // to force update, sometimes file is used by other process 
                                bool updateTestWorking = false;
                                while (updateTestWorking == false)
                                {
                                    try
                                    {
                                        // Create testWorking File
                                        StreamWriter writeWork = File.CreateText(testWorkingFile);
                                        writeWork.WriteLine("true");
                                        writeWork.Close();

                                        Log(testPC + " Set Working to true:" + currentBuildLocation);
                                        logger.LogMessageToFile(testPC + " Set Working to true:" + currentBuildLocation, sLogCount, sLogInterval);
                                        updateTestWorking = true;
                                    }
                                    catch (Exception ex)
                                    {
                                        string msg = testPC + " Exception Set Working to true:" + currentBuildLocation + "=====" + ex.Message + "" + ex.StackTrace;
                                        Log(msg);
                                        logger.LogMessageToFile(msg, sLogCount, sLogInterval);
                                        Thread.Sleep(10000);
                                        updateTestWorking = false;
                                    }
                                }

                                validBuild.Add(currentBuildNr, currentBuildLocation);  // will change to found validated build
                            }
                            else
                            {
                                Log(testPC + " is Tested for thisbuild:" + currentBuildLocation + " -- Build Definition:" + mTestDef);
                                logger.LogMessageToFile(testPC + " is Tested for thisbuild:" + currentBuildNr + " -- Build Definition:" + mTestDef,
                                    sLogCount, sLogInterval);
                            }
                        }
                        else
                        {
                            Log("Test is Working..." + currentBuildLocation);
                            logger.LogMessageToFile("Test is Working..." + currentBuildLocation, sLogCount, sLogInterval);
                        }
                    }
                    else // create testinfo file and create TestWorking File
                    {
                        //ToDO AnyCPU+x86 --> decide current platform
                        StreamWriter writeInfo = File.CreateText(testInfoTxtFile);


                        if (sMsgDebug.StartsWith("true"))
                            System.Windows.Forms.MessageBox.Show("EPIA4 Added Protected folder, now should see Normal or Protected App:testApp" + testApp);

                        if (testApp.Equals(Constants.EPIA4) || testApp.Equals(Constants.ETRICCUI) )
                        {   // EPIA4 Added Protected folder, now should see Normal or Protected App

                            if (currentPlatformApp.Equals("AnyCPUNormal"))
                                writeInfo.WriteLine("TestInfo" + System.Environment.NewLine
                                        + testPC + "+" + testApp + "+" + currentPlatform+ "Normal" + "=" + mTestDef + "==Starting");
                            else if (currentPlatformApp.Equals("AnyCPUProtected"))
                                writeInfo.WriteLine("TestInfo" + System.Environment.NewLine
                                       + testPC + "+" + testApp + "+" + currentPlatform + "Protected" + "=" + mTestDef + "==Starting");
                            else if (currentPlatformApp.Equals("x86Normal"))
                                writeInfo.WriteLine("TestInfo" + System.Environment.NewLine
                                        + testPC + "+" + testApp + "+" + currentPlatform+ "Normal" + "=" + mTestDef + "==Starting");
                            else if (currentPlatformApp.Equals("x86Protected"))
                                writeInfo.WriteLine("TestInfo" + System.Environment.NewLine
                                      + testPC + "+" + testApp + "+" + currentPlatform + "Protected" + "=" + mTestDef + "==Starting");
                        }
                        else
                            writeInfo.WriteLine("TestInfo" + System.Environment.NewLine
                           + testPC + "+" + testApp + "+" + currentPlatform + "=" + mTestDef + "==Starting");

                        writeInfo.Close();

                        // Create testWorking File
                        StreamWriter writeWork = File.CreateText(testWorkingFile);
                        writeWork.WriteLine("true");
                        writeWork.Close();
                        Log(currentBuildNr + " Deployment starting... ");
                        logger.LogMessageToFile(currentBuildNr + " Deployment starting...", sLogCount, sLogInterval);
                        validBuild.Add(currentBuildNr, currentBuildLocation);
                    }
                //}
                //else
                //{
                //    Log(currentBuildNr + " build is not secceeded..." + errorMsg);
                //    logger.LogMessageToFile(currentBuildNr + " build is not secceeded..." + errorMsg, sLogCount, sLogInterval);
                //}
                #endregion

                #region check msi file exist
                /*  not needed anymore, build status alreaady is succedded
                // check msi file exist 
                
                string msiPath = currentBuildLocation + msiRelativePath;
                try
                {
                    DirectoryInfo DirInfo = new DirectoryInfo(msiPath);
                    FileInfo[] msiFiles = DirInfo.GetFiles("*.msi");
                    if (msiFiles.Length == 0)
                    {
                        logger.LogMessageToFile(testPC + " no msi files deployed:" + msiPath, sLogCount, 0);
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    // check 
                    DateTime timeNow = DateTime.Now;
                    DirectoryInfo BuildDirInfo = new DirectoryInfo(currentBuildLocation);
                    DateTime buildDateTime = BuildDirInfo.CreationTime;
                    TimeSpan duration = timeNow - buildDateTime;

                    if (duration.TotalHours >= 10)
                    {
                        logger.LogMessageToFile(testPC + " After more than 10 hours, Still no msi files deployed::" + msiPath + "---" + ex.Message + "--" + ex.StackTrace, sLogCount, 0);
                        // ToDo :update build quality and log file  
                        #region    // Update build quality   "Deployment Failed"
                        if (TFSConnected)
                        {

                            m_Uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, BuildUtilities.GetProjectName(mTestApp), currentBuildNr);
                            string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, BuildUtilities.GetProjectName(testApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                System.Windows.Forms.MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }

                            if (!File.Exists(testInfoTxtFile))
                            {
                                //ToDO AnyCPU+x86 --> decide current platform
                                StreamWriter writeInfo = File.CreateText(testInfoTxtFile);
                                writeInfo.WriteLine("TestInfo" + System.Environment.NewLine
                                    + testPC + "+" + testApp + "+" + currentPlatform + "=" + mTestDef + "==Starting");

                                writeInfo.Close();

                                // Create testWorking File
                                StreamWriter writeWork = File.CreateText(testWorkingFile);
                                writeWork.WriteLine("false");
                                writeWork.Close();
                                Log(currentBuildNr + " Deployment starting... ");
                                logger.LogMessageToFile(currentBuildNr + " Deployment starting...", sLogCount, sLogInterval);
                                //validBuild.Add(currentBuildNr, currentBuildLocation);
                            }
                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                if (mTestApp.Equals(Constants.EPIA4))
                                {
                                    if (m_installScriptDir.IndexOf("Protected") > 0)
                                        TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no msi file found", mTestApp + "+" + mCurrentPlatform + "Protected");
                                    else
                                        TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "Deployment Failed: no msi file found", mTestApp + "+" + mCurrentPlatform + "Normal");
                                }
                                else
                                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(currentBuildLocation+"\\TestResults", "Deployment Failed: no msi file found", mTestApp + "+" + mCurrentPlatform);
                                logger.LogMessageToFile(currentBuildLocation + " Deployment Failed : testApp " + testApp, 0, 0);
                            }
                        }
                        #endregion

                    }
                    continue;
                }
                */
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
            try
            {
                StreamReader readerInfo = File.OpenText(testInfoTxtFile);
                string info = readerInfo.ReadToEnd();
                readerInfo.Close();

                if (info.IndexOf(testPC + "+" + testApp + "+" + platform + "=") > 0)
                {
                    //Log("but "+testPC + " is already in test info file");
                    //logger.LogMessageToFile(testPC + " is already in test info file", sLogCount, sLogInterval);
                    return true;
                }
                else
                {
                    Log(testPC + "+" + testApp + "+" + platform + "=" + " ---------->> is not in test info file");
                    logger.LogMessageToFile(testPC + "+" + testApp + "+" + platform + "=" + " ----------> is not in test info file",
                        sLogCount, sLogInterval);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log("IsTestWorking exception:" + testPC + " - message:" + ex.Message + " --- " + ex.StackTrace);
                logger.LogMessageToFile("IsTestWorking exception:" + testPC + "+" + testApp + "+" + platform + "=" + " - message:" + ex.Message + " --- " + ex.StackTrace,
                    sLogCount, sLogInterval);
                return true;
            }
        }

        public bool IsTestWorking(string path)
        {
            try
            {
                string testWorkingFile = Path.Combine(path, ConstCommon.TESTWORKING_FILENAME);
                // Check test Working file
                StreamReader readerWorking = File.OpenText(testWorkingFile);
                string worker = readerWorking.ReadLine();
                readerWorking.Close();

                if (worker.Trim().ToLower().Equals("true"))
                {
                    Log(path + " =========>> is UNDER TESTING...");
                    logger.LogMessageToFile(path + " =========>> is UNDER TESTING...", sLogCount, sLogInterval);
                    return true;
                }
                else
                {
                    //Log(path + " =========>> is NOT under TESTING...");
                    logger.LogMessageToFile(path + " =========>> is NOT under TESTING...", sLogCount, sLogInterval);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log("IsTestWorking exception:" + path + " - message:" + ex.Message + " --- " + ex.StackTrace);
                logger.LogMessageToFile("IsTestWorking exception:" + path + " - message:" + ex.Message + " --- " + ex.StackTrace,
                    sLogCount, sLogInterval);
                return true;
            }
        }

        private void RemoveSetup(string AppName)
        {
            logger.LogMessageToFile("------------Start Removed Setup AppName: " + AppName, 0, 0);

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

                logger.LogMessageToFile("(1)Removed Setup:" + AppName + " -- InstallPath " + InstallPath + " and InstallName " + InstallName, 0, 0);

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
                    "DeployTestLogic.Tester  RemoveSetup: find register Key");
                logger.LogMessageToFile(ex.ToString() + System.Environment.NewLine + ex.StackTrace, 0, 0);

            }

            if ((InstallPath == string.Empty) || (InstallName == string.Empty))
                return;

            string unattendedXmlFilePath = m_CurrentDrive + @"Epia 3\Testing\Automatic\AutomaticTests\TestData\Etricc5";
            if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))
                unattendedXmlFilePath = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY;

            //remove the setup in silent mode
            string xmlFile = System.IO.Path.Combine(unattendedXmlFilePath, "UnattendedXmlFile.xml");
            string uninstallParm = Path.Combine(InstallPath, InstallName);
            string args = "/passive /x " + '"' + uninstallParm + '"' + " " + "UnattendedXmlFile=" + '"' + xmlFile + '"';

            logger.LogMessageToFile("(1)Removed MsiExec.exe :" + args + " -- WorkingDirectory " + InstallPath, 0, 0);
            // remove Yes button
            AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnRemoveSetuoScreenEvent);
            // Add Open window Event Handler
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);
            Thread.Sleep(5000);

            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.EnableRaisingEvents = false;
            proc.StartInfo.FileName = "MsiExec.exe";
            proc.StartInfo.Arguments = args;
            proc.StartInfo.WorkingDirectory = InstallPath;
            proc.Start();
            proc.WaitForExit();

            Thread.Sleep(10000);

            Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                          AutomationElement.RootElement,
                         UIAShellEventHandler);

            //MessageBox.Show("args="+args);
            //Thread.Sleep(10000000);
            //msiexec /passive /x "D:\ZZ Dev\Epia\Setup.NotObfuscated\EpiaSetup.NotObfuscated.msi" UnattendedXmlFile="D:\ZZ Dev\Epia\Development\dotNET\Installer\UnattendedXmlFile.xml"
            Log("Removed Setup: " + AppName + " -- " + InstallName + " at " + InstallPath);
            try
            {
                //if (m_Settings.EnableLog)
                //{
                //string logPath = Configuration.BuildInformationfilePath;
                //string path = Path.Combine( logPath, Configuration.LogFilename );
                //Logger logger = new Logger(path );
                logger.LogMessageToFile("Removed Setup " + InstallName + " at " + InstallPath, 0, 0);

                //}
            }
            catch (Exception ex)
            {
                logger.LogMessageToFile(AppName + "------ Test Exception : " + ex.Message + "\r\n" + ex.StackTrace, 0, 0);
            }

            //Remove the install information from the registry
            string RegName = AppName;
            if (AppName.EndsWith(TestTools.ConstCommon.ETRICC_UI))
                RegName = Constants.ETRICCUI;
            WriteInstallationToReg(RegName, string.Empty, string.Empty);
        }

        private bool InstallEpiaSetup(string FilePath)
        {
            //MessageBox.Show("InstallEpiaSetup FilePath: " + FilePath);
            logger.LogMessageToFile("::: InstallEpiaSetup : " + FilePath, sLogCount, sLogInterval);

            //find the msi in the filepath
            string msiName = string.Empty;
            DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
            FileInfo[] files = DirInfo.GetFiles("*" + sEpia4InstallerName);
            bool installed = false;

            try
            {
                if (files[0] != null)
                    msiName = files[0].Name;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("cscript exception: " + ex.Message + " -- " + ex.StackTrace, "InstallSetup");
            }

            try
            {
                #region Install Epia Old version

                string[] SetupStepDescriptions = new string[100];
                SetupStepDescriptions[0] = "Welcome";
                SetupStepDescriptions[1] = "Welcome to the E'pia Framework 2010.05.11.1 Setup Wizard";
                SetupStepDescriptions[2] = "Components";
                SetupStepDescriptions[3] = "Installation Folders";
                SetupStepDescriptions[4] = "Confirm Installation";
                SetupStepDescriptions[5] = "Installing E'pia Framework ...";
                SetupStepDescriptions[6] = "E'pia Framework 2010.12.22.* Information";
                SetupStepDescriptions[7] = "Installation Complete";

                string SetupWindowName = string.Empty;
                string sErrorMessage = string.Empty;
                AutomationElement aeForm = null;

                // install Epia   UIAutomation
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);

                System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                Thread.Sleep(9000);
                aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);

                if (aeForm == null)
                {
                    logger.LogMessageToFile("aeForm  not found : ", sLogCount, sLogInterval);
                    System.Windows.MessageBox.Show("aeForm  not found : ");
                    return false;
                }
                else
                {
                    SetupWindowName = aeForm.Current.Name;
                    //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                    Log("aeForm found name : " + aeForm.Current.Name);
                    logger.LogMessageToFile("aeForm found name : " + aeForm.Current.Name, sLogCount, sLogInterval);

                    if (SetupWindowName.ToLower().EndsWith("windows installer"))
                    {
                        // check if only have one Cancel button---> do nothing
                        // else removing    
                        while (WindowHasOnlyThisButton(aeForm, "Cancel"))
                        {
                            logger.LogMessageToFile(aeForm.Current.Name + "< ONLY HAS CANCEL BUTTON> - Do nothing - ", sLogCount, sLogInterval);
                            Thread.Sleep(5000);
                        }

                        aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
                        SetupWindowName = aeForm.Current.Name;
                        logger.LogMessageToFile("New aeForm name is: " + aeForm.Current.Name, sLogCount, sLogInterval);
                        /*else
                        {
                            logger.LogMessageToFile("@@@@@ : ", sLogCount, sLogInterval);
                            logger.LogMessageToFile("@@@@@  aeForm is windows installer and should remove previous first : ", sLogCount, sLogInterval);
                            aeForm = RemovePreviousSetupandInstallCurrentEpia(aeForm);
                            SetupWindowName = aeForm.Current.Name;
                        }*/
                    }

                    Thread.Sleep(5000);

                }

                AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 6);
                //AutomationElement aeNextButton = AUIUtilities.GetElementByNameProperty(aeForm, "Next >");
                if (aeNextButton == null)
                {
                    // if Finished button found
                    // find "ControlType.RadioButton"    with name started with "Remove E'pia Framework" 
                    //  select remove radio button
                    // click Finish button
                    // wait until Close button found and click Close button
                    #region // check finish button and remove application
                    logger.LogMessageToFile("aeNextButton  not found : ", sLogCount, sLogInterval);
                    //MessageBox.Show("aeNextButton  not found : ");

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Finish", 5);
                    if (aeNextButton == null)
                        System.Windows.MessageBox.Show("aeNextButton Finish not found : ");
                    else
                    {
                        // find radio button
                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition c2 = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.RadioButton);

                        AutomationElementCollection aeAllRadioButtons = aeForm.FindAll(TreeScope.Element | TreeScope.Descendants, c2);
                        Thread.Sleep(1000);
                        foreach (AutomationElement s in aeAllRadioButtons)
                        {
                            if (s.Current.Name.StartsWith("Remove"))
                            {
                                SelectionItemPattern itemRadioPattern = s.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemRadioPattern.Select();
                                Thread.Sleep(3000);
                            }
                        }

                        System.Windows.Point FinishPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(FinishPt);
                        Thread.Sleep(1000);
                        Input.ClickAtPoint(FinishPt);
                        Thread.Sleep(2000);

                        logger.LogMessageToFile("<-----> Finish button clicked : ", sLogCount, sLogInterval);

                        // add event monitor and remove yes button
                        AutomationEventHandler UIRemoveRepairEventHandler = new AutomationEventHandler(OnRemoveRepairSetupScreenEvent);
                        // Add Open window Event Handler
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement, TreeScope.Descendants, UIRemoveRepairEventHandler);

                        // wait event end true
                        DateTime mStartTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - mStartTime;
                        Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                        //while (sEventEnd == false && mTime.Seconds <= 600)
                        //{
                        //    logger.LogMessageToFile("yes event is: "+sEventEnd +"<-> wait found yes button : " + mTime.Seconds , sLogCount, sLogInterval);
                        if (System.Environment.MachineName.IndexOf("TEAMTESTETRICC5") >= 0)
                        {
                            logger.LogMessageToFile("<-----> This is TEAMTESTETRICC  wait 35 sec: ", 0, 0);
                            Thread.Sleep(60000);
                        }
                        else
                            Thread.Sleep(10000);

                        //    mTime = DateTime.Now - mStartTime;
                        //}

                        // remove event monitor
                        Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                      AutomationElement.RootElement,
                                     UIRemoveRepairEventHandler);

                        //
                        // wait until close button found 
                        //
                        bool closeButtonfound = false;
                        int findCloseCnt = 0;
                        DateTime CloseButtonStartTime = DateTime.Now;
                        TimeSpan CloseButtonTime = DateTime.Now - CloseButtonStartTime;

                        while (closeButtonfound == false && findCloseCnt < 600)
                        {
                            //find install Window screen
                            System.Windows.Automation.Condition c2p = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, SetupWindowName),
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
                                System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
                            }
                            else
                            {
                                AutomationElement aeCloseButton = null;
                                StartTime = DateTime.Now;
                                Time = DateTime.Now - StartTime;
                                int wt = 0;
                                while (aeCloseButton == null && wt < 7)
                                {
                                    aeCloseButton = FindSpecificButtonByName(aeWindow, "Close");
                                    Time = DateTime.Now - StartTime;
                                    Thread.Sleep(2000);
                                    wt = wt + 2;
                                    logger.LogMessageToFile("<--XXX--> Close button not found yet: " + wt, sLogCount, sLogInterval);

                                }

                                if (aeCloseButton == null)
                                {
                                    //MessageBox.Show(" <--1111---> Close Button not found ;" + name, SetupStepDescriptions[0]);
                                    CloseButtonTime = DateTime.Now - CloseButtonStartTime;
                                    findCloseCnt++;
                                    logger.LogMessageToFile("<--2222---> Close button not found yet: " + findCloseCnt, sLogCount, sLogInterval);
                                    continue;
                                }
                                else
                                {
                                    closeButtonfound = true;
                                    System.Windows.Point ClosePt = AUIUtilities.GetElementCenterPoint(aeCloseButton);
                                    Thread.Sleep(1000);
                                    Input.MoveTo(ClosePt);
                                    Thread.Sleep(1000);
                                    Input.ClickAtPoint(ClosePt);
                                }
                            }
                        }
                    }
                    #endregion
                    // start msi again and find aeNextButton 
                    #region // restart again

                    // install Etricc 5   UIAutomation
                    Utilities.CloseProcess("msiexec");
                    Thread.Sleep(3000);
                    System.Diagnostics.Process SetupProc2 = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                    Thread.Sleep(2000);
                    aeForm = AutomationElement.FromHandle(SetupProc2.MainWindowHandle);

                    if (aeForm == null)
                    {
                        logger.LogMessageToFile("aeForm 2 not found : ", sLogCount, sLogInterval);
                        System.Windows.MessageBox.Show("aeForm 2 not found : ");
                    }
                    else
                    {
                        SetupWindowName = aeForm.Current.Name;
                        //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                        Log("aeForm found 2 name : " + aeForm.Current.Name);
                        logger.LogMessageToFile("aeForm found 2 name : " + aeForm.Current.Name, sLogCount, sLogInterval);
                        Thread.Sleep(5000);

                    }

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 10);

                    #endregion
                }

                if (aeNextButton == null)
                    System.Windows.Forms.MessageBox.Show("next button not found", "After remove, Reinstall App");
                else
                {
                    System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                    Thread.Sleep(2000);
                    Input.MoveTo(OptionPt);
                    Thread.Sleep(2000);
                    logger.LogMessageToFile("1<--->Next Button clicking ... ", sLogCount, sLogInterval);
                    Input.ClickAtPoint(OptionPt);
                    Thread.Sleep(2000);

                    //check is window visual state is minimized
                    // niot work aeForm is not a window review later
                    /*WindowPattern windowPattern = aeForm.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                    if (windowPattern.Current.WindowVisualState == WindowVisualState.Minimized)
                    {
                        System.Windows.Forms.MessageBox.Show("windowPattern.SetWindowVisualState is Minimized", " before InstallEpiaSetupByStep2");
                        windowPattern.SetWindowVisualState(WindowVisualState.Normal);
                    }
                    */
                    
                    /*for (int i = 2; i < 9; i++)
                    {
                        Thread.Sleep(4000);
                        logger.LogMessageToFile(" (" + i + ") " + SetupStepDescriptions[i], sLogCount, sLogInterval);
                        InstallEpiaSetupByStep(SetupWindowName, SetupStepDescriptions[i], i);
                    }
                    */
                    int i = 2;
                    while (InstallEpiaSetupByStep2(SetupWindowName, SetupStepDescriptions[i], i) == false)
                    {
                        i++;
                    }

                    installed = true;
                }
                #endregion

            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                System.Windows.Forms.MessageBox.Show(msg, "InstallSetup");
                throw new Exception(msg + "  during <InstallSetup>");
                // WIP log exception and return 
            }

            //Log("End Install Setup " + msiName + " at " + FilePath);


            if (installed)
            {
                try
                {
                    //save the name and the path in the registry to remove the setup
                    WriteInstallationToReg(Constants.EPIA4, FilePath, msiName);

                    // save deployment log  
                    AutoDeploymentOutputLog(Constants.EPIA4, m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Epia Server");
                    // 
                    string testResultsFile = Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "TestResults.txt");
                    if (!File.Exists(testResultsFile))
                    {
                        logger.LogMessageToFile("create test result file:" + testResultsFile, 0, 0);
                    }
                    else // empty file
                    {
                        logger.LogMessageToFile("empty test result to:" + testResultsFile, 0, 0);
                        StreamWriter writer = File.CreateText(testResultsFile);
                        writer.Close();
                    }

                    StreamWriter writerWorker = File.CreateText(testResultsFile);
                    writerWorker.WriteLine("failed");
                    writerWorker.Close();

                    logger.LogMessageToFile("write testResults: " + Path.Combine(@"C:\EtriccTests", "TestResults.txt"), 0, 0);

                    logger.LogMessageToFile(" **************** ( END " + "*" + sEpia4InstallerName + " Deployment )************************** ", 0, 0);

                }
                catch (Exception ex)
                {
                    string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                    System.Windows.Forms.MessageBox.Show(msg, "InstallSetupNA");
                    throw new Exception(msg + "  during <InstallSetup>");
                    // WIP log exception and return 
                }
            }

            return installed;

        }

        private bool InstallEtricc5Setup(string FilePath, string TestedVersion)
        {
            logger.LogMessageToFile("::: InstallEtricc5Setup : " + FilePath, sLogCount, sLogInterval);

            //find the msi in the filepath
            string msiName = string.Empty;
            DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
            FileInfo[] files = DirInfo.GetFiles(sEtriccMsiName);
            bool installed = false;

            TestedVersion = string.Empty;
            try
            {
                if (files[0] != null)
                    msiName = files[0].Name;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("cscript exception: " + ex.Message + " -- " + ex.StackTrace, "InstallSetup");
            }

            try
            {
                #region Install Etricc 5

                string[] SetupStepDescriptions = new string[100];
                SetupStepDescriptions[0] = "Welcome";
                SetupStepDescriptions[1] = "Welcome to the E'tricc 5.5.1.* Setup Wizard";
                SetupStepDescriptions[2] = "License Agreement";
                SetupStepDescriptions[3] = "Select components to install.";
                SetupStepDescriptions[4] = "Environment Configuration";
                SetupStepDescriptions[5] = "Select Installation Folder";
                SetupStepDescriptions[6] = "Confirm Installation";
                SetupStepDescriptions[7] = "Installing E'tricc 5.5.1.*";
                SetupStepDescriptions[8] = "Choose which functionality of the Launcher should be enabled.";
                SetupStepDescriptions[9] = "Security options";
                SetupStepDescriptions[10] = "Installation Complete";

                string SetupWindowName = string.Empty;
                string sErrorMessage = string.Empty;
                AutomationElement aeForm = null;

                // install Etricc 5   UIAutomation
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                Thread.Sleep(2000);
                aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);

                if (aeForm == null)
                {
                    logger.LogMessageToFile("aeForm  not found : ", sLogCount, sLogInterval);
                    System.Windows.MessageBox.Show("aeForm  not found : ");
                }
                else
                {
                    SetupWindowName = aeForm.Current.Name;
                    //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                    Log("aeForm found name : " + aeForm.Current.Name);
                    logger.LogMessageToFile("aeForm found name : " + aeForm.Current.Name, sLogCount, sLogInterval);

                    if (SetupWindowName.ToLower().EndsWith("windows installer"))
                    {
                        // check if only have one Cancel button---> do nothing
                        // else removing    
                        while (WindowHasOnlyThisButton(aeForm, "Cancel"))
                        {
                            logger.LogMessageToFile(aeForm.Current.Name + "< ONLY HAS CANCEL BUTTON> - Do nothing - ", sLogCount, sLogInterval);
                            Thread.Sleep(5000);
                        }

                        aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
                        SetupWindowName = aeForm.Current.Name;
                        logger.LogMessageToFile("New aeForm name is: " + aeForm.Current.Name, sLogCount, sLogInterval);
                        /*else
                        {
                            logger.LogMessageToFile("@@@@@ : ", sLogCount, sLogInterval);
                            logger.LogMessageToFile("@@@@@  aeForm is windows installer and should remove previous first : ", sLogCount, sLogInterval);
                            aeForm = RemovePreviousSetupandInstallCurrentEpia(aeForm);
                            SetupWindowName = aeForm.Current.Name;
                        }*/
                    }

                    Thread.Sleep(5000);

                }

                AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 7);
                if (aeNextButton == null)
                {
                    // if Finished button found
                    // find "ControlType.RadioButton"    with name started with "Remove E'tricc 5.5.1.*" 
                    //  select remove radio button
                    // click Finish button
                    // wait until Close button found and click Close button
                    #region // check finish button and remove application
                    logger.LogMessageToFile("aeNextButton  not found : ", sLogCount, sLogInterval);
                    //MessageBox.Show("aeNextButton  not found : ");

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Finish", 6);
                    if (aeNextButton == null)
                        System.Windows.MessageBox.Show("aeNextButton Finish not found : ");
                    else
                    {
                        // find radio button
                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition c2 = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.RadioButton);

                        AutomationElementCollection aeAllRadioButtons = aeForm.FindAll(TreeScope.Element | TreeScope.Descendants, c2);
                        Thread.Sleep(1000);
                        foreach (AutomationElement s in aeAllRadioButtons)
                        {
                            if (s.Current.Name.StartsWith("Remove"))
                            {
                                SelectionItemPattern itemRadioPattern = s.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemRadioPattern.Select();
                                Thread.Sleep(3000);
                            }
                        }

                        System.Windows.Point FinishPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(FinishPt);
                        Thread.Sleep(1000);
                        Input.ClickAtPoint(FinishPt);
                        Thread.Sleep(2000);

                        logger.LogMessageToFile("<-----> Finish button clicked : ", sLogCount, sLogInterval);

                        // add event monitor and remove yes button
                        AutomationEventHandler UIRemoveRepairEventHandler = new AutomationEventHandler(OnRemoveRepairSetupScreenEvent);
                        // Add Open window Event Handler
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement, TreeScope.Descendants, UIRemoveRepairEventHandler);

                        // wait event end true
                        DateTime mStartTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - mStartTime;
                        Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                        //while (sEventEnd == false && mTime.Seconds <= 600)
                        //{
                        //    logger.LogMessageToFile("yes event is: "+sEventEnd +"<-> wait found yes button : " + mTime.Seconds , sLogCount, sLogInterval);
                        if (System.Environment.MachineName.IndexOf("TEAMTESTETRICC5") >= 0)
                        {
                            logger.LogMessageToFile("<-----> This is TEAMTESTETRICC  wait 35 sec: ", 0, 0);
                            Thread.Sleep(60000);
                        }
                        else
                            Thread.Sleep(10000);

                        //    mTime = DateTime.Now - mStartTime;
                        //}

                        // remove event monitor
                        Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                      AutomationElement.RootElement,
                                     UIRemoveRepairEventHandler);

                        //
                        // wait until close button found 
                        //
                        bool closeButtonfound = false;
                        int findCloseCnt = 0;
                        DateTime CloseButtonStartTime = DateTime.Now;
                        TimeSpan CloseButtonTime = DateTime.Now - CloseButtonStartTime;

                        while (closeButtonfound == false && findCloseCnt < 600)
                        {
                            //find install Window screen
                            System.Windows.Automation.Condition c2p = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, SetupWindowName),
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
                                System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
                            }
                            else
                            {
                                AutomationElement aeCloseButton = null;
                                StartTime = DateTime.Now;
                                Time = DateTime.Now - StartTime;
                                int wt = 0;
                                while (aeCloseButton == null && wt < 7)
                                {
                                    aeCloseButton = FindSpecificButtonByName(aeWindow, "Close");
                                    Time = DateTime.Now - StartTime;
                                    Thread.Sleep(2000);
                                    wt = wt + 2;
                                    logger.LogMessageToFile("<--XXX--> Close button not found yet: " + wt, sLogCount, sLogInterval);

                                }

                                if (aeCloseButton == null)
                                {
                                    //MessageBox.Show(" <--1111---> Close Button not found ;" + name, SetupStepDescriptions[0]);
                                    CloseButtonTime = DateTime.Now - CloseButtonStartTime;
                                    findCloseCnt++;
                                    logger.LogMessageToFile("<--2222---> Close button not found yet: " + findCloseCnt, sLogCount, sLogInterval);
                                    continue;
                                }
                                else
                                {
                                    closeButtonfound = true;
                                    System.Windows.Point ClosePt = AUIUtilities.GetElementCenterPoint(aeCloseButton);
                                    Thread.Sleep(1000);
                                    Input.MoveTo(ClosePt);
                                    Thread.Sleep(1000);
                                    Input.ClickAtPoint(ClosePt);
                                }
                            }
                        }
                    }
                    #endregion
                    // start msi again and find aeNextButton 
                    #region // restart again

                    // install Etricc 5   UIAutomation
                    Utilities.CloseProcess("msiexec");
                    Thread.Sleep(3000);
                    System.Diagnostics.Process SetupProc2 = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                    Thread.Sleep(2000);
                    aeForm = AutomationElement.FromHandle(SetupProc2.MainWindowHandle);

                    if (aeForm == null)
                    {
                        logger.LogMessageToFile("aeForm  not found : ", sLogCount, sLogInterval);
                        System.Windows.MessageBox.Show("aeForm  not found : ");
                    }
                    else
                    {
                        SetupWindowName = aeForm.Current.Name;
                        //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                        Log("aeForm found name : " + aeForm.Current.Name);
                        logger.LogMessageToFile("aeForm found name : " + aeForm.Current.Name, sLogCount, sLogInterval);
                        Thread.Sleep(5000);

                    }

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 11);

                    #endregion
                }

                if (aeNextButton == null)
                    System.Windows.Forms.MessageBox.Show("next button not found", "After remove, Reinstall App");
                else
                {
                    System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                    Thread.Sleep(2000);
                    Input.MoveTo(OptionPt);
                    Thread.Sleep(2000);
                    logger.LogMessageToFile("<--->Next Button clicking ... ", sLogCount, sLogInterval);
                    Input.ClickAtPoint(OptionPt);
                    Thread.Sleep(2000);

                    if (InstallEtricc5SetupByStep(ref sErrorMessage) == false)
                    {
                        System.Windows.Forms.MessageBox.Show(sErrorMessage, "InstallEtricc  Setup ");
                        throw new Exception(sErrorMessage + "  during <InstallEtriccSetup>");
                    }

                    installed = true;
                }
                #endregion
            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + System.Environment.NewLine + ex.StackTrace;
                System.Windows.Forms.MessageBox.Show(msg, "InstallSetup ");
                throw new Exception(msg + "  during <InstallSetup>");
                // WIP log exception and return 
            }

            //Log("End Install Setup " + msiName + " at " + FilePath);


            if (installed)
            {
                //save the name and the path in the registry to remove the setup
                WriteInstallationToReg(Constants.ETRICC5, FilePath, msiName);

                // save deployment log
                AutoDeploymentOutputLog(mTestApp, m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Etricc Server");
                // 
                string testResultsFile = Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "TestResults.txt");
                if (!File.Exists(testResultsFile))
                {
                    logger.LogMessageToFile("create test result file:" + testResultsFile, 0, 0);
                }
                else // empty file
                {
                    logger.LogMessageToFile("empty test result to:" + testResultsFile, 0, 0);
                    StreamWriter writer = File.CreateText(testResultsFile);
                    writer.Close();
                }

                StreamWriter writerWorker = File.CreateText(testResultsFile);
                writerWorker.WriteLine("failed");
                writerWorker.Close();

                logger.LogMessageToFile("write testResults: " + Path.Combine(@"C:\EtriccTests", "TestResults.txt"), 0, 0);

                logger.LogMessageToFile(" **************** ( END Etricc ?.msi Deployment )************************** ", 0, 0);
            }

            return installed;

        }

        private string RecompileTestRuns()
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

                string origPath = sEtricc5InstallationFolder + "Egemin*.dll";
                string destPath = ConstCommon.ETRICC_TESTS_DIRECTORY + @"\TestRuns\bin\debug\";
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
                // DOTNET Version 2
                if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))  //  Deployed Testing
                {
                    arg = "/debug /target:library /out:" + Qmark + ConstCommon.ETRICC_TESTS_DIRECTORY + @"\TestRuns\bin\debug\TestRuns.dll" + Qmark;
                    arg = arg + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestWorker.cs" + Qmark
                         + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestProjectWorker.cs" + Qmark
                         + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestConstants.cs" + Qmark
                         + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\Logger.cs" + Qmark
                         + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestData.cs" + Qmark
                         + space + Qmark + m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + @"\TestUtility.cs" + Qmark;
                    arg = arg + space + "/reference:";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Design.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Interfaces.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Security.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Security.SSPI.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.UI.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Definitions.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.WCS.dll" + '"' + ";";
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
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Design.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Interfaces.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Security.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.Security.SSPI.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Common.UI.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.Definitions.dll" + '"' + ";";
                    arg = arg + '"' + dllPath + "Egemin.EPIA.WCS.dll" + '"' + ";";
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
                
                string DotnetVersionPath = m_SystemDrive + @"WINDOWS\Microsoft.NET\Framework\"+Constants.sRecompileDotnetVersion;
                string exePath = Path.Combine(DotnetVersionPath, "csc.exe");

                //System.Windows.Forms.MessageBox.Show("  exePath :" + exePath, "RecompileTestRuns");
                logger.LogMessageToFile("  exePath :" + exePath, 0, 0);

                // Run recompile Process
                output = Utilities.RunProcessAndGetOutput(exePath, arg);
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

        private bool InstallEtriccUISetup(string FilePath, string TestedVersion)
        {
            //MessageBox.Show("InstallEtriccSetup FilePath: " + FilePath);
            logger.LogMessageToFile("::: InstallEtriccUISetup : " + FilePath, sLogCount, sLogInterval);

            //find the msi in the filepath
            string msiName = string.Empty;
            DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
            FileInfo[] files = DirInfo.GetFiles("*Shell.msi");
            bool installed = false;

            TestedVersion = string.Empty;
            try
            {
                if (files[0] != null)
                    msiName = files[0].Name;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("InstallEtriccUISetup exception : " + ex.Message + " -- " + ex.StackTrace, "InstallEtriccUISetup(1) with filepath:" + FilePath);
            }

            try
            {
                #region Install EtriccUI

                string[] SetupStepDescriptions = new string[100];
                SetupStepDescriptions[0] = "Welcome";
                SetupStepDescriptions[1] = "Welcome to the E'tricc UI 2010.05.11.1 Setup Wizard";
                SetupStepDescriptions[2] = "Components";
                SetupStepDescriptions[3] = "Installation Folders";
                SetupStepDescriptions[4] = "Confirm Installation";
                SetupStepDescriptions[5] = "Installing E'pia Framework ...";
                SetupStepDescriptions[6] = "Installation Complete";

                string SetupWindowName = string.Empty;
                string sErrorMessage = string.Empty;
                AutomationElement aeForm = null;

                // install Etricc   UIAutomation
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                Thread.Sleep(2000);

                aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);

                if (aeForm == null)
                {
                    logger.LogMessageToFile("aeForm  not found : ", sLogCount, sLogInterval);
                    System.Windows.MessageBox.Show("aeForm  not found : ");
                }
                else
                {
                    SetupWindowName = aeForm.Current.Name;
                    //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                    Log("aeForm found name : " + SetupWindowName);
                    logger.LogMessageToFile("aeForm found name : " + SetupWindowName, sLogCount, sLogInterval);

                    if (SetupWindowName.ToLower().EndsWith("windows installer"))
                    {
                        // check if only have one Cancel button---> do nothing
                        // else removing    
                        while (WindowHasOnlyThisButton(aeForm, "Cancel"))
                        {
                            logger.LogMessageToFile(aeForm.Current.Name + "< ONLY HAS CANCEL BUTTON> - Do nothing - ", sLogCount, sLogInterval);
                            Thread.Sleep(5000);
                        }

                        aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
                        SetupWindowName = aeForm.Current.Name;
                        logger.LogMessageToFile("New aeForm name is: " + SetupWindowName, sLogCount, sLogInterval);
                        /*else
                        {
                            logger.LogMessageToFile("@@@@@ : ", sLogCount, sLogInterval);
                            logger.LogMessageToFile("@@@@@  aeForm is windows installer and should remove previous first : ", sLogCount, sLogInterval);
                            aeForm = RemovePreviousSetupandInstallCurrentEpia(aeForm);
                            SetupWindowName = aeForm.Current.Name;
                        }*/
                    }

                    Thread.Sleep(5000);

                }

                AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 8);
                if (aeNextButton == null)
                {
                    // if Finished button found
                    // find "ControlType.RadioButton"    with name started with "Remove E'pia Framework" 
                    //  select remove radio button
                    // click Finish button
                    // wait until Close button found and click Close button
                    #region // check finish button and remove application
                    logger.LogMessageToFile("aeNextButton  not found : ", sLogCount, sLogInterval);
                    //MessageBox.Show("aeNextButton  not found : ");

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Finish", 7);
                    if (aeNextButton == null)
                        System.Windows.MessageBox.Show("aeNextButton Finish not found : ");
                    else
                    {
                        // find radio button
                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition c2 = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.RadioButton);

                        AutomationElementCollection aeAllRadioButtons = aeForm.FindAll(TreeScope.Element | TreeScope.Descendants, c2);
                        Thread.Sleep(1000);
                        foreach (AutomationElement s in aeAllRadioButtons)
                        {
                            if (s.Current.Name.StartsWith("Remove"))
                            {
                                SelectionItemPattern itemRadioPattern = s.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemRadioPattern.Select();
                                Thread.Sleep(3000);
                            }
                        }

                        System.Windows.Point FinishPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(FinishPt);
                        Thread.Sleep(1000);
                        Input.ClickAtPoint(FinishPt);
                        Thread.Sleep(2000);

                        logger.LogMessageToFile("<-----> Finish button clicked : ", sLogCount, sLogInterval);

                        // add event monitor and remove yes button
                        AutomationEventHandler UIRemoveRepairEventHandler = new AutomationEventHandler(OnRemoveRepairSetupScreenEvent);
                        // Add Open window Event Handler
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement, TreeScope.Descendants, UIRemoveRepairEventHandler);

                        // wait event end true
                        DateTime mStartTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - mStartTime;
                        Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                        //while (sEventEnd == false && mTime.Seconds <= 600)
                        //{
                        //    logger.LogMessageToFile("yes event is: "+sEventEnd +"<-> wait found yes button : " + mTime.Seconds , sLogCount, sLogInterval);
                        if (System.Environment.MachineName.IndexOf("TEAMTESTETRICC5") >= 0)
                        {
                            logger.LogMessageToFile("<-----> This is TEAMTESTETRICC  wait 35 sec: ", 0, 0);
                            Thread.Sleep(60000);
                        }
                        else
                            Thread.Sleep(10000);

                        //    mTime = DateTime.Now - mStartTime;
                        //}

                        // remove event monitor
                        Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                      AutomationElement.RootElement,
                                     UIRemoveRepairEventHandler);

                        //
                        // wait until close button found 
                        //
                        bool closeButtonfound = false;
                        int findCloseCnt = 0;
                        DateTime CloseButtonStartTime = DateTime.Now;
                        TimeSpan CloseButtonTime = DateTime.Now - CloseButtonStartTime;

                        while (closeButtonfound == false && findCloseCnt < 600)
                        {
                            //find install Window screen
                            System.Windows.Automation.Condition c2p = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, SetupWindowName),
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
                                System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
                            }
                            else
                            {
                                AutomationElement aeCloseButton = null;
                                StartTime = DateTime.Now;
                                Time = DateTime.Now - StartTime;
                                int wt = 0;
                                while (aeCloseButton == null && wt < 7)
                                {
                                    aeCloseButton = FindSpecificButtonByName(aeWindow, "Close");
                                    Time = DateTime.Now - StartTime;
                                    Thread.Sleep(2000);
                                    wt = wt + 2;
                                    logger.LogMessageToFile("<--XXX--> Close button not found yet: " + wt, sLogCount, sLogInterval);

                                }

                                if (aeCloseButton == null)
                                {
                                    //MessageBox.Show(" <--1111---> Close Button not found ;" + name, SetupStepDescriptions[0]);
                                    CloseButtonTime = DateTime.Now - CloseButtonStartTime;
                                    findCloseCnt++;
                                    logger.LogMessageToFile("<--2222---> Close button not found yet: " + findCloseCnt, sLogCount, sLogInterval);
                                    continue;
                                }
                                else
                                {
                                    closeButtonfound = true;
                                    System.Windows.Point ClosePt = AUIUtilities.GetElementCenterPoint(aeCloseButton);
                                    Thread.Sleep(1000);
                                    Input.MoveTo(ClosePt);
                                    Thread.Sleep(1000);
                                    Input.ClickAtPoint(ClosePt);
                                }
                            }
                        }
                    }
                    #endregion
                    // start msi again and find aeNextButton 
                    #region // restart again

                    // install Etricc UIAutomation
                    Utilities.CloseProcess("msiexec");
                    Thread.Sleep(3000);
                    System.Diagnostics.Process SetupProc2 = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                    Thread.Sleep(2000);
                    aeForm = AutomationElement.FromHandle(SetupProc2.MainWindowHandle);

                    if (aeForm == null)
                    {
                        logger.LogMessageToFile("aeForm  not found : ", sLogCount, sLogInterval);
                        System.Windows.MessageBox.Show("aeForm  not found : ");
                    }
                    else
                    {
                        SetupWindowName = aeForm.Current.Name;
                        //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                        Log("aeForm found name : " + aeForm.Current.Name);
                        logger.LogMessageToFile("aeForm found name : " + aeForm.Current.Name, sLogCount, sLogInterval);
                        Thread.Sleep(5000);

                    }

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 12);

                    #endregion
                }

                if (aeNextButton == null)
                    System.Windows.Forms.MessageBox.Show("next button not found", "After remove, Reinstall App");
                else
                {
                    System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                    Thread.Sleep(2000);
                    Input.MoveTo(OptionPt);
                    Thread.Sleep(2000);
                    logger.LogMessageToFile("<--->Next Button clicking ... ", sLogCount, sLogInterval);
                    Input.ClickAtPoint(OptionPt);
                    Thread.Sleep(2000);

                    for (int i = 2; i < 8; i++)
                    {
                        Thread.Sleep(4000);
                        logger.LogMessageToFile(" (" + i + ") " + SetupStepDescriptions[i], sLogCount, sLogInterval);
                        InstallEtriccUISetupByStep(SetupWindowName, SetupStepDescriptions[i], i);
                    }

                    installed = true;
                }
                #endregion
            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + System.Environment.NewLine + ex.StackTrace;
                System.Windows.Forms.MessageBox.Show(msg, "InstallSetup");
                throw new Exception(msg + "  during <InstallSetup>");
                // WIP log exception and return 
            }

            //Log("End Install Setup " + msiName + " at " + FilePath);


            if (installed)
            {
                //save the name and the path in the registry to remove the setup
                WriteInstallationToReg(Constants.ETRICCUI, FilePath, msiName);

                // save deployment log
                AutoDeploymentOutputLog(mTestApp, m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Etricc Server ");
                // 
                string testResultsFile = Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "TestResults.txt");
                if (!File.Exists(testResultsFile))
                {
                    logger.LogMessageToFile("create test result file:" + testResultsFile, 0, 0);
                }
                else // empty file
                {
                    logger.LogMessageToFile("empty test result to:" + testResultsFile, 0, 0);
                    StreamWriter writer = File.CreateText(testResultsFile);
                    writer.Close();
                }

                StreamWriter writerWorker = File.CreateText(testResultsFile);
                writerWorker.WriteLine("failed");
                writerWorker.Close();

                logger.LogMessageToFile("write testResults: " + Path.Combine(@"C:\EtriccTests", "TestResults.txt"), 0, 0);

                logger.LogMessageToFile(" **************** ( END Etricc UI.msi Deployment )************************** ", 0, 0);
            }

            return installed;

        }

        private bool InstallEtriccStatisticsParserSetup(string FilePath)
        {
            //MessageBox.Show("InstallEtriccStatisticsParserSetup FilePath: " + FilePath);
            logger.LogMessageToFile("::: InstallEtriccStatisticsParserSetup : " + FilePath, sLogCount, sLogInterval);

            //find the msi in the filepath
            string msiName = string.Empty;
            DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
            FileInfo[] files = DirInfo.GetFiles("Etricc.Statistics.Parser.msi");
            bool installed = false;

            try
            {
                if (files[0] != null)
                    msiName = files[0].Name;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("cscript exception: " + ex.Message + " -- " + ex.StackTrace, "InstallEtriccStatisticsParserSetup");
            }

            try
            {
                #region Install Etricc statistics parser

                string[] SetupStepDescriptions = new string[100];
                SetupStepDescriptions[0] = "Welcome";
                SetupStepDescriptions[1] = "Welcome to the E'tricc Statistics Parser 201!.05.12.* Setup Wizard";
                SetupStepDescriptions[2] = "Select Installation Folder";
                SetupStepDescriptions[3] = "Confirm Installation";
                SetupStepDescriptions[4] = "Installing E'tricc Statistics Parser ...";
                SetupStepDescriptions[5] = "Installation Complete";

                string SetupWindowName = string.Empty;
                string sErrorMessage = string.Empty;
                AutomationElement aeForm = null;

                // install Epia   UIAutomation
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);

                System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                Thread.Sleep(9000);
                aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);

                if (aeForm == null)
                {
                    logger.LogMessageToFile("aeForm EtriccStatisticsParser not found : ", sLogCount, sLogInterval);
                    System.Windows.MessageBox.Show("aeForm EtriccStatisticsParser not found : ");
                    return false;
                }
                else
                {
                    SetupWindowName = aeForm.Current.Name;
                    //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                    Log("aeForm EtriccStatisticsParser found name : " + aeForm.Current.Name);
                    logger.LogMessageToFile("aeForm EtriccStatisticsParser found name : " + aeForm.Current.Name, sLogCount, sLogInterval);

                    if (SetupWindowName.ToLower().EndsWith("windows installer"))
                    {
                        // check if only have one Cancel button---> do nothing
                        // else removing    
                        while (WindowHasOnlyThisButton(aeForm, "Cancel"))
                        {
                            logger.LogMessageToFile(aeForm.Current.Name + "< ONLY HAS CANCEL BUTTON> - Do nothing - ", sLogCount, sLogInterval);
                            Thread.Sleep(5000);
                        }

                        aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
                        SetupWindowName = aeForm.Current.Name;
                        logger.LogMessageToFile("New aeForm EtriccStatisticsParser name is: " + aeForm.Current.Name, sLogCount, sLogInterval);
                        /*else
                        {
                            logger.LogMessageToFile("@@@@@ : ", sLogCount, sLogInterval);
                            logger.LogMessageToFile("@@@@@  aeForm is windows installer and should remove previous first : ", sLogCount, sLogInterval);
                            aeForm = RemovePreviousSetupandInstallCurrentEpia(aeForm);
                            SetupWindowName = aeForm.Current.Name;
                        }*/
                    }

                    Thread.Sleep(5000);

                }



                AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 6);
                //AutomationElement aeNextButton = AUIUtilities.GetElementByNameProperty(aeForm, "Next >");
                if (aeNextButton == null)
                {
                    // if Finished button found
                    // find "ControlType.RadioButton"    with name started with "Remove E'pia Framework" 
                    //  select remove radio button
                    // click Finish button
                    // wait until Close button found and click Close button
                    #region // check finish button and remove application
                    logger.LogMessageToFile("aeNextButton  not found : ", sLogCount, sLogInterval);
                    //MessageBox.Show("aeNextButton  not found : ");

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Finish", 5);
                    if (aeNextButton == null)
                        System.Windows.MessageBox.Show("aeNextButton Finish not found : ");
                    else
                    {
                        // find radio button
                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition c2 = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.RadioButton);

                        AutomationElementCollection aeAllRadioButtons = aeForm.FindAll(TreeScope.Element | TreeScope.Descendants, c2);
                        Thread.Sleep(1000);
                        foreach (AutomationElement s in aeAllRadioButtons)
                        {
                            if (s.Current.Name.StartsWith("Remove"))
                            {
                                SelectionItemPattern itemRadioPattern = s.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemRadioPattern.Select();
                                Thread.Sleep(3000);
                            }
                        }

                        System.Windows.Point FinishPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(FinishPt);
                        Thread.Sleep(1000);
                        Input.ClickAtPoint(FinishPt);
                        Thread.Sleep(2000);

                        logger.LogMessageToFile("<-----> Finish button clicked : ", sLogCount, sLogInterval);

                        // add event monitor and remove yes button
                        AutomationEventHandler UIRemoveRepairEventHandler = new AutomationEventHandler(OnRemoveRepairSetupScreenEvent);
                        // Add Open window Event Handler
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement, TreeScope.Descendants, UIRemoveRepairEventHandler);

                        // wait event end true
                        DateTime mStartTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - mStartTime;
                        Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                        //while (sEventEnd == false && mTime.Seconds <= 600)
                        //{
                        //    logger.LogMessageToFile("yes event is: "+sEventEnd +"<-> wait found yes button : " + mTime.Seconds , sLogCount, sLogInterval);
                        if (System.Environment.MachineName.IndexOf("TEAMTESTETRICC5") >= 0)
                        {
                            logger.LogMessageToFile("<-----> This is TEAMTESTETRICC  wait 35 sec: ", 0, 0);
                            Thread.Sleep(60000);
                        }
                        else
                            Thread.Sleep(10000);

                        //    mTime = DateTime.Now - mStartTime;
                        //}

                        // remove event monitor
                        Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                      AutomationElement.RootElement,
                                     UIRemoveRepairEventHandler);

                        //
                        // wait until close button found 
                        //
                        bool closeButtonfound = false;
                        int findCloseCnt = 0;
                        DateTime CloseButtonStartTime = DateTime.Now;
                        TimeSpan CloseButtonTime = DateTime.Now - CloseButtonStartTime;

                        while (closeButtonfound == false && findCloseCnt < 600)
                        {
                            //find install Window screen
                            System.Windows.Automation.Condition c2p = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, SetupWindowName),
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
                                System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
                            }
                            else
                            {
                                AutomationElement aeCloseButton = null;
                                StartTime = DateTime.Now;
                                Time = DateTime.Now - StartTime;
                                int wt = 0;
                                while (aeCloseButton == null && wt < 7)
                                {
                                    aeCloseButton = FindSpecificButtonByName(aeWindow, "Close");
                                    Time = DateTime.Now - StartTime;
                                    Thread.Sleep(2000);
                                    wt = wt + 2;
                                    logger.LogMessageToFile("<--XXX--> Close button not found yet: " + wt, sLogCount, sLogInterval);

                                }

                                if (aeCloseButton == null)
                                {
                                    //MessageBox.Show(" <--1111---> Close Button not found ;" + name, SetupStepDescriptions[0]);
                                    CloseButtonTime = DateTime.Now - CloseButtonStartTime;
                                    findCloseCnt++;
                                    logger.LogMessageToFile("<--2222---> Close button not found yet: " + findCloseCnt, sLogCount, sLogInterval);
                                    continue;
                                }
                                else
                                {
                                    closeButtonfound = true;
                                    System.Windows.Point ClosePt = AUIUtilities.GetElementCenterPoint(aeCloseButton);
                                    Thread.Sleep(1000);
                                    Input.MoveTo(ClosePt);
                                    Thread.Sleep(1000);
                                    Input.ClickAtPoint(ClosePt);
                                }
                            }
                        }
                    }
                    #endregion
                    // start msi again and find aeNextButton 
                    #region // restart again

                    // install EtriccStatisticsParser   UIAutomation
                    Utilities.CloseProcess("msiexec");
                    Thread.Sleep(3000);
                    System.Diagnostics.Process SetupProc2 = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                    Thread.Sleep(2000);
                    aeForm = AutomationElement.FromHandle(SetupProc2.MainWindowHandle);

                    if (aeForm == null)
                    {
                        logger.LogMessageToFile("aeForm  not found : ", sLogCount, sLogInterval);
                        System.Windows.MessageBox.Show("aeForm  not found : ");
                    }
                    else
                    {
                        SetupWindowName = aeForm.Current.Name;
                        //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                        Log("aeForm found name : " + aeForm.Current.Name);
                        logger.LogMessageToFile("aeForm found name : " + aeForm.Current.Name, sLogCount, sLogInterval);
                        Thread.Sleep(5000);

                    }

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 10);

                    #endregion
                }

                if (aeNextButton == null)
                    System.Windows.Forms.MessageBox.Show("next button not found", "After remove, Reinstall App");
                else
                {
                    System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                    Thread.Sleep(2000);
                    Input.MoveTo(OptionPt);
                    Thread.Sleep(2000);
                    logger.LogMessageToFile("<--->Next Button clicking on Welcome screen... ", sLogCount, sLogInterval);
                    Input.ClickAtPoint(OptionPt);
                    Thread.Sleep(2000);


                    int i = 2;
                    while (InstallEtriccStatisticsParserSetupByStep(SetupWindowName, SetupStepDescriptions[i], i) == false)
                    {
                        i++;
                    }

                    installed = true;
                }
                #endregion

            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                System.Windows.Forms.MessageBox.Show(msg, "InstallSetup");
                throw new Exception(msg + "  during <InstallSetup>");
                // WIP log exception and return 
            }

            //Log("End Install Setup " + msiName + " at " + FilePath);


            if (installed)
            {
                try
                {
                    //save the name and the path in the registry to remove the setup
                    WriteInstallationToReg(Constants.ETRICCSTATISTICS_PARSER_SETUP, FilePath, msiName);

                    // save deployment log  
                    AutoDeploymentOutputLog(Constants.ETRICCSTATISTICS_PARSER_SETUP, m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Etricc Statistics Parser");
                    // 
                    string testResultsFile = Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "TestResults.txt");
                    if (!File.Exists(testResultsFile))
                    {
                        logger.LogMessageToFile("create test result file:" + testResultsFile, 0, 0);
                    }
                    else // empty file
                    {
                        logger.LogMessageToFile("empty test result to:" + testResultsFile, 0, 0);
                        StreamWriter writer = File.CreateText(testResultsFile);
                        writer.Close();
                    }

                    StreamWriter writerWorker = File.CreateText(testResultsFile);
                    writerWorker.WriteLine("failed");
                    writerWorker.Close();

                    logger.LogMessageToFile("write testResults: " + Path.Combine(@"C:\EtriccTests", "TestResults.txt"), 0, 0);

                    logger.LogMessageToFile(" **************** ( END " + "Etricc.Statistics.Parser.msi Deployment )************************** ", 0, 0);

                }
                catch (Exception ex)
                {
                    string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                    System.Windows.Forms.MessageBox.Show(msg, "InstallEtriccStatisticsParserSetup");
                    throw new Exception(msg + "  during <InstallEtriccStatisticsParserSetup>");
                    // WIP log exception and return 
                }
            }

            return installed;

        }

        private bool InstallEtriccStatisticsParserConfiguratorSetup(string FilePath)
        {
            //MessageBox.Show("InstallEtriccStatisticsParserSetup FilePath: " + FilePath);
            logger.LogMessageToFile("::: InstallEtriccStatisticsParserConfiguratorSetup : " + FilePath, sLogCount, sLogInterval);

            //find the msi in the filepath
            string msiName = string.Empty;
            DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
            FileInfo[] files = DirInfo.GetFiles("Etricc.Statistics.ParserConfigurator.msi");
            bool installed = false;

            try
            {
                if (files[0] != null)
                    msiName = files[0].Name;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("cscript exception: " + ex.Message + " -- " + ex.StackTrace, "InstallEtriccStatisticsParserConfiguratorSetup");
            }

            try
            {
                #region Install Epia Old version

                string[] SetupStepDescriptions = new string[100];
                SetupStepDescriptions[0] = "Welcome";
                SetupStepDescriptions[1] = "Welcome to the E'tricc Statistics ParserConfiguration 2011.05.12.* Setup Wizard";
                SetupStepDescriptions[2] = "Select Installation Folder";
                SetupStepDescriptions[3] = "Confirm Installation";
                SetupStepDescriptions[4] = "Installing E'tricc Statistics ParserConfiguration ...";
                SetupStepDescriptions[5] = "Installation Complete";

                string SetupWindowName = string.Empty;
                string sErrorMessage = string.Empty;
                AutomationElement aeForm = null;

                // install Epia   UIAutomation
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);

                System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                Thread.Sleep(9000);
                aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);

                if (aeForm == null)
                {
                    logger.LogMessageToFile("aeForm EtriccStatisticsParserConfiguration not found : ", sLogCount, sLogInterval);
                    System.Windows.MessageBox.Show("aeForm EtriccStatisticsParserConfiguration not found : ");
                    return false;
                }
                else
                {
                    SetupWindowName = aeForm.Current.Name;
                    //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                    Log("aeForm EtriccStatisticsParserConfiguration found name : " + aeForm.Current.Name);
                    logger.LogMessageToFile("aeForm EtriccStatisticsParserConfiguration found name : " + aeForm.Current.Name, sLogCount, sLogInterval);

                    if (SetupWindowName.ToLower().EndsWith("windows installer"))
                    {
                        // check if only have one Cancel button---> do nothing
                        // else removing    
                        while (WindowHasOnlyThisButton(aeForm, "Cancel"))
                        {
                            logger.LogMessageToFile(aeForm.Current.Name + "< ONLY HAS CANCEL BUTTON> - Do nothing - ", sLogCount, sLogInterval);
                            Thread.Sleep(5000);
                        }

                        aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
                        SetupWindowName = aeForm.Current.Name;
                        logger.LogMessageToFile("New aeForm EtriccStatisticsParser name is: " + aeForm.Current.Name, sLogCount, sLogInterval);
                        /*else
                        {
                            logger.LogMessageToFile("@@@@@ : ", sLogCount, sLogInterval);
                            logger.LogMessageToFile("@@@@@  aeForm is windows installer and should remove previous first : ", sLogCount, sLogInterval);
                            aeForm = RemovePreviousSetupandInstallCurrentEpia(aeForm);
                            SetupWindowName = aeForm.Current.Name;
                        }*/
                    }

                    Thread.Sleep(5000);

                }



                AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 6);
                //AutomationElement aeNextButton = AUIUtilities.GetElementByNameProperty(aeForm, "Next >");
                if (aeNextButton == null)
                {
                    // if Finished button found
                    // find "ControlType.RadioButton"    with name started with "Remove E'pia Framework" 
                    //  select remove radio button
                    // click Finish button
                    // wait until Close button found and click Close button
                    #region // check finish button and remove application
                    logger.LogMessageToFile("aeNextButton  not found : ", sLogCount, sLogInterval);
                    //MessageBox.Show("aeNextButton  not found : ");

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Finish", 5);
                    if (aeNextButton == null)
                        System.Windows.MessageBox.Show("aeNextButton Finish not found : ");
                    else
                    {
                        // find radio button
                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition c2 = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.RadioButton);

                        AutomationElementCollection aeAllRadioButtons = aeForm.FindAll(TreeScope.Element | TreeScope.Descendants, c2);
                        Thread.Sleep(1000);
                        foreach (AutomationElement s in aeAllRadioButtons)
                        {
                            if (s.Current.Name.StartsWith("Remove"))
                            {
                                SelectionItemPattern itemRadioPattern = s.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemRadioPattern.Select();
                                Thread.Sleep(3000);
                            }
                        }

                        System.Windows.Point FinishPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(FinishPt);
                        Thread.Sleep(1000);
                        Input.ClickAtPoint(FinishPt);
                        Thread.Sleep(2000);

                        logger.LogMessageToFile("<-----> Finish button clicked : ", sLogCount, sLogInterval);

                        // add event monitor and remove yes button
                        AutomationEventHandler UIRemoveRepairEventHandler = new AutomationEventHandler(OnRemoveRepairSetupScreenEvent);
                        // Add Open window Event Handler
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement, TreeScope.Descendants, UIRemoveRepairEventHandler);

                        // wait event end true
                        DateTime mStartTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - mStartTime;
                        Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                        //while (sEventEnd == false && mTime.Seconds <= 600)
                        //{
                        //    logger.LogMessageToFile("yes event is: "+sEventEnd +"<-> wait found yes button : " + mTime.Seconds , sLogCount, sLogInterval);
                        if (System.Environment.MachineName.IndexOf("TEAMTESTETRICC5") >= 0)
                        {
                            logger.LogMessageToFile("<-----> This is TEAMTESTETRICC  wait 35 sec: ", 0, 0);
                            Thread.Sleep(60000);
                        }
                        else
                            Thread.Sleep(10000);

                        //    mTime = DateTime.Now - mStartTime;
                        //}

                        // remove event monitor
                        Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                      AutomationElement.RootElement,
                                     UIRemoveRepairEventHandler);

                        //
                        // wait until close button found 
                        //
                        bool closeButtonfound = false;
                        int findCloseCnt = 0;
                        DateTime CloseButtonStartTime = DateTime.Now;
                        TimeSpan CloseButtonTime = DateTime.Now - CloseButtonStartTime;

                        while (closeButtonfound == false && findCloseCnt < 600)
                        {
                            //find install Window screen
                            System.Windows.Automation.Condition c2p = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, SetupWindowName),
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
                                System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
                            }
                            else
                            {
                                AutomationElement aeCloseButton = null;
                                StartTime = DateTime.Now;
                                Time = DateTime.Now - StartTime;
                                int wt = 0;
                                while (aeCloseButton == null && wt < 7)
                                {
                                    aeCloseButton = FindSpecificButtonByName(aeWindow, "Close");
                                    Time = DateTime.Now - StartTime;
                                    Thread.Sleep(2000);
                                    wt = wt + 2;
                                    logger.LogMessageToFile("<--XXX--> Close button not found yet: " + wt, sLogCount, sLogInterval);

                                }

                                if (aeCloseButton == null)
                                {
                                    //MessageBox.Show(" <--1111---> Close Button not found ;" + name, SetupStepDescriptions[0]);
                                    CloseButtonTime = DateTime.Now - CloseButtonStartTime;
                                    findCloseCnt++;
                                    logger.LogMessageToFile("<--2222---> Close button not found yet: " + findCloseCnt, sLogCount, sLogInterval);
                                    continue;
                                }
                                else
                                {
                                    closeButtonfound = true;
                                    System.Windows.Point ClosePt = AUIUtilities.GetElementCenterPoint(aeCloseButton);
                                    Thread.Sleep(1000);
                                    Input.MoveTo(ClosePt);
                                    Thread.Sleep(1000);
                                    Input.ClickAtPoint(ClosePt);
                                }
                            }
                        }
                    }
                    #endregion
                    // start msi again and find aeNextButton 
                    #region // restart again

                    // install EtriccStatisticsParser   UIAutomation
                    Utilities.CloseProcess("msiexec");
                    Thread.Sleep(3000);
                    System.Diagnostics.Process SetupProc2 = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                    Thread.Sleep(2000);
                    aeForm = AutomationElement.FromHandle(SetupProc2.MainWindowHandle);

                    if (aeForm == null)
                    {
                        logger.LogMessageToFile("aeForm  not found : ", sLogCount, sLogInterval);
                        System.Windows.MessageBox.Show("aeForm  not found : ");
                    }
                    else
                    {
                        SetupWindowName = aeForm.Current.Name;
                        //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                        Log("aeForm found name : " + aeForm.Current.Name);
                        logger.LogMessageToFile("aeForm found name : " + aeForm.Current.Name, sLogCount, sLogInterval);
                        Thread.Sleep(5000);

                    }

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 10);

                    #endregion
                }

                if (aeNextButton == null)
                    System.Windows.Forms.MessageBox.Show("next button not found", "After remove, Reinstall App");
                else
                {
                    System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                    Thread.Sleep(2000);
                    Input.MoveTo(OptionPt);
                    Thread.Sleep(2000);
                    logger.LogMessageToFile("<--->Next Button clicking ... ", sLogCount, sLogInterval);
                    Input.ClickAtPoint(OptionPt);
                    Thread.Sleep(2000);

                    int i = 2;
                    while ( InstallEtriccStatisticsParserConfiguratorSetupByStep(SetupWindowName, SetupStepDescriptions[i], i) == false)
                    {
                        i++;
                    }

                    installed = true;
                }
                #endregion

            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                System.Windows.Forms.MessageBox.Show(msg, "InstallEtriccStatisticsParserConfiguratorSetup");
                throw new Exception(msg + "  during <InstallEtriccStatisticsParserConfiguratorSetup>");
                // WIP log exception and return 
            }

            //Log("End Install Setup " + msiName + " at " + FilePath);


            if (installed)
            {
                try
                {
                    //save the name and the path in the registry to remove the setup
                    WriteInstallationToReg(Constants.ETRICCSTATISTICS_PARSERCONFIGURATOR, FilePath, msiName);

                    // save deployment log  
                    AutoDeploymentOutputLog(Constants.ETRICCSTATISTICS_PARSERCONFIGURATOR, m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Etricc Statistics ParserConfigurator");
                    // 
                    string testResultsFile = Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "TestResults.txt");
                    if (!File.Exists(testResultsFile))
                    {
                        logger.LogMessageToFile("create test result file:" + testResultsFile, 0, 0);
                    }
                    else // empty file
                    {
                        logger.LogMessageToFile("empty test result to:" + testResultsFile, 0, 0);
                        StreamWriter writer = File.CreateText(testResultsFile);
                        writer.Close();
                    }

                    StreamWriter writerWorker = File.CreateText(testResultsFile);
                    writerWorker.WriteLine("failed");
                    writerWorker.Close();

                    logger.LogMessageToFile("write testResults: " + Path.Combine(@"C:\EtriccTests", "TestResults.txt"), 0, 0);

                    logger.LogMessageToFile(" **************** ( END " + "Etricc.Statistics.ParserConfigurator.msi Deployment )************************** ", 0, 0);

                }
                catch (Exception ex)
                {
                    string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                    System.Windows.Forms.MessageBox.Show(msg, "InstallEtriccStatisticsParserConfiguratorSetup");
                    throw new Exception(msg + "  during <InstallEtriccStatisticsParserConfiguratorSetup>");
                    // WIP log exception and return 
                }
            }

            return installed;

        }

        private bool InstallEtriccStatisticsUISetup(string FilePath)
        {
            //MessageBox.Show("InstallEtriccStatisticsUI FilePath: " + FilePath);
            logger.LogMessageToFile("::: InstallEtriccStatisticsUISetup : " + FilePath, sLogCount, sLogInterval);

            //find the msi in the filepath
            string msiName = string.Empty;
            DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
            FileInfo[] files = DirInfo.GetFiles("Etricc.Statistics.UI.msi");
            bool installed = false;

            try
            {
                if (files[0] != null)
                    msiName = files[0].Name;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("cscript exception: " + ex.Message + " -- " + ex.StackTrace, "InstallEtriccStatisticsUISetup");
            }

            try
            {
                #region Install Epia Old version

                string[] SetupStepDescriptions = new string[100];
                SetupStepDescriptions[0] = "Welcome";
                SetupStepDescriptions[1] = "Welcome to the E'tricc Statistics UI 2011.05.12.* Setup Wizard";
                SetupStepDescriptions[2] = "Components";
                SetupStepDescriptions[3] = "Installation Folder";
                SetupStepDescriptions[4] = "Confirm Installation";
                SetupStepDescriptions[5] = "Installing E'tricc Statistics ParserConfiguration ...";
                SetupStepDescriptions[6] = "Installation Complete";

                string SetupWindowName = string.Empty;
                string sErrorMessage = string.Empty;
                AutomationElement aeForm = null;

                // install Epia   UIAutomation
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);
                Utilities.CloseProcess("msiexec");
                Thread.Sleep(3000);

                System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                Thread.Sleep(9000);
                aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);

                if (aeForm == null)
                {
                    logger.LogMessageToFile("aeForm EtriccStatisticsUI not found : ", sLogCount, sLogInterval);
                    System.Windows.MessageBox.Show("aeForm EtriccStatisticsUI not found : ");
                    return false;
                }
                else
                {
                    SetupWindowName = aeForm.Current.Name;
                    //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                    Log("aeForm EtriccStatisticsUI found name : " + aeForm.Current.Name);
                    logger.LogMessageToFile("aeForm EtriccStatisticsUI found name : " + aeForm.Current.Name, sLogCount, sLogInterval);

                    if (SetupWindowName.ToLower().EndsWith("windows installer"))
                    {
                        // check if only have one Cancel button---> do nothing
                        // else removing    
                        while (WindowHasOnlyThisButton(aeForm, "Cancel"))
                        {
                            logger.LogMessageToFile(aeForm.Current.Name + "< ONLY HAS CANCEL BUTTON> - Do nothing - ", sLogCount, sLogInterval);
                            Thread.Sleep(5000);
                        }

                        aeForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
                        SetupWindowName = aeForm.Current.Name;
                        logger.LogMessageToFile("New aeForm EtriccStatisticsParser name is: " + aeForm.Current.Name, sLogCount, sLogInterval);
                        /*else
                        {
                            logger.LogMessageToFile("@@@@@ : ", sLogCount, sLogInterval);
                            logger.LogMessageToFile("@@@@@  aeForm is windows installer and should remove previous first : ", sLogCount, sLogInterval);
                            aeForm = RemovePreviousSetupandInstallCurrentEpia(aeForm);
                            SetupWindowName = aeForm.Current.Name;
                        }*/
                    }

                    Thread.Sleep(5000);

                }



                AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 6);
                //AutomationElement aeNextButton = AUIUtilities.GetElementByNameProperty(aeForm, "Next >");
                if (aeNextButton == null)
                {
                    // if Finished button found
                    // find "ControlType.RadioButton"    with name started with "Remove E'pia Framework" 
                    //  select remove radio button
                    // click Finish button
                    // wait until Close button found and click Close button
                    #region // check finish button and remove application
                    logger.LogMessageToFile("aeNextButton  not found : ", sLogCount, sLogInterval);
                    //MessageBox.Show("aeNextButton  not found : ");

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Finish", 5);
                    if (aeNextButton == null)
                        System.Windows.MessageBox.Show("aeNextButton Finish not found : ");
                    else
                    {
                        // find radio button
                        // Set a property condition that will be used to find the control.
                        System.Windows.Automation.Condition c2 = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, ControlType.RadioButton);

                        AutomationElementCollection aeAllRadioButtons = aeForm.FindAll(TreeScope.Element | TreeScope.Descendants, c2);
                        Thread.Sleep(1000);
                        foreach (AutomationElement s in aeAllRadioButtons)
                        {
                            if (s.Current.Name.StartsWith("Remove"))
                            {
                                SelectionItemPattern itemRadioPattern = s.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemRadioPattern.Select();
                                Thread.Sleep(3000);
                            }
                        }

                        System.Windows.Point FinishPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(FinishPt);
                        Thread.Sleep(1000);
                        Input.ClickAtPoint(FinishPt);
                        Thread.Sleep(2000);

                        logger.LogMessageToFile("<-----> Finish button clicked : ", sLogCount, sLogInterval);

                        // add event monitor and remove yes button
                        AutomationEventHandler UIRemoveRepairEventHandler = new AutomationEventHandler(OnRemoveRepairSetupScreenEvent);
                        // Add Open window Event Handler
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement, TreeScope.Descendants, UIRemoveRepairEventHandler);

                        // wait event end true
                        DateTime mStartTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - mStartTime;
                        Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                        //while (sEventEnd == false && mTime.Seconds <= 600)
                        //{
                        //    logger.LogMessageToFile("yes event is: "+sEventEnd +"<-> wait found yes button : " + mTime.Seconds , sLogCount, sLogInterval);
                        if (System.Environment.MachineName.IndexOf("TEAMTESTETRICC5") >= 0)
                        {
                            logger.LogMessageToFile("<-----> This is TEAMTESTETRICC  wait 35 sec: ", 0, 0);
                            Thread.Sleep(60000);
                        }
                        else
                            Thread.Sleep(10000);

                        //    mTime = DateTime.Now - mStartTime;
                        //}

                        // remove event monitor
                        Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                      AutomationElement.RootElement,
                                     UIRemoveRepairEventHandler);

                        //
                        // wait until close button found 
                        //
                        bool closeButtonfound = false;
                        int findCloseCnt = 0;
                        DateTime CloseButtonStartTime = DateTime.Now;
                        TimeSpan CloseButtonTime = DateTime.Now - CloseButtonStartTime;

                        while (closeButtonfound == false && findCloseCnt < 600)
                        {
                            //find install Window screen
                            System.Windows.Automation.Condition c2p = new AndCondition(
                              new PropertyCondition(AutomationElement.NameProperty, SetupWindowName),
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
                                System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
                            }
                            else
                            {
                                AutomationElement aeCloseButton = null;
                                StartTime = DateTime.Now;
                                Time = DateTime.Now - StartTime;
                                int wt = 0;
                                while (aeCloseButton == null && wt < 7)
                                {
                                    aeCloseButton = FindSpecificButtonByName(aeWindow, "Close");
                                    Time = DateTime.Now - StartTime;
                                    Thread.Sleep(2000);
                                    wt = wt + 2;
                                    logger.LogMessageToFile("<--XXX--> Close button not found yet: " + wt, sLogCount, sLogInterval);

                                }

                                if (aeCloseButton == null)
                                {
                                    //MessageBox.Show(" <--1111---> Close Button not found ;" + name, SetupStepDescriptions[0]);
                                    CloseButtonTime = DateTime.Now - CloseButtonStartTime;
                                    findCloseCnt++;
                                    logger.LogMessageToFile("<--2222---> Close button not found yet: " + findCloseCnt, sLogCount, sLogInterval);
                                    continue;
                                }
                                else
                                {
                                    closeButtonfound = true;
                                    System.Windows.Point ClosePt = AUIUtilities.GetElementCenterPoint(aeCloseButton);
                                    Thread.Sleep(1000);
                                    Input.MoveTo(ClosePt);
                                    Thread.Sleep(1000);
                                    Input.ClickAtPoint(ClosePt);
                                }
                            }
                        }
                    }
                    #endregion
                    // start msi again and find aeNextButton 
                    #region // restart again

                    // install EtriccStatisticsParser   UIAutomation
                    Utilities.CloseProcess("msiexec");
                    Thread.Sleep(3000);
                    System.Diagnostics.Process SetupProc2 = Utilities.StartProcessNoWait(FilePath, msiName, string.Empty);
                    Thread.Sleep(2000);
                    aeForm = AutomationElement.FromHandle(SetupProc2.MainWindowHandle);

                    if (aeForm == null)
                    {
                        logger.LogMessageToFile("aeForm  not found : ", sLogCount, sLogInterval);
                        System.Windows.MessageBox.Show("aeForm  not found : ");
                    }
                    else
                    {
                        SetupWindowName = aeForm.Current.Name;
                        //MessageBox.Show("aeForm found name : " + aeForm.Current.Name);
                        Log("aeForm found name : " + aeForm.Current.Name);
                        logger.LogMessageToFile("aeForm found name : " + aeForm.Current.Name, sLogCount, sLogInterval);
                        Thread.Sleep(5000);

                    }

                    aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 10);

                    #endregion
                }

                if (aeNextButton == null)
                    System.Windows.Forms.MessageBox.Show("next button not found", "After remove, Reinstall App");
                else
                {
                    System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                    Thread.Sleep(2000);
                    Input.MoveTo(OptionPt);
                    Thread.Sleep(2000);
                    logger.LogMessageToFile("<--->Next Button clicking ... ", sLogCount, sLogInterval);
                    Input.ClickAtPoint(OptionPt);
                    Thread.Sleep(2000);

                    /*for (int i = 2; i < 8; i++)
                    {
                        Thread.Sleep(4000);
                        logger.LogMessageToFile(" (" + i + ") " + SetupStepDescriptions[i], sLogCount, sLogInterval);
                        InstallEtriccStatisticsUISetupByStep(SetupWindowName, SetupStepDescriptions[i], i);
                    }*/

                    int i = 2;
                    while (InstallEtriccStatisticsUISetupByStep(SetupWindowName, SetupStepDescriptions[i], i) == false)
                    {
                        i++;
                    }

                    installed = true;
                }
                #endregion

            }
            catch (Exception ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                System.Windows.Forms.MessageBox.Show(msg, "InstallSetup");
                throw new Exception(msg + "  during <InstallSetup>");
                // WIP log exception and return 
            }

            //Log("End Install Setup " + msiName + " at " + FilePath);


            if (installed)
            {
                try
                {
                    //save the name and the path in the registry to remove the setup
                    WriteInstallationToReg(Constants.ETRICCSTATISTICS_UI, FilePath, msiName);

                    // save deployment log  
                    AutoDeploymentOutputLog(Constants.ETRICCSTATISTICS_UI, m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Etricc Statistics Parser");
                    // 
                    string testResultsFile = Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "TestResults.txt");
                    if (!File.Exists(testResultsFile))
                    {
                        logger.LogMessageToFile("create test result file:" + testResultsFile, 0, 0);
                    }
                    else // empty file
                    {
                        logger.LogMessageToFile("empty test result to:" + testResultsFile, 0, 0);
                        StreamWriter writer = File.CreateText(testResultsFile);
                        writer.Close();
                    }

                    StreamWriter writerWorker = File.CreateText(testResultsFile);
                    writerWorker.WriteLine("failed");
                    writerWorker.Close();

                    logger.LogMessageToFile("write testResults: " + Path.Combine(@"C:\EtriccTests", "TestResults.txt"), 0, 0);

                    logger.LogMessageToFile(" **************** ( END " + "Etricc.Statistics.Ui.msi Deployment )************************** ", 0, 0);

                }
                catch (Exception ex)
                {
                    string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                    System.Windows.Forms.MessageBox.Show(msg, "InstallEtriccStatisticsUiSetup");
                    throw new Exception(msg + "  during <InstallEtriccStatisticsUiSetup>");
                    // WIP log exception and return 
                }
            }

            return installed;

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

        public void InstallEpiaSetupByStep(string WindowName, string stepMsg, int step)
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
                System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", stepMsg);
            }
            else
            {
                if (step == 2)
                {
                    AutomationElement aeIAgreeRadioButton
                        = AUIUtilities.FindElementByName("E'pia Server", aeWindow);
                    TogglePattern tg = aeIAgreeRadioButton.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                    ToggleState tgState = tg.Current.ToggleState;
                    if (tgState == ToggleState.Off)
                        tg.Toggle();

                    AutomationElement aeShellckb
                        = AUIUtilities.FindElementByName("E'pia Shell", aeWindow);
                    TogglePattern tg2 = aeShellckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                    ToggleState tg2State = tg2.Current.ToggleState;
                    if (tg2State == ToggleState.Off)
                        tg2.Toggle();

                    logger.LogMessageToFile("<---> Components Sever and shell checkbox state set on  ... ", sLogCount, sLogInterval);
                    Thread.Sleep(3000);
                }


                if (step == 6)   // E'pia Framework 2010.12.22.* Information.
                {
                    logger.LogMessageToFile("<----->wait extra 15 second  : ", sLogCount, sLogInterval);
                    WaitUntilMyButtonFoundInThisWindowWithStatusEnable(WindowName, "Next >", 600);
                    Thread.Sleep(3000);
                }
                else if (step == 5)
                {
                    logger.LogMessageToFile("<----->Do nothing  : ", sLogCount, sLogInterval);
                    Thread.Sleep(3000);
                }
                else
                {
                    string buttonName = (step == 8) ? "Close" : "Next >";
                    AutomationElement aeNextButton = null;
                    StartTime = DateTime.Now;
                    Time = DateTime.Now - StartTime;
                    int wt = 0;
                    while (aeNextButton == null && wt < 500)
                    {
                        aeNextButton = FindSpecificButtonByName(aeWindow, (step == 8) ? "Close" : "Next >");
                        Time = DateTime.Now - StartTime;
                        Thread.Sleep(2000);
                        wt = wt + 2;
                    }

                    if (aeNextButton == null)
                    {
                        System.Windows.Forms.MessageBox.Show(" <-----> Button not found ;" + buttonName, stepMsg);
                    }
                    else
                    {
                        System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(OptionPt);
                        Thread.Sleep(2000);
                        logger.LogMessageToFile("<---> " + buttonName + " button clicking ... ", sLogCount, sLogInterval);
                        Input.ClickAtPoint(OptionPt);
                        Thread.Sleep(2000);
                    }
                }
            }
        }

        public bool InstallEtricc5SetupByStep(ref string errorMsg)
        {
            bool installed = true;
            mInstallEtriccLauncher = m_TestConfigSettings.Element.InstallEtriccLauncher;
            mInstallOldEtricc5Service = m_TestConfigSettings.Element.InstallOldEtriccService;
            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");

            AutomationElement aeWindow = null;
            DateTime StartTime = DateTime.Now;
            TimeSpan Time = DateTime.Now - StartTime;

            while (aeWindow == null && Time.TotalSeconds <= 600)
            {
                aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                //aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c2p);
                Thread.Sleep(2000);
                Time = DateTime.Now - StartTime;
                logger.LogMessageToFile("<-----> find Window: " + Time.Seconds, sLogCount, sLogInterval);
            }

            if (aeWindow == null)
            {
                errorMsg = " <-----> Window not found ";
                installed = false;
            }
            else
            {
                Thread.Sleep(3000);
                Console.WriteLine("Welcom Etricc Core window opend...");
                logger.LogMessageToFile("Welcom Etricc Core window opend...", sLogCount, sLogInterval);
                Console.WriteLine("Searching next button...");
                AutomationElement btnNext = AUIUtilities.GetElementByNameProperty(aeWindow, "Next >");
                if (btnNext != null)
                {
                    Input.MoveToAndClick(btnNext);
                    Thread.Sleep(2000);
                    aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                    Thread.Sleep(2000);
                    Console.WriteLine("License agreement window opend...");
                    if (aeWindow != null)
                        logger.LogMessageToFile("License agreement window opend...", sLogCount, sLogInterval);
                    else
                    {
                        logger.LogMessageToFile("License agreement window not opend...", sLogCount, sLogInterval);
                        return false;
                    }

                    Console.WriteLine("Searching I Agree button...");
                    AutomationElement aeBtnAgree = AUIUtilities.GetElementByNameProperty(aeWindow, "I Agree");
                    if (aeBtnAgree == null)
                    {
                        errorMsg = "Agree button not found...";
                        installed = false;
                        return installed;
                    }

                    Point pt = AUIUtilities.GetElementCenterPoint(aeBtnAgree);
                    Input.MoveTo(pt);
                    Thread.Sleep(2000);
                    AUIUtilities.ClickElement(aeBtnAgree);
                    
                    Thread.Sleep(2000);
                    btnNext = AUIUtilities.GetElementByNameProperty(aeWindow, "Next >");
                    if (btnNext != null)
                    {
                        Input.MoveToAndClick(btnNext);
                        Thread.Sleep(3000);
                    }
                    else
                    {
                        errorMsg = "Next > button not found..." + "in License agreement window";
                        installed = false;
                        return installed;
                    }

                    aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                    Console.WriteLine("Select componts to install window opend...");
                    Console.WriteLine("Searching Next button...");

                    AutomationElement aeLauncherkb
                        = AUIUtilities.FindElementByName("Launcher", aeWindow);
                    TogglePattern tgaeLauncherkb = aeLauncherkb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                    ToggleState tgLauncherState = tgaeLauncherkb.Current.ToggleState;
                    logger.LogMessageToFile("<----->mInstallEtriccLauncher  : " + mInstallEtriccLauncher, sLogCount, sLogInterval);
                    
                    if (mInstallEtriccLauncher)
                    {
                        if (tgLauncherState == ToggleState.Off)
                        {
                            Thread.Sleep(2000);
                            tgaeLauncherkb.Toggle();
                        }
                    }
                    else
                    {
                        if (tgLauncherState == ToggleState.On)
                        {
                            logger.LogMessageToFile("<----->tgLauncherState == ToggleState.On  : ", sLogCount, sLogInterval);
                            Thread.Sleep(2000);
                            //tgaeLauncherkb.Toggle();
                            Input.MoveToAndClick(aeLauncherkb);
                        }
                    }
                    
                    AutomationElement aeServicekb
                        = AUIUtilities.FindElementByName("Service", aeWindow);
                    TogglePattern tgaeServicekb = aeServicekb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                    ToggleState tgServiceState = tgaeServicekb.Current.ToggleState;
                    logger.LogMessageToFile("<----->mInstallOldEtricc5Service  : " + mInstallOldEtricc5Service, sLogCount, sLogInterval);
                    if (mInstallOldEtricc5Service)
                    {
                        if (tgServiceState == ToggleState.Off)
                        {
                            Thread.Sleep(2000);
                            tgaeServicekb.Toggle();
                        }
                    }
                    else
                    {
                        if (tgServiceState == ToggleState.On)
                        {
                            logger.LogMessageToFile("<----->tgServiceState == ToggleState.On  : ", sLogCount, sLogInterval);
                            Thread.Sleep(2000);
                            //tgaeServicekb.Toggle();
                            Input.MoveToAndClick(aeServicekb);
                            Thread.Sleep(2000);
                        }
                    }

                    btnNext = AUIUtilities.GetElementByNameProperty(aeWindow, "Next >");
                    if (btnNext != null)
                    {
                        Input.MoveToAndClick(btnNext);
                        Thread.Sleep(3000);
                        aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                        Console.WriteLine("Environment Configuration window opend...");
                        Console.WriteLine("Searching Next button...");
                        btnNext = AUIUtilities.GetElementByNameProperty(aeWindow, "Next >");
                        if (btnNext != null)
                        {
                            Input.MoveToAndClick(btnNext);
                            Thread.Sleep(3000);
                            aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                            Console.WriteLine("Select Installation Folder window opend...");
                            Console.WriteLine("Searching Next button...");
                            btnNext = AUIUtilities.GetElementByNameProperty(aeWindow, "Next >");
                            if (btnNext != null)
                            {
                                Input.MoveToAndClick(btnNext);
                                Thread.Sleep(3000);
                                aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                Console.WriteLine("Confirm Installation window opend...");
                                Console.WriteLine("Searching Next button...");
                                btnNext = AUIUtilities.GetElementByNameProperty(aeWindow, "Next >");
                                if (btnNext != null)
                                {
                                    AUIUtilities.ClickElement(btnNext);
                                    Thread.Sleep(3000);
                                    aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                    Console.WriteLine("Installing Etricc  window opend...");
                                    Console.WriteLine("Installing Etricc  window move to left...");

                                    AutomationElement aeTitleBar =
                                        AUIUtilities.FindElementByID("TitleBar", aeWindow);

                                    Point pt1 = new Point(
                                        (aeTitleBar.Current.BoundingRectangle.Left + aeTitleBar.Current.BoundingRectangle.Right) / 2,
                                        (aeTitleBar.Current.BoundingRectangle.Bottom + aeTitleBar.Current.BoundingRectangle.Top) / 2);

                                    Point newPt1 = new Point(200, 100);
                                    Input.MoveTo(pt1);

                                    Thread.Sleep(1000);
                                    Input.SendMouseInput(pt1.X, pt1.Y, 0, SendMouseInputFlags.LeftDown | SendMouseInputFlags.Absolute);
                                    Thread.Sleep(1000);
                                    Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.Move | SendMouseInputFlags.Absolute);
                                    Thread.Sleep(1000);
                                    Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);
                                    Thread.Sleep(3000);

                                    Console.WriteLine("try to find  FrmLauncherFunctionality window...");

                                    System.Windows.Automation.Condition conditionFrm = new PropertyCondition(AutomationElement.AutomationIdProperty, "FrmLauncherFunctionality");
                                    System.Windows.Automation.Condition conditionFrm2 = new PropertyCondition(AutomationElement.AutomationIdProperty, "FrmSecurityFunctionality");

                                    AutomationElement frmElement = null;
                                    if (mInstallEtriccLauncher)
                                    {
                                        while (frmElement == null)
                                        {
                                            Thread.Sleep(5000);
                                            Console.WriteLine("Wait until Next button found...");
                                            frmElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, conditionFrm);
                                            if (frmElement == null)
                                                Console.WriteLine("Frm Window  not found");
                                            else
                                            {
                                                System.Windows.Automation.Condition cNt = new AndCondition(
                                                new PropertyCondition(AutomationElement.NameProperty, "Next >"),
                                                new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                                                );

                                                AutomationElement aeBtnNext = frmElement.FindFirst(TreeScope.Element | TreeScope.Descendants, cNt);
                                                Console.WriteLine("Next button found: " + aeBtnNext.Current.Name);
                                                Console.WriteLine("Nex button found... ---> Click NExt button");
                                                Input.MoveToAndClick(aeBtnNext);
                                            }
                                        }
                                    }

                                    AutomationElement frmElement2 = null;
                                    if (mInstallOldEtricc5Service)
                                    {
                                        while (frmElement2 == null)
                                        {
                                            Thread.Sleep(5000);
                                            Console.WriteLine("Wait until Next button found...");
                                            frmElement2 = AutomationElement.RootElement.FindFirst(TreeScope.Children, conditionFrm2);
                                            if (frmElement2 == null)
                                                Console.WriteLine("Frm2 Window  not found");
                                            else
                                            {
                                                System.Windows.Automation.Condition cNt = new AndCondition(
                                                new PropertyCondition(AutomationElement.NameProperty, "Next >"),
                                                new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                                                );

                                                AutomationElement aeBtnNext = frmElement2.FindFirst(TreeScope.Element | TreeScope.Descendants, cNt);
                                                Console.WriteLine("Next button found: " + aeBtnNext.Current.Name);
                                                Console.WriteLine("Nex button found... ---> Click NExt button");
                                                Input.MoveToAndClick(aeBtnNext);
                                            }
                                        }
                                    }
                                    Console.WriteLine("Installation complete  window opend...");
                                    Console.WriteLine("Searching close button...");

                                    // Wait until Close button Found
                                    aeWindow = null;
                                    while (aeWindow == null)
                                    {

                                        try
                                        {
                                            aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);

                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine(ex.Message + ".." + ex.StackTrace);
                                            aeWindow = null;
                                            Thread.Sleep(5000);
                                        }
                                    }

                                    System.Windows.Automation.Condition c2 = new AndCondition(
                                            new PropertyCondition(AutomationElement.NameProperty, "Close"),
                                            new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                                        );
                                    AutomationElement aeBtnClose
                                        = aeWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);

                                    while (aeBtnClose == null)
                                    {
                                        Thread.Sleep(5000);
                                        Console.WriteLine("Wait until Close button found...");
                                        aeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                        if (aeWindow == null)
                                            Console.WriteLine("Installer Window  not found");
                                        else
                                        {
                                            aeBtnClose = aeWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                                            Console.WriteLine("Installer Window found: " + aeWindow.Current.Name);
                                        }
                                    }
                                    Console.WriteLine("Close button found... ---> Close Installer Window");
                                    Input.MoveToAndClick(aeBtnClose);
                                }
                            }
                        }
                    }
                }
                else
                {
                    errorMsg = "Next > button not found..." + "in Select Components window";
                    installed = false;
                }
            }

            return installed;
        }

        private void WriteInstallationToReg(string AppName, string FilePath, string MsiName)
        {
            logger.LogMessageToFile("------------WriteInstallationToReg: " + AppName + "InstallationPath"
                + System.Environment.NewLine + "FilePath: " + FilePath + " & " + "MsiName: " + MsiName, 0, 0);
            RegistryKey key = Registry.CurrentUser.CreateSubKey(REGKEY);
            key.SetValue(AppName + "InstallationPath", FilePath);
            key.SetValue(AppName + "InstallationName", MsiName);
        }

        private void AutoDeploymentOutputLog(string InstalledApp, string appDeployedPath)
        {
            // After Epia installed, write the log record to output file: Format
            //2007-4-27 17:47:36: , Installed, C:\Program Files\Egemin\Epia 1.9.10, msi path info, rootpath                                                                            // 0 Time 
            string msg = ", Installed:" + InstalledApp                              // 1 Installed App
                + " , " + appDeployedPath                                           // 2 Deployed Path
                + " , " + mTestDef                                                  // 3 Definition: CI, Nightly... 
                + " , " + mTestApp                                                  // 4 APP
                + " , " + TESTTOOL_VERSION                                          // 5 TestTool Version
                + " , " + m_TestConfigSettings.Element.SelectedProjectFile.Substring(0,
                            m_TestConfigSettings.Element.SelectedProjectFile.LastIndexOf("."))       // 6 Default Project File
                + " , " + m_installScriptDir                                        // 7 Build install path
                + " , " + m_SystemDrive                                             // 8 Build base path
                + " , " + sDemonstration.ToString().ToLower()                       // 9 demo          
                + " , " + m_TestConfigSettings.Element.Mail.ToString().ToLower()    // 10 send mail
                + " , " + m_TestConfigSettings.Element.Excel.ToString()             // 11 Excel Visible
                + " , " + TestTools.HelpUtilities.GetPCOS()                         // 12 OS
                + " , " + ConstCommon.ETRICC_TESTS_DIRECTORY                        // 13 
                + " , " + m_TestWorkingDirectory                                    // 14 Working directory
                + " , " + m_TestAutoMode;                                           // 15 AutoMode?

            string path = Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, "TestAutoDeploymentOutputLog.txt");
            if (!File.Exists(path))
            {
                FileStream fs = File.Create(path);
                fs.Close();
                logger.LogMessageToFile("create install info to:" + path, 0, 0);
            }
            else // empty file
            {
                logger.LogMessageToFile("empty install info to:" + path, 0, 0);
                StreamWriter writer = File.CreateText(path);
                writer.Close();
            }

            logger.LogMessageToFile("write install info to:" + path, 0, 0);
            System.IO.StreamWriter sw = System.IO.File.CreateText(path);
            //System.IO.StreamWriter sw = System.IO.File.AppendText(path);
            //Path.Combine( logFilePath, logFileName ));
            try
            {
                string logLine = System.String.Format(
                    "{0:G}: {1}.", System.DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }

        public void InstallEtriccUISetupByStep(string WindowName, string stepMsg, int step)
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
                System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", stepMsg);
            }
            else
            {
                if (step == 2)
                {
                    AutomationElement aeIAgreeRadioButton
                        = AUIUtilities.FindElementByName("E'pia Server Extensions (Resource & Security Files)", aeWindow);
                    TogglePattern tg = aeIAgreeRadioButton.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                    ToggleState tgState = tg.Current.ToggleState;
                    if (tgState == ToggleState.Off)
                        tg.Toggle();

                    AutomationElement aeShellckb
                        = AUIUtilities.FindElementByName("E'pia Shell Extensions (Shell Module & Config)", aeWindow);
                    TogglePattern tg2 = aeShellckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                    ToggleState tg2State = tg2.Current.ToggleState;
                    if (tg2State == ToggleState.Off)
                        tg2.Toggle();

                    AutomationElement aeEtriccckb
                        = AUIUtilities.FindElementByName("E'tricc Core Extensions (Wrappers)", aeWindow);
                    TogglePattern tg3 = aeEtriccckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                    ToggleState tg3State = tg3.Current.ToggleState;
                    if (tg3State == ToggleState.Off)
                        tg3.Toggle();

                    logger.LogMessageToFile("<---> Components Sever and shell checkbox state set on  ... ", sLogCount, sLogInterval);
                    Thread.Sleep(3000);
                }


                if (step == 6)
                {
                    logger.LogMessageToFile("<----->wait extra 15 second  : ", sLogCount, sLogInterval);
                    WaitUntilMyButtonFoundInThisWindow(WindowName, "Close", 601);
                    Thread.Sleep(3000);
                }
                else if (step == 5)
                {
                    logger.LogMessageToFile("<----->Do nothing  : ", sLogCount, sLogInterval);
                    Thread.Sleep(3000);
                }
                else
                {
                    string buttonName = (step == 7) ? "Close" : "Next >";
                    AutomationElement aeNextButton = null;
                    StartTime = DateTime.Now;
                    Time = DateTime.Now - StartTime;
                    int wt = 0;
                    while (aeNextButton == null && wt < 500)
                    {
                        aeNextButton = FindSpecificButtonByName(aeWindow, (step == 7) ? "Close" : "Next >");
                        Time = DateTime.Now - StartTime;
                        Thread.Sleep(2000);
                        wt = wt + 2;
                    }

                    if (aeNextButton == null)
                    {
                        System.Windows.Forms.MessageBox.Show(" <-----> Button not found ;" + buttonName, stepMsg);
                    }
                    else
                    {
                        System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(OptionPt);
                        Thread.Sleep(2000);
                        logger.LogMessageToFile("<---> " + buttonName + " button clicking ... ", sLogCount, sLogInterval);
                        Input.ClickAtPoint(OptionPt);
                        Thread.Sleep(2000);
                    }
                }
            }
        }

        public bool InstallEtriccStatisticsParserSetupByStep(string WindowName, string stepMsg, int step)
        {
            bool clickCloseButton = false;
            AutomationElement aeWindow = null;
            AutomationElementCollection aeAllWindows = null;
            AutomationElement aeClickButton = null;
            //find all install Window screen texts
            System.Windows.Automation.Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);
           

            DateTime StartTime = DateTime.Now;
            TimeSpan Time = DateTime.Now - StartTime;
            try
            {
                while (aeWindow == null && Time.TotalSeconds <= 600)
                {
                    aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                    Thread.Sleep(3000);
                    logger.LogMessageToFile("<----->aeWindows count  : " + aeAllWindows.Count, sLogCount, sLogInterval);
                    
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        logger.LogMessageToFile("<----->Window title(" + i + ")  : " + aeAllWindows[i].Current.Name, sLogCount, sLogInterval);
                        if (aeAllWindows[i].Current.Name.StartsWith("E'tricc Statistics Parser"))
                        {
                            aeWindow = aeAllWindows[i];
                            logger.LogMessageToFile("<----->Window title(" + i + ")  : " + aeWindow.Current.Name, sLogCount, sLogInterval);
                            break;
                        }
                    }
                    Time = DateTime.Now - StartTime;
                }

                logger.LogMessageToFile("<-----> install step is : " + step, sLogCount, sLogInterval);

           
                if (aeWindow == null)
                {
                    System.Windows.Forms.MessageBox.Show(" <-----> after 10 minute Window not found ", stepMsg);
                }
                else
                {
                    logger.LogMessageToFile("<-----> find aeWindow name is:" + aeWindow.Current.Name, sLogCount, sLogInterval);
                    //find all install Window screen texts
                    System.Windows.Automation.Condition cText = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Text);

                    AutomationElementCollection aeAllTexts = aeWindow.FindAll(TreeScope.Children, cText);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllTexts.Count; i++)
                    {
                        logger.LogMessageToFile("<----->Window text(" + i + ")  : " + aeAllTexts[i].Current.Name, sLogCount, sLogInterval);
                    }

                    System.Windows.Automation.Condition cButton = new AndCondition(
                        new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                        new PropertyCondition(AutomationElement.IsContentElementProperty, true),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                    );

                    AutomationElementCollection aeButtons = aeWindow.FindAll(TreeScope.Children, cButton);
                    Thread.Sleep(3000);
                    bool clickButtonFound = false;
                    for (int i = 0; i < aeButtons.Count; i++)
                    {
                        logger.LogMessageToFile("<----->Window enabled button(" + i + ")  : " + aeButtons[i].Current.Name, sLogCount, sLogInterval);
                        if (aeButtons[i].Current.Name.StartsWith("Next"))
                        {
                            aeClickButton = aeButtons[i];
                            clickButtonFound = true;
                            break;
                        }
                        else if (aeButtons[i].Current.Name.StartsWith("Close"))
                        {
                            aeClickButton = aeButtons[i];
                            clickButtonFound = true;
                            clickCloseButton = true;
                            break;
                        }
                    }

                    if (clickButtonFound)
                    {
                        System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(OptionPt);
                        Thread.Sleep(1000);
                        logger.LogMessageToFile("<---> " + aeClickButton.Current.Name + " button clicking ... ", sLogCount, sLogInterval);
                        Input.ClickAtPoint(OptionPt);
                        Thread.Sleep(5000);
                    }
                }
            }
            catch (System.Windows.Automation.ElementNotAvailableException ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                if (sMsgDebug.StartsWith("true"))
                    System.Windows.Forms.MessageBox.Show(msg, "InstallEtriccStatisticsParserSetup");
                logger.LogMessageToFile("<---> " + msg, sLogCount, sLogInterval);
                clickCloseButton = false;
            }

            return clickCloseButton;
        }

        public bool InstallEtriccStatisticsParserConfiguratorSetupByStep(string WindowName, string stepMsg, int step)
        {
            bool clickCloseButton = false;
            AutomationElement aeWindow = null;
            AutomationElementCollection aeAllWindows = null;
            AutomationElement aeClickButton = null;
            //find all install Window screen texts
            System.Windows.Automation.Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            DateTime StartTime = DateTime.Now;
            TimeSpan Time = DateTime.Now - StartTime;
            try
            {
                while (aeWindow == null && Time.TotalMinutes <= 2)
                {
                    aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        if (aeAllWindows[i].Current.Name.StartsWith("E'tricc Statistics ParserConfigurator"))
                        {
                            aeWindow = aeAllWindows[i];
                            logger.LogMessageToFile("<----->Window title(" + i + ")  : " + aeWindow.Current.Name, sLogCount, sLogInterval);
                            break;
                        }
                    }
                    Time = DateTime.Now - StartTime;
                }

                logger.LogMessageToFile("<-----> install step is : " + step, sLogCount, sLogInterval);
           
                if (aeWindow == null)
                {
                    System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", stepMsg);
                }
                else
                {
                    logger.LogMessageToFile("<-----> find aeWindow name is:" + aeWindow.Current.Name, sLogCount, sLogInterval);
                    //find all install Window screen texts
                    System.Windows.Automation.Condition cText = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Text);

                    AutomationElementCollection aeAllTexts = aeWindow.FindAll(TreeScope.Children, cText);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllTexts.Count; i++)
                    {
                        logger.LogMessageToFile("<----->Window text(" + i + ")  : " + aeAllTexts[i].Current.Name, sLogCount, sLogInterval);
                    }

                    System.Windows.Automation.Condition cButton = new AndCondition(
                        new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                        new PropertyCondition(AutomationElement.IsContentElementProperty, true),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                    );

                    AutomationElementCollection aeButtons = aeWindow.FindAll(TreeScope.Children, cButton);
                    Thread.Sleep(3000);
                    bool clickButtonFound = false;
                    for (int i = 0; i < aeButtons.Count; i++)
                    {
                        logger.LogMessageToFile("<----->Window enabled button(" + i + ")  : " + aeButtons[i].Current.Name, sLogCount, sLogInterval);
                        if (aeButtons[i].Current.Name.StartsWith("Next"))
                        {
                            aeClickButton = aeButtons[i];
                            clickButtonFound = true;
                            break;
                        }
                        else if (aeButtons[i].Current.Name.StartsWith("Close"))
                        {
                            aeClickButton = aeButtons[i];
                            clickButtonFound = true;
                            clickCloseButton = true;
                            break;
                        }
                    }

                    if (clickButtonFound)
                    {
                        System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(OptionPt);
                        Thread.Sleep(1000);
                        logger.LogMessageToFile("<---> " + aeClickButton.Current.Name + " button clicking ... ", sLogCount, sLogInterval);
                        Input.ClickAtPoint(OptionPt);
                        Thread.Sleep(5000);
                    }
                    else
                        Thread.Sleep(5000);
                }
            }
            catch (System.Windows.Automation.ElementNotAvailableException ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                if (sMsgDebug.StartsWith("true"))
                    System.Windows.Forms.MessageBox.Show(msg, "InstallEtriccStatisticsParserConfiguratorSetup");
                logger.LogMessageToFile("<---> " + msg, sLogCount, sLogInterval);
                clickCloseButton = false;
            }

            return clickCloseButton;
        }

        public bool InstallEtriccStatisticsUISetupByStep(string WindowName, string stepMsg, int step)
        {
            logger.LogMessageToFile("<-----> Start InstallEtriccStatisticsUISetupByStep : " + step, sLogCount, sLogInterval);
            bool clickCloseButton = false;
            AutomationElement aeWindow = null;

            AutomationElementCollection aeAllWindows = null;
            AutomationElement aeClickButton = null;
            //find all install Window screen texts
            System.Windows.Automation.Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            DateTime StartTime = DateTime.Now;
            TimeSpan Time = DateTime.Now - StartTime;
            logger.LogMessageToFile("<-----> Start1 InstallEtriccStatisticsUISetupByStep : " + step, sLogCount, sLogInterval);
            try
            {
                while (aeWindow == null && Time.TotalMinutes <= 2)
                {
                    aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                    logger.LogMessageToFile("<-----> Start2 InstallEtriccStatisticsUISetupByStep  aeAllWindows.Count: " + aeAllWindows.Count, sLogCount, sLogInterval);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        logger.LogMessageToFile("<----->Window title(" + i + ")  : " + aeAllWindows[i].Current.Name, sLogCount, sLogInterval);
                        if (aeAllWindows[i].Current.Name.StartsWith("E'tricc Statistics UI"))
                        {
                            aeWindow = aeAllWindows[i];
                            logger.LogMessageToFile("<----->Window title(" + i + ")  : " + aeWindow.Current.Name, sLogCount, sLogInterval);
                            break;
                        }
                    }
                    Time = DateTime.Now - StartTime;
                }

                logger.LogMessageToFile("<-----> install step is : " + step, sLogCount, sLogInterval);

                if (aeWindow == null)
                {
                    System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", stepMsg);
                }
                else
                {
                    logger.LogMessageToFile("<-----> find aeWindow name is:" + aeWindow.Current.Name, sLogCount, sLogInterval);
                    //find all install Window screen texts
                    System.Windows.Automation.Condition cText = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Text);

                    AutomationElementCollection aeAllTexts = aeWindow.FindAll(TreeScope.Children, cText);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllTexts.Count; i++)
                    {
                        logger.LogMessageToFile("<----->Window text(" + i + ")  : " + aeAllTexts[i].Current.Name, sLogCount, sLogInterval);
                        if (aeAllTexts[i].Current.Name.StartsWith("Components"))
                        {
                            AutomationElement aeIAgreeRadioButton
                                = AUIUtilities.FindElementByName("E'pia Server Extensions (Resource & Security Files && Server Components)", aeWindow);
                            TogglePattern tg = aeIAgreeRadioButton.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                            ToggleState tgState = tg.Current.ToggleState;
                            if (tgState == ToggleState.Off)
                                tg.Toggle();

                            AutomationElement aeShellckb
                                = AUIUtilities.FindElementByName("E'pia Shell Extensions (Shell Module & Config)", aeWindow);
                            TogglePattern tg2 = aeShellckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                            ToggleState tg2State = tg2.Current.ToggleState;
                            if (tg2State == ToggleState.Off)
                                tg2.Toggle();

                            logger.LogMessageToFile("<---> Components Sever and shell checkbox state set on  ... ", sLogCount, sLogInterval);
                            Thread.Sleep(3000);
                        }
                    }

                    System.Windows.Automation.Condition cButton = new AndCondition(
                        new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                        new PropertyCondition(AutomationElement.IsContentElementProperty, true),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                    );

                    AutomationElementCollection aeButtons = aeWindow.FindAll(TreeScope.Children, cButton);
                    Thread.Sleep(3000);
                    bool clickButtonFound = false;
                    for (int i = 0; i < aeButtons.Count; i++)
                    {
                        logger.LogMessageToFile("<----->Window enabled button(" + i + ")  : " + aeButtons[i].Current.Name, sLogCount, sLogInterval);
                        if (aeButtons[i].Current.Name.StartsWith("Next"))
                        {
                            aeClickButton = aeButtons[i];
                            clickButtonFound = true;
                            break;
                        }
                        else if (aeButtons[i].Current.Name.StartsWith("Close"))
                        {
                            aeClickButton = aeButtons[i];
                            clickButtonFound = true;
                            clickCloseButton = true;
                            break;
                        }
                    }

                    if (clickButtonFound)
                    {
                        System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(OptionPt);
                        Thread.Sleep(1000);
                        logger.LogMessageToFile("<---> " + aeClickButton.Current.Name + " button clicking ... ", sLogCount, sLogInterval);
                        Input.ClickAtPoint(OptionPt);
                        Thread.Sleep(5000);
                    }
                    else
                        Thread.Sleep(5000);
                }
            }
            catch (System.Windows.Automation.ElementNotAvailableException ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                if (sMsgDebug.StartsWith("true"))
                    System.Windows.Forms.MessageBox.Show(msg, "InstallEtriccStatisticsUISetup");
                logger.LogMessageToFile("<---> " + msg, sLogCount, sLogInterval);
                clickCloseButton = false;
            }

            return clickCloseButton;
        }

        public bool InstallEpiaSetupByStep2(string WindowName, string stepMsg, int step)
        {
            logger.LogMessageToFile("<-----> Start InstallEpiaSetupByStep2 : " + step, sLogCount, sLogInterval);
            bool clickCloseButton = false;
            AutomationElement aeWindow = null;

            AutomationElementCollection aeAllWindows = null;
            AutomationElement aeClickButton = null;
            //find all install Window screen texts
            System.Windows.Automation.Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            DateTime StartTime = DateTime.Now;
            TimeSpan Time = DateTime.Now - StartTime;
            logger.LogMessageToFile("<-----> Start1 InstallEpiaSetupByStep2 : " + step, sLogCount, sLogInterval);
            try
            {
                while (aeWindow == null && Time.TotalMinutes <= 2)
                {
                    aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                    logger.LogMessageToFile("<-----> Start2 InstallEpiaSetupByStep2  aeAllWindows.Count: " + aeAllWindows.Count, sLogCount, sLogInterval);
                    Thread.Sleep(3000);
                    try
                    {
                        for (int i = 0; i < aeAllWindows.Count; i++)
                        {
                            logger.LogMessageToFile("<----->Window title(" + i + ")  : " + aeAllWindows[i].Current.Name, sLogCount, sLogInterval);
                            if (aeAllWindows[i].Current.Name.StartsWith("E'pia Framework"))
                            {
                                aeWindow = aeAllWindows[i];
                                logger.LogMessageToFile("<----->Window title(" + i + ")  : " + aeWindow.Current.Name, sLogCount, sLogInterval);

                                // Make sure our window is usable.
                                // WaitForInputIdle will return before the specified time 
                                // if the window is ready.
                                //WindowPattern windowPattern = aeWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                                //if (false == windowPattern.WaitForInputIdle(30000))
                                //{
                                //    System.Windows.Forms.MessageBox.Show("Object not responding in a timely manner, click OK continue", stepMsg);
                                //}

                                break;
                            }
                        }
                     }
                    catch (System.Windows.Automation.ElementNotAvailableException ex)
                    {
                        string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                        if (sMsgDebug.StartsWith("true"))
                            System.Windows.Forms.MessageBox.Show(msg, "InstallEpiaSetupByStep2");
                        logger.LogMessageToFile("<---> " + msg, sLogCount, sLogInterval);
                        aeWindow = null;
                    }
                    Time = DateTime.Now - StartTime;
                }

                logger.LogMessageToFile("<-----> install step is : " + step, sLogCount, sLogInterval);

                if (aeWindow == null)
                {
                    System.Windows.Forms.MessageBox.Show(" <-----> Window not found ", stepMsg);
                }
                else
                {
                    WindowPattern wp = aeWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
					if ( wp.Current.WindowVisualState == WindowVisualState.Minimized )
                    {
                       
                        //System.Windows.Forms.MessageBox.Show("wp.Current.WindowVisualState == WindowVisualState.Minimized", stepMsg);
                        wp.SetWindowVisualState(WindowVisualState.Normal);
                        Thread.Sleep(1000);
                    }
						


                    logger.LogMessageToFile("<-----> find aeWindow name is:" + aeWindow.Current.Name, sLogCount, sLogInterval);
                    //find all install Window screen texts
                    System.Windows.Automation.Condition cText = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Text);

                    AutomationElementCollection aeAllTexts = aeWindow.FindAll(TreeScope.Children, cText);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllTexts.Count; i++)
                    {
                        logger.LogMessageToFile("<----->Window text(" + i + ")  : " + aeAllTexts[i].Current.Name, sLogCount, sLogInterval);
                        if (aeAllTexts[i].Current.Name.StartsWith("Components"))
                        {
                            AutomationElement aeIAgreeRadioButton
                                = AUIUtilities.FindElementByName("E'pia Server", aeWindow);
                            TogglePattern tg = aeIAgreeRadioButton.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                            ToggleState tgState = tg.Current.ToggleState;
                            if (tgState == ToggleState.Off)
                                tg.Toggle();

                            AutomationElement aeShellckb
                                = AUIUtilities.FindElementByName("E'pia Shell", aeWindow);
                            TogglePattern tg2 = aeShellckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                            ToggleState tg2State = tg2.Current.ToggleState;
                            if (tg2State == ToggleState.Off)
                                tg2.Toggle();

                            logger.LogMessageToFile("<---> Components Sever and shell checkbox state set on  ... ", sLogCount, sLogInterval);
                            Thread.Sleep(3000);
                        }
                    }

                    System.Windows.Automation.Condition cButton = new AndCondition(
                        new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                        new PropertyCondition(AutomationElement.IsContentElementProperty, true),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                    );

                    AutomationElementCollection aeButtons = aeWindow.FindAll(TreeScope.Children, cButton);
                    Thread.Sleep(3000);
                    bool clickButtonFound = false;
                    for (int i = 0; i < aeButtons.Count; i++)
                    {
                        logger.LogMessageToFile("<----->Window enabled button(" + i + ")  : " + aeButtons[i].Current.Name, sLogCount, sLogInterval);
                        if (aeButtons[i].Current.Name.StartsWith("Next"))
                        {
                            aeClickButton = aeButtons[i];
                            clickButtonFound = true;
                            break;
                        }
                        else if (aeButtons[i].Current.Name.StartsWith("Close"))
                        {
                            aeClickButton = aeButtons[i];
                            clickButtonFound = true;
                            clickCloseButton = true;
                            break;
                        }
                    }

                    if (clickButtonFound)
                    {
                        logger.LogMessageToFile("<---> clickButtonFound... " + aeClickButton.Current.Name, sLogCount, sLogInterval);
                        System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                        Thread.Sleep(1000);
                        Input.MoveTo(OptionPt);
                        Thread.Sleep(1000);
                        logger.LogMessageToFile("<---> " + aeClickButton.Current.Name + " button clicking ... ", sLogCount, sLogInterval);
                        Input.ClickAtPoint(OptionPt);
                        Thread.Sleep(5000);
                    }
                    else
                        Thread.Sleep(5000);
                }
            }
            catch (System.Windows.Automation.ElementNotAvailableException ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                if (sMsgDebug.StartsWith("true"))
                    System.Windows.Forms.MessageBox.Show(msg, "InstallEpiaSetupByStep2");
                logger.LogMessageToFile("<---> " + msg, sLogCount, sLogInterval);
                clickCloseButton = false;
            }

            return clickCloseButton;

           
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
            try{
                procExplorer.Start();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("explor start exception:"+ex.Message +"-"+ex.StackTrace);
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
            System.Windows.Forms.MessageBox.Show("Destination=" + Destination, "Create Map result:"+result);
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

        #region OnUIAShellEvent
        public static void OnRemoveSetuoScreenEvent(object src, AutomationEventArgs args)
        {
            sEventEnd = false;
            AutomationElement aeRemoveForm = null;
            AutomationElement aeYesButton = null;
            string removeFormID = "FrmRemoveRegistryKeysDialog";
            DateTime mAppTime = DateTime.Now;
            AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeRemoveForm, removeFormID, mAppTime, 30);

            if (aeRemoveForm == null)
            {
                //MessageBox.Show("aeRemoveForm not found", "OnRemoveSetuoScreenEvent");
            }
            else
            {
                aeYesButton = null;
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.Button);

                AutomationElementCollection aeAllButtons = aeRemoveForm.FindAll(TreeScope.Element | TreeScope.Descendants, c);
                Thread.Sleep(1000);

                // get Yes button
                foreach (AutomationElement s in aeAllButtons)
                {
                    if (s.Current.Name.Equals("Yes") && s.Current.IsContentElement)
                    {
                        aeYesButton = s;
                    }
                }

                if (aeYesButton == null)
                    System.Windows.MessageBox.Show("Yes button not found ");
                else
                {
                    System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeYesButton);
                    Thread.Sleep(1000);
                    Input.MoveTo(OptionPt);
                    Thread.Sleep(2000);
                    Input.ClickAtPoint(OptionPt);
                    Thread.Sleep(3000);
                    sEventEnd = true;
                }
            }
        }

        public static void OnRemoveRepairSetupScreenEvent(object src, AutomationEventArgs args)
        {
            Thread.Sleep(12000);
            sEventEnd = false;
            AutomationElement aeRemoveForm = null;
            AutomationElement aeYesButton = null;
            string removeFormID = "FrmRemoveRegistryKeysDialog";
            DateTime mAppTime = DateTime.Now;

            AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeRemoveForm, removeFormID, mAppTime, 30);

            if (aeRemoveForm == null)
            {
                Thread.Sleep(1000);
                //MessageBox.Show("aeRemoveForm not found", "OnRemoveRepairSetupScreenEvent");
            }
            else
            {
                Thread.Sleep(1000);
                //MessageBox.Show("aeRemoveForm found", "OnRemoveRepairSetupScreenEvent");
                aeYesButton = null;
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.Button);

                AutomationElementCollection aeAllButtons = aeRemoveForm.FindAll(TreeScope.Element | TreeScope.Descendants, c);
                Thread.Sleep(1000);

                // get Yes button
                foreach (AutomationElement s in aeAllButtons)
                {
                    if (s.Current.Name.Equals("Yes") && s.Current.IsContentElement)
                    {
                        aeYesButton = s;
                    }
                }

                if (aeYesButton == null)
                    System.Windows.MessageBox.Show("Yes button not found ");
                else
                {
                    System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeYesButton);
                    Thread.Sleep(1000);
                    Input.MoveTo(OptionPt);
                    Thread.Sleep(2000);
                    Input.ClickAtPoint(OptionPt);
                    Thread.Sleep(3000);
                    sEventEnd = true;
                }
            }
        }

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
                    System.Windows.Forms.MessageBox.Show("find button in " + ButtonName + " <-----> Window not found: windows name is: " + WindowName, "WaitUntilMyButtonFoundInThisWindow:"+searchCnt);
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
            Thread.Sleep(1000);

            if (aeAllButtons.Count == 1)
            {
                if (aeAllButtons[0].Current.Name.Equals(buttonName))
                {
                    status = true;
                }

            }

            return status;
        }


        #region OnUIAShellEvent
        public static void OnUIAShellEvent(object src, AutomationEventArgs args)
        {
            string sErrorMessage = string.Empty;
            Console.WriteLine("OnUIAShellEvent");
            logger.LogMessageToFile("OnUIAShellEvent  ====", 0, 0);
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
            logger.LogMessageToFile("OnUIAShellEvent  ==== " + str, 0, 0);


            Thread.Sleep(2000);
            if (name.Equals("Error"))
            {
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                //logger.LogMessageToFile(" start deployment error window <-----> ", sLogCount, sLogInterval);
                Thread.Sleep(6000);
            }
            else if (name.Equals("Open File - Security Warning"))
            {
                //Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
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
                //Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
            }
            sEventEnd = true;
        }
        #endregion
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

        public enum STATE
        {
            UNDEFINED,
            PENDING,
            INPROGRESS,
            EXCEPTION,
        }
    }
}
