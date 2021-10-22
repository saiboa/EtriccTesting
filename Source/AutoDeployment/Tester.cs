using System;
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


namespace Epia3Deployment
{
	public class Tester 
	{
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Fields of Tester (17)
		// --- LOG
		internal static TestTools.Logger logger = null;
		internal static string sLogFilename = string.Empty;

		private static string m_BuildNumber = string.Empty;
		
		static string sEpiaInstallationFolder = string.Empty;
		static string sEtriccInstallationFolder = string.Empty;
		static string sEtricc5InstallationFolder = string.Empty;

        static string sEpia4InstallerName = "Epia.msi";
		private string mTestedVersion = string.Empty;
		private string mEpiaPath = string.Empty;

		internal string TESTTOOL_VERSION = "1.10.3.1";
		internal string m_TestWorkingDirectory = string.Empty;
		string m_installScriptDir = string.Empty;
		internal StringCollection m_Logging = new StringCollection();
		static bool sEventEnd = false;
 
		// --- TEST PARAMS
		static string m_MapNetworkDrive_Root = ConstCommon.DRIVE_MAP_LETTER;
		static string m_CurrentDrive = string.Empty;
		static string m_SystemDrive = string.Empty;
		private Settings m_Settings;
		internal STATE m_State;
		internal bool m_TestAutoMode = true;
		internal string m_TestPC = string.Empty;

		// tested build info
		string m_ValidatedBuildDirectory = string.Empty;
		private Uri m_Uri = null;
		string m_testApp = string.Empty;     //  Epia, Etricc UI, Etricc 5, Kimberly Clark
        string m_testBranch = string.Empty;     //  Main, Dev01, Dev02, Dev03
		string m_testDef = string.Empty;     //  CI, Nightly, Weekly, Version

        string sEpiaRelativePath = string.Empty;
        string sEtriccUIRelativePath = string.Empty;
        string sEtricc5RelativePath = string.Empty;

        string sEpiaBuildLogFile = string.Empty;
        string sEtriccUIBuildLogFile = string.Empty;
        string sEtricc5BuildLogFile = string.Empty;
   
		static int sLogCount    = 0;
		static int sLogInterval = 0;
		public DateTime sTestStartUpTime;

		public bool sIsDeployed = false;
		public static DateTime sDeploymentEndTime;

        public static string sMsgDebug = string.Empty;
		public static string sDemonstration = string.Empty;
		public static string sProjectFile = string.Empty;
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

		
	  
		// --- BUILD
		static string sTFSServerUrl = string.Empty;
		//TeamFoundationServer TFS;
        TfsTeamProjectCollection tfsProjectCollection = null;
		IBuildServer m_BuildSvc;
		private bool TFSConnected = true;

		private const string REGKEY = "Software\\Egemin\\Automatic testing\\";

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

		private System.Diagnostics.Process procLauncher;
		private System.Diagnostics.Process procExplorer;

		#endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Events of Tester (1)
		public event EventHandler OnLoggingChanged;
		#endregion // —— Events •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Constructors/Destructors/Cleanup of Tester (1)
		public Tester()
		{
			sTFSServerUrl = System.Configuration.ConfigurationManager.AppSettings.Get("TFSServer");
			sDemonstration = System.Configuration.ConfigurationManager.AppSettings.Get("Demonstration");
            sMsgDebug = System.Configuration.ConfigurationManager.AppSettings.Get("MsgDebug");
  
			string logFilename = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + "-" 
				+ System.Configuration.ConfigurationManager.AppSettings.Get("LogFilename");
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

			mPreviousSetupPathKC = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\KC\\Previous";
			mCurrentSetupPathKC = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\KC\\Current";

			mPreviousSetupPathEwms = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Ewms\\Previous";
			mCurrentSetupPathEwms = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY + "\\Setup\\Ewms\\Current";

			mTestRunsDirectory = ConstCommon.ETRICC_TESTS_DIRECTORY + "\\TestRuns\\bin\\Debug";

			if (!Directory.Exists(mPreviousSetupPathEpia))
				Directory.CreateDirectory(mPreviousSetupPathEpia);

			if (!Directory.Exists(mCurrentSetupPathEpia))
				Directory.CreateDirectory(mCurrentSetupPathEpia);
			//------------------------------------------------------
			if (!Directory.Exists(mPreviousSetupPathEtricc))
				Directory.CreateDirectory(mPreviousSetupPathEtricc);

			if (!Directory.Exists(mCurrentSetupPathEtricc))
				Directory.CreateDirectory(mCurrentSetupPathEtricc);
			//------------------------------------------------------
			if (!Directory.Exists(mPreviousSetupPathEtricc5))
				Directory.CreateDirectory(mPreviousSetupPathEtricc5);

			if (!Directory.Exists(mCurrentSetupPathEtricc5))
				Directory.CreateDirectory(mCurrentSetupPathEtricc5);
			//------------------------------------------------------
			if (!Directory.Exists(mPreviousSetupPathEwms))
				Directory.CreateDirectory(mPreviousSetupPathEwms);

			if (!Directory.Exists(mCurrentSetupPathEwms))
				Directory.CreateDirectory(mCurrentSetupPathEwms);
			
			if (!Directory.Exists(mTestRunsDirectory))
				Directory.CreateDirectory(mTestRunsDirectory);

			if (!Directory.Exists(ConstCommon.ETRICC_TESTS_DIRECTORY))
				Directory.CreateDirectory(ConstCommon.ETRICC_TESTS_DIRECTORY);

			// Get build store
			try
			{
				sIsDeployed = false;
				sDeploymentEndTime = DateTime.Now;
				m_TestPC = System.Environment.MachineName;
				
				if (TFSConnected)
				{
					Log("Connect to TFS");

                    Uri serverUri = new Uri(sTFSServerUrl);
                    System.Net.ICredentials tfsCredentials
                        = new System.Net.NetworkCredential("TfsBuild", "Egemin01", "TeamSystems.Egemin.Be");

                    tfsProjectCollection
                        = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                    tfsProjectCollection.Connect(Microsoft.TeamFoundation.Framework.Common.ConnectOptions.IncludeServices);
                    TfsConfigurationServer tfsConfigurationServer = tfsProjectCollection.ConfigurationServer;

                    m_BuildSvc = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));
				}

				//TeamFoundationServer TFS = TeamFoundationServerFactory.GetServer(System.Configuration.ConfigurationManager.AppSettings.Get("TFSServer"));
                WorkItemStore store = (WorkItemStore)tfsProjectCollection.GetService(typeof(WorkItemStore));
				//WorkItemType wiType = store.Projects[8].WorkItemTypes[1];
				// project nedds to be checked 
				//MessageBox.Show("store projects: " + store.Projects[8].ToString() + "  ");

				//WorkItem newWI = new WorkItem(wiType);
				//newWI.Title = "AddNewWorkItem";
				//newWI.State = "Active";
				//newWI.Fields["System.assignedTo"].Value = "Jiemin Shi";
				//newWI.Save();

				int ret = Disconnect(m_MapNetworkDrive_Root);
				if (ret == 0)
				{
					logger.LogMessageToFile(m_TestPC + "Disconnect MAP DRIVE OK:", sLogCount, sLogInterval);
				}
				else if (ret == 2250 )
					logger.LogMessageToFile(m_TestPC 
						+ "Disconnnet: MAP DRIVE The Network connection could not be found :"+ret, 
						sLogCount, sLogInterval);
				else
					System.Windows.MessageBox.Show("Disconnect  DriveMap failed with error code:" + ret);

				Thread.Sleep(3000);

				ret = OpenDriveMap(@"\\Teamsystem\Team Systems Builds", m_MapNetworkDrive_Root);
				if (ret == 0)
				{
					logger.LogMessageToFile(m_TestPC + "Open MAP DRIVE OK:", sLogCount, sLogInterval);
				}
				else if (ret == 85 )
					logger.LogMessageToFile(m_TestPC + "Open MAP DRIVE not connected due to existing connection:", sLogCount, sLogInterval);
				else
					System.Windows.MessageBox.Show("OpenDriveMap failed with error code:" + ret);
			}           
			catch (TeamFoundationServerUnauthorizedException ex1)
			{
				System.Windows.MessageBox.Show(ex1.Message + System.Environment.NewLine + ex1.StackTrace, "Tester Constructor");
				Log(ex1.Message + System.Environment.NewLine + ex1.StackTrace);
				TFSConnected = false;
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(ex.Message + System.Environment.NewLine + ex.StackTrace, "Tester Constructor");
				Log(ex.Message + System.Environment.NewLine + ex.StackTrace);
				TFSConnected = false;
			}
		}
		#endregion // —— Constructors/Destructors/Cleanup •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

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

		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Methods of Tester (14)
		public string GetTFSServerName()
		{
			return sTFSServerUrl;
		}

		public string getLogPath()
		{
			return sLogFilename;
		}
		/// <summary>
		/// Method will start new tests
		/// </summary>
		public void Start(ref DateTime StartUpTime)
		{
            if (m_Settings.PlatformTarget.Equals("Any CPU"))
            {   
                sEpiaRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaRelativePath");
                sEtriccUIRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIRelativePath");
                sEtricc5RelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5RelativePath");

                sEpiaBuildLogFile = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaBuildLogFile");
                sEtriccUIBuildLogFile = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIBuildLogFile");
                sEtricc5BuildLogFile = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5BuildLogFile");

            }
            else if (m_Settings.PlatformTarget.Equals("x86"))
            {
                sEpiaRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Epiax86RelativePath");
                sEtriccUIRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx86RelativePath");
                sEtricc5RelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5x86RelativePath");

                sEpiaBuildLogFile = System.Configuration.ConfigurationManager.AppSettings.Get("Epiax86BuildLogFile");
                sEtriccUIBuildLogFile = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx86BuildLogFile");
                sEtricc5BuildLogFile = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5x86BuildLogFile");
            }
            else if (m_Settings.PlatformTarget.Equals("x64"))
            {
                sEpiaRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Epiax64RelativePath");
                sEtriccUIRelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx64RelativePath");
                sEtricc5RelativePath = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5x64RelativePath");

                 sEpiaBuildLogFile = System.Configuration.ConfigurationManager.AppSettings.Get("Epiax64BuildLogFile");
                sEtriccUIBuildLogFile = System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIx64BuildLogFile");
                sEtricc5BuildLogFile = System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5x64BuildLogFile");
            }
            else
            {
                MessageBox.Show("Wrong Platform TArget:" + m_Settings.PlatformTarget);
            }

			Start(string.Empty, ref StartUpTime);
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
		public void Start(string BuildPath, ref DateTime upTime)
		{
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
				+ System.Configuration.ConfigurationManager.AppSettings.Get("LogFilename");
			if (!logFilename.Equals(logger.GetLogPath()))   // if another day, create new log file
				logger = new Logger(System.IO.Path.Combine(ConstCommon.ETRICC_TESTS_DIRECTORY, logFilename));

			logger.LogMessageToFile("Start:" + sLogCount, sLogCount, sLogInterval);
			Log("===> searching for available build...");

			m_State = STATE.INPROGRESS;

			if ( CheckTFSConnection() == false)
            //if (CheckTFSConnection() == false && m_TestAutoMode == true )
			{
				Log("No connection to TFS");
				logger.LogMessageToFile("No connection to TFS", sLogCount, sLogInterval);
				ClickUiScreenActionToAvoidScreenStandBy();
				sLogCount++;
				m_State = STATE.PENDING;
				return;
			}

			#region // Get All Builds From TFS depend on the setting from configuration file
			if (BuildPath == string.Empty)
			{
				// should check build quality
				List<string> allBuilds = getAllBuildDirectorys(m_Settings, m_MapNetworkDrive_Root);
				if (allBuilds.Count == 0)
				{
					Log("No Any build dircetory found:");
					logger.LogMessageToFile("No Any build dircetory found:", sLogCount, sLogInterval);
					ClickUiScreenActionToAvoidScreenStandBy();
					sLogCount++;
					m_State = STATE.PENDING;
					return;
				}

				// Get Validated build , that can be tested by thisPC
				// X:\Nightly\Etricc 5\Etricc - Nightly_20100202.1
                //MessageBox.Show("m_Settings.BuildApplication" + m_Settings.BuildApplication);
                m_ValidatedBuildDirectory = GetValidatedBuildDirectory(allBuilds, m_TestPC, m_Settings.BuildApplication, m_Settings.Branch, out m_testApp);
				if (m_ValidatedBuildDirectory == null)
				{
					logger.LogMessageToFile("No new build found:", sLogCount, sLogInterval);
					ClickUiScreenActionToAvoidScreenStandBy();
					m_State = STATE.PENDING;
					sLogCount++;
					return;
				}

				//X:\CI\Etricc 5\Etricc - CI_20100301.1
				logger.LogMessageToFile("===== <This build will be tested>===== >" + m_ValidatedBuildDirectory, 0, 0);
				// Tested buildnr Etricc - Nightly_20100202.1
				m_BuildNumber = BuildUtilities.getBuildnr(m_ValidatedBuildDirectory);
				Log("testing  build nr: " + m_BuildNumber);
				logger.LogMessageToFile("===== <testing  build nr > m_BuildNumber=: " + m_BuildNumber, 0, 0);

				//m_testApp = BuildUtilities.getTestApplication(m_ValidatedBuildDirectory);
                /*if (m_Settings.BuildApplication.Equals("Etricc UI"))
                    m_testApp = m_Settings.BuildApplication;
                else if (m_Settings.BuildApplication.Equals("Epia"))
                    m_testApp = m_Settings.BuildApplication;
                else if (m_Settings.BuildApplication.Equals("Etricc 5"))
                    m_testApp = m_Settings.BuildApplication;
                */
				Log("testing  application : " + m_testApp);
				logger.LogMessageToFile("===== <testing  application > m_testApp =: " + m_testApp, 0, 0);

				// Tested build type. Nightly
				m_testDef = BuildUtilities.getTestDefinition(m_ValidatedBuildDirectory);
				Log("testing definition : " + m_testDef);
				logger.LogMessageToFile("===== <testing  Definition > m_testDef=: " + m_testDef, 0, 0);
				//MessageBox.Show("ValidatedBuildDirectory:" + m_ValidatedBuildDirectory + "--testApp-" + m_testApp + " --testDef-" + m_testDef, "m_BuildNumber:" + m_BuildNumber);

				m_TestAutoMode = true;
				
			}
			else
			{
				Log(" manual test starting : " + BuildPath);
                MessageBox.Show("BuildPath:" + BuildPath);
				m_MapNetworkDrive_Root = System.IO.Path.GetPathRoot(BuildPath);
                MessageBox.Show("m_MapNetworkDrive_Root:" + m_MapNetworkDrive_Root);
				m_ValidatedBuildDirectory = BuildUtilities.getBuildBasePath(BuildPath);
                MessageBox.Show("m_ValidatedBuildDirectory:" + m_ValidatedBuildDirectory);
				m_installScriptDir = BuildPath;
				m_BuildNumber = BuildUtilities.getBuildnr(BuildPath);
                MessageBox.Show("m_BuildNumber:" + m_BuildNumber);
				
                
                m_testApp = BuildUtilities.getTestApplication(m_ValidatedBuildDirectory);
                if (m_Settings.BuildApplication.Equals("Etricc UI"))
                    m_testApp = m_Settings.BuildApplication;
                else if (m_Settings.BuildApplication.Equals("Epia"))
                    m_testApp = m_Settings.BuildApplication;
                else if (m_Settings.BuildApplication.Equals("Etricc 5"))
                    m_testApp = m_Settings.BuildApplication;
                MessageBox.Show("m_testApp:" + m_testApp);
				
                m_testDef = BuildUtilities.getTestDefinition(m_ValidatedBuildDirectory);
                MessageBox.Show("m_testDef:" + m_testDef);
				Log(" manual testing : " + m_BuildNumber);
				logger.LogMessageToFile(" manual testing : " + m_BuildNumber, 0, 0);
				Log(" m_MapNetWorkDrive_Root : " + m_MapNetworkDrive_Root);
				logger.LogMessageToFile(" root : " + m_MapNetworkDrive_Root, 0, 0);

				m_TestAutoMode = false;
				TFSConnected = false;
			}

			if (TFSConnected)
			{
				//m_Uri = m_buildStore.GetBuildUri(ProjectName(m_testApp), m_BuildNumber);
				m_Uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, TestTools.TfsUtilities.GetProjectName(m_testApp), m_BuildNumber);
				//MessageBox.Show("m_Uri_1" + m_Uri, "m_Uri_2" + m_Uri_2);
			}
			#endregion

			// Prepare deployment
            //MessageBox.Show("m_testApp:" + m_testApp, "prepare deployment 430");
			#region // Prepare deployment
			// We have BuildNumber now, now Deploy application by check m_Settings.Application
			if ( m_testApp.Equals(Constants.KIMBERLY_CLARK))
			{
				m_installScriptDir = m_ValidatedBuildDirectory + @"\\Debug\\Installation\\Setup\\";

				Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_TOOLS_DATABASE_FILLER);
			}
			else if (m_testApp.Equals(Constants.ETRICC_5) )
			{
                m_installScriptDir = m_ValidatedBuildDirectory +  sEtricc5RelativePath;
				Utilities.CloseProcess("EPIA.Launcher");
				Utilities.CloseProcess("EPIA.Explorer");
			}
            else if (m_testApp.Equals(Constants.ETRICC_UI) )
			{
                m_installScriptDir = m_ValidatedBuildDirectory + sEtriccUIRelativePath;
                
				Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_TOOLS_DATABASE_FILLER);

            }
			else if (m_testApp.Equals(Constants.EPIA) )
			{
                m_installScriptDir = m_ValidatedBuildDirectory + sEpiaRelativePath;
             
				Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_TOOLS_DATABASE_FILLER);
			}
			else
			{
				System.Windows.MessageBox.Show("Unknown Application, try other application again...   " + m_testApp);
				return;
			}
			#endregion

			#region    // Update build quality   "Deployment Started"
			if (TFSConnected)
			{
                string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(m_testApp), "Deployment Started", m_BuildSvc, sDemonstration);
                
                Log(updateResult);
                if (updateResult.StartsWith("Error"))
                {
                    MessageBox.Show(updateResult, "Update quality error");
                    throw new Exception(updateResult);
                }

				//UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                if (m_TestAutoMode)
                {
                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, "Deployment Started", m_testApp);
                    logger.LogMessageToFile(m_ValidatedBuildDirectory + " Deployment Started : m_testApp " + m_testApp, 0, 0);
                }
			}
			#endregion

			Thread.Sleep(3000);
			string installPath = m_CurrentDrive;   // only Etricc 5 depende on Drive, all other apps are on C:
			string dir = string.Empty;
			// Start Deployment   .......
			try
			{
                sEpia4InstallerName = System.Configuration.ConfigurationManager.AppSettings.Get("Epia4InstallerName");

				if ( m_testApp.Equals(Constants.KIMBERLY_CLARK))
				{    
					#region //KC
					if (System.IO.Directory.Exists(m_SystemDrive+"Program Files\\Egemin\\Etricc Server"))
					{
						//System.Windows.MessageBox.Show(m_SystemDrive + "Program Files\\Egemin\\Etricc Server");
						System.IO.Directory.Delete(m_SystemDrive + "Program Files\\Egemin\\Etricc Server", true);
					}
					logger.LogMessageToFile("<--------> Start KC depmoyment:" + m_BuildNumber, 0, 0);
					System.Threading.Thread.Sleep(2000);

					//  Install setup
					//Install new setup
					mTestedVersion = GetTestedVersion(m_ValidatedBuildDirectory, string.Empty);
					Log(" Tested Version :" + mTestedVersion);
					logger.LogMessageToFile(" Tested Version :" + mTestedVersion, 0, 0);

					mEpiaPath = m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT + "\\";
					//System.Windows.MessageBox.Show(" Install path :" + mEpiaPath);
					Log(" Install path :" + mEpiaPath);
					logger.LogMessageToFile(" Install path :" + mEpiaPath, 0, 0);

					// Remove Old SetUP
					RemoveSetup(Constants.ETRICCUI);
					RemoveSetup(Constants.KC+"TestProgram");
					RemoveSetup(Constants.KC);
					RemoveSetup(Constants.EPIA);

					//Move the current Setup files to a backup location
					if (!CopySetup(mCurrentSetupPathKC, mPreviousSetupPathKC))
						return;

					if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathKC, "Ewcs*.msi"))
						return;

					if (!CopySetup(mCurrentSetupPathEtricc, mPreviousSetupPathEtricc))
						return;

					if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtricc, "*Shell.msi"))
						return;

					if (!CopySetup(mCurrentSetupPathEpia, mPreviousSetupPathEpia))
						return;

                    if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia, "*" + sEpia4InstallerName))
						return;

					// remove Ewcs Service in case this service is still exist
					System.Diagnostics.Process procRemoveService = new System.Diagnostics.Process();
					procRemoveService.EnableRaisingEvents = false;
					procRemoveService.StartInfo.FileName = "sc";
					procRemoveService.StartInfo.Arguments = "delete " + '"' + "Ewcs Service" + '"';
					procRemoveService.StartInfo.WorkingDirectory = m_TestWorkingDirectory;
					procRemoveService.Start();
					procRemoveService.WaitForExit();

					if (System.IO.Directory.Exists(m_SystemDrive+ConstCommon.EPIA_SERVER_ROOT))
					{
						System.IO.Directory.Delete(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT, true);
						System.IO.Directory.CreateDirectory(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT);
					}

					// Install Current Setup
					if (InstallEpiaSetup(mCurrentSetupPathEpia, mTestedVersion))
					{
						if (System.IO.Directory.Exists(m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT))
						{
							System.IO.Directory.Delete(m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT, true);
							System.IO.Directory.CreateDirectory(m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT);

							FileManipulation.CopyDirectory(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
								m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT);
						}

						if (InstallEtriccUISetup(mCurrentSetupPathEtricc, mTestedVersion))
						{
							if (InstallKCSetup(mCurrentSetupPathKC, mTestedVersion))
							{
								if (InstallKCTestProgramSetup(mCurrentSetupPathKC, mTestedVersion)) 
								{
									FileManipulation.CopyDirectory(m_SystemDrive + @"Program Files\Egemin\Etricc Server",
											m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT);
								}
							}
						}
					}
					#endregion
				}
				else if (m_testApp.Equals(Constants.ETRICC_5))
				{
					#region // Etricc 5
					//MessageBox.Show("Start depmoyment:" + m_BuildNumber);
					logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
					System.Threading.Thread.Sleep(2000);

					//  Install setup
					//Install new setup and recompile Worker
					//mTestedVersion = GetTestedVersion(m_ValidatedBuildDirectory, string.Empty);
					//MessageBox.Show(" Tested Version :" + mTestedVersion);
					Log(" Tested Version :" + mTestedVersion);
					logger.LogMessageToFile(" Tested Version :" + mTestedVersion, 0, 0);

					mEpiaPath = m_SystemDrive + @"Program Files\Egemin\Epia " + mTestedVersion + "\\";
					//MessageBox.Show(" Install path :" + mEpiaPath);
					Log(" Install path :" + mEpiaPath);
					logger.LogMessageToFile(" Install path :" + mEpiaPath, 0, 0);

					// Remove Old SetUP
					RemoveSetup(Constants.ETRICC5);   

					//Move the current Setup files to a backup location
					if (!CopySetup(mCurrentSetupPathEtricc5, mPreviousSetupPathEtricc5))
						return;

					if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtricc5, "Etricc 5.msi"))
						return;

                    // check if current folder is empty
                    FileInfo[] FilesToCopy;
                    DirectoryInfo DirInfo = new DirectoryInfo(mCurrentSetupPathEtricc5);
                    FilesToCopy = DirInfo.GetFiles("Etricc 5.msi");
                    if (FilesToCopy.Length == 0)
                    {
                        #region    // Update build quality   "Deployment Failed"
                        if (TFSConnected)
                        {
                            string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(m_testApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, "Deployment Failed: no 'Etricc 5.msi' file found", m_testApp);
                                logger.LogMessageToFile(m_ValidatedBuildDirectory + " Deployment Failed : m_testApp " + m_testApp, 0, 0);
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
					
					if (InstallEtricc5Setup(mCurrentSetupPathEtricc5, mTestedVersion))
					{
						string cscOut = RecompileTestRuns();
						logger.LogMessageToFile("------ReCOMPILE TestRUNS output -------:"+cscOut, 0, 0);
					}
					#endregion
				}
				else if (m_testApp.Equals(Constants.EPIA) )
				{
					#region // Epia
					//MessageBox.Show("Start depmoyment:" + m_BuildNumber);
					logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
					System.Threading.Thread.Sleep(2000);

					if (System.IO.Directory.Exists(m_SystemDrive+"Program Files\\Egemin\\Epia Server"))
					{
						System.IO.Directory.Delete(m_SystemDrive + "Program Files\\Egemin\\Epia Server", true);
					}

					if (System.IO.Directory.Exists(m_SystemDrive + "Program Files\\Egemin\\Epia Shell"))
					{
						System.IO.Directory.Delete(m_SystemDrive + "Program Files\\Egemin\\Epia Shell", true);
					}

					//  Install Epia setup
					//Install new setup
					//mTestedVersion = GetTestedVersion(m_ValidatedBuildDirectory,  System.Configuration.ConfigurationManager.AppSettings.Get("EpiaBuildLogFile"));
					//MessageBox.Show(" Tested Version :" + mTestedVersion);
					Log(" Tested Version :" + mTestedVersion);
					logger.LogMessageToFile(" Tested Version :" + mTestedVersion, 0, 0);

					mEpiaPath = m_SystemDrive + @"Program Files\Egemin\Epia Server";
					// Clean up epia server folder
					if (System.IO.Directory.Exists(mEpiaPath))
					{
						logger.LogMessageToFile(" Clean up Server folder :", 0, 0);
						if (System.IO.Directory.Exists(mEpiaPath+"\\Data"))
							System.IO.Directory.Delete(mEpiaPath + "\\Data", true);
						if (System.IO.Directory.Exists(mEpiaPath + "\\Log"))
							System.IO.Directory.Delete(mEpiaPath + "\\Log", true);
					}

					//MessageBox.Show(" Install path :" + mEpiaPath);
					//MessageBox.Show(" mCurrentSetupPath :" + mCurrentSetupPathEpia);
					Log(" Install path :" + mEpiaPath);
					logger.LogMessageToFile(" Install path :" + mEpiaPath, 0, 0);

					// Remove Old Epia SetUP
					RemoveSetup(Constants.EPIA);  

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
                        if (TFSConnected)
                        {
                            string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(m_testApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, "Deployment Failed: no msi file found", m_testApp);
                                logger.LogMessageToFile(m_ValidatedBuildDirectory + " Deployment Failed : m_testApp " + m_testApp, 0, 0);
                            }
                        }
                        #endregion
                        return;
                    }

					// Install Current Setup
                    if (InstallEpiaSetup(mCurrentSetupPathEpia, mTestedVersion))
                    {
                        //MessageBox.Show(" Epia installed :" + mTestedVersion);
                        //string cscOut = RecompileTestRuns();
                        //logger.LogMessageToFile("------ReCOMPILE TestRUNS output -------:" + cscOut, 0, 0);
                    }
					#endregion
				}
				else if (m_testApp.Equals(Constants.ETRICC_UI) )
				{
					#region // Etricc
					//MessageBox.Show("Start depmoyment:" + m_BuildNumber);
					logger.LogMessageToFile("<--------> Start depmoyment:" + m_BuildNumber, 0, 0);
					System.Threading.Thread.Sleep(2000);

					//  Install setup
					//Install new setup
					//mTestedVersion = GetTestedVersion(m_ValidatedBuildDirectory, string.Empty);
					//MessageBox.Show(" Tested Version :" + mTestedVersion);
					Log(" Tested Version :" + mTestedVersion);
					logger.LogMessageToFile(" Tested Version :" + mTestedVersion, 0, 0);

					// Remove Old SetUP
					RemoveSetup(Constants.EPIA);
                    RemoveSetup(Constants.ETRICCUI);   // because the register is EtriccInstallation
					RemoveSetup(Constants.ETRICC5);
				  
					//Move the current Etricc Setup files to a backup location
					//---------------------------------------------------------------
					if (!CopySetup(mCurrentSetupPathEtricc, mPreviousSetupPathEtricc))
						return;

					//MessageBox.Show(" m_installScriptDir :" + m_installScriptDir);
					if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtricc, "*Shell.msi"))
						return;

                    // check if current folder is empty
                    FileInfo[] FilesToCopy;
                    DirectoryInfo DirInfo = new DirectoryInfo(mCurrentSetupPathEtricc);
                    FilesToCopy = DirInfo.GetFiles("*Shell.msi");
                    if (FilesToCopy.Length == 0)
                    {
                        #region    // Update build quality   "Deployment Failed"
                        if (TFSConnected)
                        {
                            string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(m_testApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, "Deployment Failed: no '*Shell.msi' file found", m_testApp);
                                logger.LogMessageToFile(m_ValidatedBuildDirectory + " Deployment Failed : m_testApp " + m_testApp, 0, 0);
                            }
                        }
                        #endregion
                        return;
                    }
					//---------------------------------------------------------------
					//Move the current Epia Setup files to a backup location
					if (!CopySetup(mCurrentSetupPathEpia, mPreviousSetupPathEpia))
						return;

					if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEpia,"*" + sEpia4InstallerName))
						return;

                    // check if current folder is empty
                    DirInfo = new DirectoryInfo(mCurrentSetupPathEpia);
                    FilesToCopy = DirInfo.GetFiles("*" + sEpia4InstallerName);
                    if (FilesToCopy.Length == 0)
                    {
                        #region    // Update build quality   "Deployment Failed"
                        if (TFSConnected)
                        {
                            string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(m_testApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, "Deployment Failed: no epia msi file found", m_testApp);
                                logger.LogMessageToFile(m_ValidatedBuildDirectory + " Deployment Failed : m_testApp " + m_testApp, 0, 0);
                            }
                        }
                        #endregion
                        return;
                    }
					//---------------------------------------------------------------
					//Move the current Etricc5 Setup files to a backup location
					if (!CopySetup(mCurrentSetupPathEtricc5, mPreviousSetupPathEtricc5))
						return;

					if (!CopySetupFilesWithWildcards(m_installScriptDir, mCurrentSetupPathEtricc5, "Etricc 5.msi"))
						return;

                    DirInfo = new DirectoryInfo(mCurrentSetupPathEtricc5);
                    FilesToCopy = DirInfo.GetFiles("Etricc 5.msi");
                    if (FilesToCopy.Length == 0)
                    {
                        #region    // Update build quality   "Deployment Failed"
                        if (TFSConnected)
                        {
                            string updateResult = TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(m_testApp), "Deployment Failed", m_BuildSvc, sDemonstration);

                            Log(updateResult);
                            if (updateResult.StartsWith("Error"))
                            {
                                MessageBox.Show(updateResult, "Update quality error");
                                throw new Exception(updateResult);
                            }

                            //UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Started", m_BuildSvc);
                            if (m_TestAutoMode)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, "Deployment Failed: no 'Etricc 5.msi' file found", m_testApp);
                                logger.LogMessageToFile(m_ValidatedBuildDirectory + " Deployment Failed : m_testApp " + m_testApp, 0, 0);
                            }
                        }
                        #endregion
                        return;
                    }
					if (System.IO.Directory.Exists(m_SystemDrive + "Program Files\\Egemin\\Epia Server"))
					{
						System.IO.Directory.Delete(m_SystemDrive + "Program Files\\Egemin\\Epia Server", true);
					}

					if (System.IO.Directory.Exists(m_SystemDrive + "Program Files\\Egemin\\Epia Shell"))
					{
						System.IO.Directory.Delete(m_SystemDrive + "Program Files\\Egemin\\Epia Shell", true);
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
					
					// Install Current Setup 
					if (InstallEpiaSetup(mCurrentSetupPathEpia, mTestedVersion))
					{
						//MessageBox.Show(" Epia installed :" + mTestedVersion);
						logger.LogMessageToFile(" Epia installed :" + mTestedVersion, 0, 0);
						Thread.Sleep(1000);
					}
					
					if (InstallEtricc5Setup(mCurrentSetupPathEtricc5, mTestedVersion))
					{
						//MessageBox.Show(" Etricc5 installed :" + mTestedVersion);
						logger.LogMessageToFile(" Etricc5 installed :" + mTestedVersion, 0, 0);
					}

					if (InstallEtriccUISetup(mCurrentSetupPathEtricc, mTestedVersion))
					{
						//MessageBox.Show(" Etricc installed :" + mTestedVersion);
						logger.LogMessageToFile(" Etricc installed :" + mTestedVersion, 0, 0);
					}

					mEpiaPath = sEtricc5InstallationFolder;
					Log(" Install path :" + mEpiaPath);
					logger.LogMessageToFile(" Install path :" + mEpiaPath, 0, 0);
					
					#region //   Check Project file
					if (m_testApp.Equals(Constants.ETRICC_UI))
					{
						string projectXml = m_Settings.SelectedProjectFile;
						string xmlPath = mEpiaPath +"\\Data\\Etricc\\Demo.xml";
						logger.LogMessageToFile("---- Check Project file:" + xmlPath, 0, 0);
						// Etricc deployment take long time than Epia, It copy demo.xml, check if copied
						while (!System.IO.File.Exists(xmlPath))
						{
							logger.LogMessageToFile("---- Not copied yet, wait:" + xmlPath, 0, 0);
							Thread.Sleep(3000);
						}

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
						else
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
					}
					#endregion
					#endregion
				}
				else
				{
					System.Windows.MessageBox.Show("Unknown Application, try other application again2..." + m_testApp);
					return;
				}

				// end deployment 
				System.Configuration.Configuration config =
					System.Configuration.ConfigurationManager.OpenExeConfiguration
					(System.Configuration.ConfigurationUserLevel.None);
				// Add an Application Setting.
				config.AppSettings.Settings.Add("LastDeploymentedBuild",
					   m_installScriptDir);
				// Save the changes in App.config file.
				config.Save(System.Configuration.ConfigurationSaveMode.Modified);
				// Force a reload of a changed section.
				System.Configuration.ConfigurationManager.RefreshSection("appSettings");


				#region // update build quality to "Deployment Completed"
				// only if this is first time test and build quality not = "GUI Tests Failed"
				if (TFSConnected)
				{
					Log(TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(m_testApp), "Deployment Completed",
						m_BuildSvc, sDemonstration));
					//UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "Deployment Completed", m_BuildSvc);
				}

				if (m_TestAutoMode)
                {
                    TestTools.FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, "Deployment Completed", m_testApp);
                    logger.LogMessageToFile(m_ValidatedBuildDirectory + "---   Deployment Completed : m_testApp " + m_testApp, 0, 0);
                }
				#endregion

				sDeploymentEndTime = DateTime.Now;
				sIsDeployed = true;

				// start testing
				// // clear test directory first : remove  xls file
				string deletePathXls = ConstCommon.ETRICC_TESTS_DIRECTORY + @"\*.xls";
				string msg = "Before testing start, first clear test directory-->delete files:" + deletePathXls;
				logger.LogMessageToFile(msg, 0, 0);
				if (!FileManipulation.DeleteFilesWithWildcards(deletePathXls, ref msg))
					throw new Exception(msg);

				// 
				if (m_testApp.Equals(Constants.EPIA) ||
					m_testApp.Equals(Constants.ETRICC_UI) ||
					m_testApp.Equals(Constants.KIMBERLY_CLARK))
				{
					string ExcelVisible = m_Settings.ExcelVisible;
					string AllowFunctionalTesting = m_Settings.FunctionalTesting.ToString().ToLower();
					string ServerRunAs = m_Settings.ServerRunAs;
					string Demo = (string)ConfigurationManager.AppSettings["Demonstration"];
					string Mail = m_Settings.Mail.ToString().ToLower();

					string arg = '"' + m_installScriptDir + '"'         //  0
						+ " " + '"' + ExcelVisible + '"'                //  1
						+ " " + '"' + m_TestAutoMode.ToString().ToLower() + '"'
						+ " " + '"' + ServerRunAs + '"'                 //  3
						+ " " + '"' + sDemonstration.ToString().ToLower() + '"'
						+ " " + '"' + Mail + '"'                        //  5
						+ " " + '"' + TESTTOOL_VERSION + '"'             //  6
						+ " " + '"' + sTFSServerUrl + '"'                   //  7
                        + " " + '"' + m_ValidatedBuildDirectory + '"'                   //  8
						+ " " + '"' + "AutoDeployment" + '"';           //  9

					dir = m_TestWorkingDirectory;
					string filename = string.Empty;
					if (m_testApp.Equals(Constants.EPIA))
					{                  
						#region
						filename = "Egemin.Epia.Testing.UIAutoTest.exe";
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
					else if (m_testApp.Equals(Constants.ETRICC_UI))
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
						arg = arg + " "
							+ '"' + AllowFunctionalTesting + '"' + " "   //  10
							+ '"' + sProjectFile + '"' + " "            //  11
							+ '"' + sEtricc5InstallationFolder + '"';  //  12

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
					else if (m_testApp.Equals(Constants.KIMBERLY_CLARK))
					{
						#region // KIMBERLY_CLARK directory
						filename = "Egemin.Epia.Testing.KimberlyClarkGUIAutoTest.exe";
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
							dir = dir + "\\KimberlyClarkGUIAutoTest\\bin\\Debug";
						}

						// remove Ewcs database 
						string dropEwcs = ConfigurationManager.AppSettings.Get("DropKCEwcsDatabase").ToString().ToLower();
						if (dropEwcs.StartsWith("true"))     
						{
							// first check Ewcs exist in sql file
							logger.LogMessageToFile("remove Ewcs database: working:dir:" + dir, 0, 0);
							TestTools.Utilities.StartProcessWaitForExit(dir, "DropDatabase.bat", 
								string.Empty);
						}

						arg = arg + " "
							+ '"' + dropEwcs + '"';  

						#endregion
					}

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
				else if (m_testApp.Equals(Constants.ETRICC_5))
				{
					#region  //Start Etricc 5 Testing
					//MessageBox.Show("Start Etricc 5 Test:" + m_BuildNumber);
					logger.LogMessageToFile("Start Etricc 5 Test:" + m_BuildNumber, 0, 0);

					// unzip project file, only test Eurobaltic Project
					string projectXml = m_Settings.SelectedProjectFile;
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
					if (TFSConnected)
					{
						Log(TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(m_testApp), "GUI Tests Started",
							m_BuildSvc, sDemonstration));
						//UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), "GUI Tests Started", m_BuildSvc);
						string quality2 = m_BuildSvc.GetMinimalBuildDetails(m_Uri).Quality;
						logger.LogMessageToFile(")))))))))))) start test worker:::::::::::::::::::quality:" + quality2, 0, 0);
					}

					if (m_TestAutoMode)
                    {
                        FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, "GUI Tests Started", m_testApp);
                        logger.LogMessageToFile(m_ValidatedBuildDirectory + "GUI Tests Started : m_testApp " + m_testApp, 0, 0);
                    }

					#endregion

					//   TestWorker starting... 
					logger.LogMessageToFile(")))))))))))) start test worker:::::::::::::::::::cnt:", 0, 0);
					StartTestWorker(sEtricc5InstallationFolder, ConstCommon.ETRICC_TESTS_DIRECTORY, sProjectFile);

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
						while ( cnt == 1 && exceptionMsg.Length == 0 && Time.Minutes <= 10)
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
					if (TFSConnected)
					{
						Log(TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(m_testApp), GUIstatus,
							m_BuildSvc, sDemonstration));
						//UpdateBuildQualityStatus(m_Uri, ProjectName(m_testApp), GUIstatus, m_BuildSvc);
					}

					if (m_TestAutoMode)
					{
						if (exceptionMsg.Length > 10)
						{
                            FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, exceptionMsg, m_testApp);
                            logger.LogMessageToFile(m_ValidatedBuildDirectory + "exceptionMsg : m_testApp " + m_testApp, 0, 0);
						}
						else
                            FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, GUIstatus, m_testApp);
                            logger.LogMessageToFile(m_ValidatedBuildDirectory + "GUIstatus : m_testApp " + m_testApp, 0, 0);
					}
					#endregion

					Utilities.CloseProcess("EPIA.Launcher");
					Utilities.CloseProcess("EPIA.Explorer");

					if (m_TestAutoMode)             
						FileManipulation.UpdateTestWorkingFile(m_ValidatedBuildDirectory, "false");

					logger.LogMessageToFile(" **************** ( END ETRICC 5 TESTING )************************** ", 0, 0);
					#endregion
				}

			}
			catch (Exception ex)
			{
				// Your error handler here
				System.Windows.MessageBox.Show(ex.Message + System.Environment.NewLine + ex.StackTrace
					+ "-- m_installScriptDir" + m_installScriptDir,
					"Started dir:" + dir);
				Log(ex.Message + System.Environment.NewLine + ex.StackTrace);

				// Update build quality
				if (TFSConnected)
				{
					Log(TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, m_Uri, TestTools.TfsUtilities.GetProjectName(m_testApp), "Deployment Failed",
						m_BuildSvc,sDemonstration));

					if (m_TestAutoMode)
					{
                        if (ex.Message.IndexOf("RecompileTestRuns") >= 0)
                        {
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, "Recompile TestRuns Exception:"
                                + ex.Message + "---" + ex.StackTrace, m_testApp);

                            logger.LogMessageToFile(m_ValidatedBuildDirectory + "Recompile TestRuns Exception: : m_testApp " + m_testApp, 0, 0);
                        }
                        else
                        {
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(m_ValidatedBuildDirectory, "Deployment Exception:"
                                + ex.Message + "---" + ex.StackTrace, m_testApp);
                            logger.LogMessageToFile(m_ValidatedBuildDirectory + "Deployment Exception: : m_testApp " + m_testApp, 0, 0);
                        }
					}
				}

				// Set TestWorking to false
				FileManipulation.UpdateTestWorkingFile(m_ValidatedBuildDirectory, "false");
			}
			finally
			{
				m_State = STATE.PENDING;
			}
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

		public bool CheckTFSConnection()
		{
			try
			{
				Log("reconnect to TFS");
                Uri serverUri = new Uri(sTFSServerUrl);
                System.Net.ICredentials tfsCredentials
                    = new System.Net.NetworkCredential("TfsBuild", "Egemin01", "TeamSystems.Egemin.Be");

                tfsProjectCollection
                    = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                tfsProjectCollection.Connect(Microsoft.TeamFoundation.Framework.Common.ConnectOptions.IncludeServices);
                TfsConfigurationServer tfsConfigurationServer = tfsProjectCollection.ConfigurationServer;

                m_BuildSvc = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));
			   
			}           
			catch (TeamFoundationServerUnauthorizedException ex)
			{
				Log(ex.Message + System.Environment.NewLine + ex.StackTrace);
				TFSConnected = false;
				return false;
			}
			catch (Exception ex)
			{
				Log(ex.Message + System.Environment.NewLine + ex.StackTrace);
				TFSConnected = false;
				return false;
			}

			if (tfsProjectCollection == null || m_BuildSvc == null)
				TFSConnected = false;
			else
				TFSConnected = true;

			return TFSConnected;

		}

		public List<string> getAllBuildDirectorys(Settings m_Settings, string root)
		{
			// only here use m_Settings.BuildApplication, after get m_ValidatedDirectory --> use m_testApp
			string buildApp = m_Settings.BuildApplication;
			List<string> searchDirs = new List<string>();

            //MessageBox.Show("1293    m_Settings.BuildApplication: " + m_Settings.BuildApplication);
			switch (m_Settings.BuildApplication)
			{
                    
				case Constants.EPIA:
                    if (m_Settings.BuildDefCI)  //  value="\\CI\\Epia 4\\Epia\\Epia - CI\\"
                    {
                        if (m_Settings.Branch.ToLower().StartsWith("main"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiAllBuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev01"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev01BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev02"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev02BuildStam"));
                        else if (m_Settings.Branch.ToLower().Equals("dev03"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev03BuildStam"));
                        else if (m_Settings.Branch.ToLower().Equals("dev03-net4"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev03-Net4BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev07"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev07BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev08"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev08BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("all"))
                        {
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiAllBuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev01BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev02BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev03BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev03-Net4BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev07BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiDev08BuildStam"));
                        }
                        else
                            MessageBox.Show("Epia CI Branch unknown");
                    }
                    if (m_Settings.BuildDefNightly) //  value="\\Nightly\\Epia 4\\Epia\\Epia - Nightly\\"
                    {
                        if (m_Settings.Branch.ToLower().StartsWith("main"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightAllBuildStam"));
                         else if (m_Settings.Branch.ToLower().StartsWith("dev01"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightDev01BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev02"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightDev02BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev03"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightDev03BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev07"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightDev07BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev08"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightDev08BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("all"))
                        {
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightAllBuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightDev01BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightDev02BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightDev03BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightDev07BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightDev08BuildStam"));
                        }
                        else
                            MessageBox.Show("Epia Night Branch unknown");
                    }
					if (m_Settings.BuildDefWeekly)
						searchDirs.Add(root + "\\Weekly\\Epia 3\\Epia");
                    if (m_Settings.BuildDefVersion) //  value="\\Version\\Epia 4\\Epia\\Epia - Version\\"
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaVersionAllBuildStam"));
					break;
				case Constants.ETRICC_UI:
                    if (m_Settings.BuildDefCI)  //  value= "\\CI\\Etricc 5\\Etricc.Main.CI\\"
                    {
                        if (m_Settings.Branch.ToLower().StartsWith("main"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiAllBuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev01"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiDev01BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev02"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiDev02BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev03"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiDev03BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev04"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiDev04BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("all"))
                        {
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiAllBuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiDev01BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiDev02BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiDev03BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiDev04BuildStam"));
                        }
                    }
                    if (m_Settings.BuildDefNightly)  // value= "\\Nightly\\Etricc 5\\Etricc.Main.Nightly\\"
                    {
                        if (m_Settings.Branch.ToLower().StartsWith("main"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightAllBuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev01"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightDev01BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev02"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightDev02BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev03"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightDev03BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev04"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightDev04BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("all"))
                        {
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightAllBuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightDev01BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightDev02BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightDev03BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightDev04BuildStam"));
                        }
                    }
					if (m_Settings.BuildDefWeekly)
						searchDirs.Add(root + "\\Weekly\\Epia 3\\Etricc");
                    if (m_Settings.BuildDefVersion)   //  value= "\\Version\\Epia 3\\Etricc UI\\Etricc UI - Version\\"
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIVersionAllBuildStam"));
					break;
				case Constants.ETRICC_5:
                    if (m_Settings.BuildDefCI)   //  value=  "\\CI\\Etricc 5"
                    {
                        if (m_Settings.Branch.ToLower().StartsWith("main"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiAllBuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev01"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiDev01BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev02"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiDev02BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev03"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiDev03BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev04"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiDev04BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("all"))
                        {
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiAllBuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiDev01BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiDev02BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiDev03BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiDev04BuildStam"));
                        }
                    }
                    if (m_Settings.BuildDefNightly) //  value= "\\Nightly\\Etricc 5"
                    {
                        if (m_Settings.Branch.ToLower().StartsWith("main"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightAllBuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev01"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightDev01BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev02"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightDev02BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev03"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightDev03BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("dev04"))
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightDev04BuildStam"));
                        else if (m_Settings.Branch.ToLower().StartsWith("all"))
                        {
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightAllBuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightDev01BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightDev02BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightDev03BuildStam"));
                            searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightDev04BuildStam"));
                        }
                    }
					//if (m_Settings.BuildDefWeekly)    // not yet implemented in teamsystem
					//    searchDirs.Add(root + "\\Weekly\\Etricc 5");
                    if (m_Settings.BuildDefVersion) //  value= "\\Version\\Etricc 5"
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5VersionAllBuildStam"));
					break;
				case Constants.ETRICC_ETRICC5:
					if (m_Settings.BuildDefCI)
					{
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiAllBuildStam"));
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiAllBuildStam"));
					}
					if (m_Settings.BuildDefNightly)
					{
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightAllBuildSta"));
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightAllBuildStam"));
					}
					if (m_Settings.BuildDefWeekly)
					{
						searchDirs.Add(root + "\\Weekly\\Epia 3\\Etricc");
						//    searchDirs.Add(root + "\\Weekly\\Etricc 5");
					}
					if (m_Settings.BuildDefVersion)
					{
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIVersionAllBuildStam"));
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5VersionAllBuildStam"));
					}
					break;
				case Constants.KIMBERLY_CLARK:
					if (m_Settings.BuildDefCI)
						searchDirs.Add(root + "\\CI\\Ewcs - Projects\\Kimberly Clark");
					if (m_Settings.BuildDefNightly)
						searchDirs.Add(root + "\\Nightly\\Ewcs - Projects\\Kimberly Clark");
					//if (m_Settings.BuildDefWeekly)    // not yet implemented in teamsystem
					//    searchDirs.Add(root + "\\Weekly\\Ewcs - Projects\\Kimberly Clark");
					if (m_Settings.BuildDefVersion)
						searchDirs.Add(root + "\\Version\\Ewcs - Projects\\Kimberly Clark");
					break;
				case "All":
					if (m_Settings.BuildDefCI)
					{
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaCiAllBuildStam"));
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUICiAllBuildStam"));
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5CiAllBuildStam"));
						searchDirs.Add(root + "\\CI\\Ewcs - Projects\\Kimberly Clark");
					}
					if (m_Settings.BuildDefNightly)
					{
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaNightAllBuildStam"));
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUINightAllBuildStam"));
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5NightAllBuildStam"));
						searchDirs.Add(root + "\\Nightly\\Ewcs - Projects\\Kimberly Clark");
					}
					if (m_Settings.BuildDefWeekly)
					{
						searchDirs.Add(root + "\\Weekly\\Epia 3\\Epia");
						searchDirs.Add(root + "\\Weekly\\Epia 3\\Etricc");
						//    searchDirs.Add(root + "\\Weekly\\Etricc 5");
						//    searchDirs.Add(root + "\\Weekly\\Ewcs - Projects\\Kimberly Clark");
					}
					if (m_Settings.BuildDefVersion)
					{
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EpiaVersionAllBuildStam"));
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("EtriccUIVersionAllBuildStam"));
                        searchDirs.Add(root + System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5VersionAllBuildStam"));
						searchDirs.Add(root + "\\Version\\Ewcs - Projects\\Kimberly Clark");
					}
					break;
				default:
					break;
			}

			List<string> allBuilds = new List<string>();

			foreach (string s in searchDirs)
			{
                //if (sMsgDebug.StartsWith("true"))
                //    MessageBox.Show("Searching dir =" + s);
				try
				{
					string[] allDirs = System.IO.Directory.GetDirectories(s);
					// start from latest build
					for (int i = allDirs.Length - 1; i >= 0; i--)
					{
                        //MessageBox.Show("allDirs["+i+"]=" + allDirs[i]);
						// Check Build Quality
						if (IsValidTestBuildNr(allDirs[i]))
						{
							allBuilds.Add(allDirs[i]);
							//Log(allDirs[i] + " is valid Directory");
							//TODO add loging level
							//logger.LogMessageToFile(allDirs[i] + " is valid Directory");
						}
					}
				}
				catch (Exception ex)
				{
					logger.LogMessageToFile(s + ":getAllBuildDirectorys" + ex.Message + " --- " + ex.StackTrace, 0, 1);
				}
			}
			return allBuilds;
		}

		public string GetValidatedBuildDirectory(List<string> BuildDirs, string testPC, string testApp, string testBranch, out string current_testApp)
		{
            current_testApp = string.Empty;
			string validBuild = null;
			foreach (string s in BuildDirs)
			{
				string buildDirectory = s + "\\TestResults";
				if (!System.IO.Directory.Exists(buildDirectory))
					System.IO.Directory.CreateDirectory(buildDirectory);

				// check build Succeeding by check Building.txt has text "0 Error(s)"
				string errorMsg = "CheckBuildSucceeding:";

                string buildlogfile = string.Empty;
                if (testApp.Equals(Constants.EPIA))
                {
                    buildlogfile = s + sEpiaBuildLogFile; 
                    current_testApp = Constants.EPIA;
                }
                else if (testApp.Equals(Constants.ETRICC_UI))
                {
                    buildlogfile = s + sEtriccUIBuildLogFile;
                    current_testApp = Constants.ETRICC_UI;
                }
                else if (testApp.Equals(Constants.ETRICC_5))
                {
                    buildlogfile = s + sEtricc5BuildLogFile;
                    current_testApp = Constants.ETRICC_5;
                }
                else if (testApp.Equals(Constants.ETRICC_ETRICC5)) //  current_testApp will be decided next
                {
                    buildlogfile = s + sEtricc5BuildLogFile;
                }
                else
                {
                    MessageBox.Show("Wrong testApp:" + testApp);
                    current_testApp = "Wrong testApp:";
                }

                logger.LogMessageToFile(testPC + " check buildlog file:" + buildlogfile, sLogCount, 0);

                if (TestTools.FileManipulation.CheckSearchTextExistInFile(buildlogfile, "0 Error(s)", ref errorMsg))
				{
					// check TestInfo.txt and TestWorking.txt files
					string testInfoTxtFile = Path.Combine(buildDirectory, ConstCommon.TESTINFO_FILENAME);
					string testWorkingFile = Path.Combine(buildDirectory, ConstCommon.TESTWORKING_FILENAME);
					// TestInfo file exist
					if (File.Exists(testInfoTxtFile))
					{
                        if (sMsgDebug.StartsWith("true"))
                             MessageBox.Show("is test working:" + buildDirectory );

						if (IsTestWorking(buildDirectory) == false)
						{
                            bool BothTested = false; 
                            if (testApp.Equals(Constants.ETRICC_ETRICC5))
                            {
                                if (IsThisPCTested(testInfoTxtFile, testPC, Constants.ETRICC_UI) == false)
                                    current_testApp = Constants.ETRICC_UI;
                                else if (IsThisPCTested(testInfoTxtFile, testPC, Constants.ETRICC_5) == false)
                                    current_testApp = Constants.ETRICC_5;
                                else
                                {
                                    Log(testPC + " is Tested for EtriccUI and Etricc 5 thisbuild:" + s + " -- Build Definition:" + m_testDef);
                                    logger.LogMessageToFile(testPC + " is Tested for thisbuild: Etricc UI and Etricc 5 --" + s + " -- Build Definition:" + m_testDef,
                                        sLogCount, 0);
                                    BothTested = true;
                                }
                            }


                            if (IsThisPCTested(testInfoTxtFile, testPC, current_testApp) == false && BothTested == false )
							{
								// Add test info
								// Check test Working file
								FileInfo workFile = new FileInfo(testWorkingFile);
								File.SetAttributes(workFile.FullName, FileAttributes.Normal);

								StreamReader readerInfo = File.OpenText(testInfoTxtFile);
								string info = readerInfo.ReadToEnd();
								readerInfo.Close();
                                info = info + testPC + "+" + current_testApp + "==Starting";
								Log(testPC + " Added into InfoFile of build:" + s);
								logger.LogMessageToFile(testPC + " Added into Info file:" + s, sLogCount, sLogInterval);

								// ------------  if write infotext file failure, not do anything , continue go to iteration 
								try
								{
									StreamWriter writeInfo = File.CreateText(testInfoTxtFile);
									writeInfo.WriteLine(info);
									writeInfo.Close();
								}
								catch (Exception ex)
								{
									string msg = testPC + " Exception Add pc to InfoText:" + s + "=====" + ex.Message + "" + ex.StackTrace;
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

										Log(testPC + " Set Working to true:" + s);
										logger.LogMessageToFile(testPC + " Set Working to true:" + s, sLogCount, sLogInterval);
										updateTestWorking = true;
									}
									catch (Exception ex)
									{
										string msg = testPC + " Exception Set Working to true:" + s + "=====" + ex.Message + "" + ex.StackTrace;
										Log(msg);
										logger.LogMessageToFile(msg, sLogCount, sLogInterval);
										Thread.Sleep(10000);
										updateTestWorking = false;
									}
								}

								validBuild = s;
							}
							else
							{
								Log(testPC + " is Tested for thisbuild:" + s + " -- Build Definition:" + m_testDef);
								logger.LogMessageToFile(testPC + " is Tested for thisbuild:" + s + " -- Build Definition:" + m_testDef,
									sLogCount, sLogInterval);
							}
						}
						else
						{
							Log("Test is Working..." + s);
							logger.LogMessageToFile("Test is Working..." + s, sLogCount, sLogInterval);
						}
					}
					else // create testinfo file and create TestWorking File
					{
						StreamWriter writeInfo = File.CreateText(testInfoTxtFile);
                        if (testApp.Equals(Constants.ETRICC_ETRICC5))
                               current_testApp = Constants.ETRICC_5;
                        else
                            current_testApp = testApp;

                        writeInfo.WriteLine("TestInfo" + System.Environment.NewLine
                            + testPC + "+" + current_testApp + "=" + m_testDef + "==Starting");

                        writeInfo.Close();

						// Create testWorking File
						StreamWriter writeWork = File.CreateText(testWorkingFile);
						writeWork.WriteLine("true");
						writeWork.Close();
						Log(s + " Deployment starting... ");
						logger.LogMessageToFile(s + " Deployment starting...", sLogCount, sLogInterval);
						validBuild = s;
					}
				}
				else
				{
					Log(s + " build is not secceeded..." + errorMsg);
					logger.LogMessageToFile(s + " build is not secceeded..."+errorMsg, sLogCount, sLogInterval);
				}

				if (validBuild != null)
				{
					break;
				}
			}

			return validBuild;
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
		
		/// <summary>
		///  For valid build quality, Check if it has tested by current PC
		/// </summary>
		/// <param name="testInfoTxtFile"></param>
		/// <param name="testPC"></param>
		/// <param name="BuildDefinition: CI, Nightly, Weekly or Version"></param>
		/// <returns></returns>
		public bool IsThisPCTested(string testInfoTxtFile, string testPC, string testApp)
		{
			try
			{
				StreamReader readerInfo = File.OpenText(testInfoTxtFile);
				string info = readerInfo.ReadToEnd();
				readerInfo.Close();

				if (info.IndexOf(testPC +"+"+testApp + "=") > 0)
				{
					//Log("but "+testPC + " is already in test info file");
					//logger.LogMessageToFile(testPC + " is already in test info file", sLogCount, sLogInterval);
					return true;
				}
				else
				{
                    Log(testPC + "+" + testApp +  "="+" ---------->> is not in test info file");
                    logger.LogMessageToFile(testPC + "+" + testApp + "=" + " ----------> is not in test info file", 
                        sLogCount, sLogInterval);
					return false;
				}
			}
			catch (Exception ex)
			{
				Log("IsTestWorking exception:" + testPC + " - message:" + ex.Message + " --- " + ex.StackTrace);
                logger.LogMessageToFile("IsTestWorking exception:" + testPC + "+" + testApp + "=" + " - message:" + ex.Message + " --- " + ex.StackTrace,
					sLogCount, sLogInterval);
				return true;
			}
		}

		public string GetTestedVersion(string path, string fileName)
		{
			string version = string.Empty;


            if (m_Settings.BuildApplication.Equals("Etricc 5"))
                return "Etricc 5 has no Epia Version";

            if (m_Settings.BuildApplication.Equals("Etricc UI"))
                return "Etricc UI has no Epia Version";

            if (m_Settings.BuildApplication.Equals(ConstCommon.ETRICC_ETRICC5))
                return ConstCommon.ETRICC_ETRICC5+ " has no Epia Version";

            //path = @"X:\CI\Etricc 5\Etricc - CI_20100210.4";
			string searchText = "Egemin.EPIA.Definitions, Version=";
			string fileFullname = System.IO.Path.GetFullPath(path+fileName);
            //MessageBox.Show("fileFullname :"+ fileFullname);

            if (m_Settings.BuildApplication.Equals("Epia"))
                return "Epia has no Epia Version";

			string strFile;
			int MyPos = 1;
            StreamReader reader2 = null;
            try
            {
                reader2 = System.IO.File.OpenText(fileFullname);
                strFile = reader2.ReadToEnd();


                MyPos = strFile.IndexOf(searchText, MyPos + 1);
                if (MyPos > 0)
                {
                    //version = strFile.Substring(MyPos + 33, 5);
                    string versionStr = strFile.Substring(MyPos + 33, 20);

                    int pointCnt = 0;
                    int order = 0;
                    while (pointCnt < 3)
                    {
                        if (versionStr.Substring(order, 1).Equals("."))
                        {
                            pointCnt++;
                        }
                        order++;
                    }
                    version = versionStr.Substring(0, order - 1);
                }
                else
                    version = null;
            }
            catch (Exception ex)
            {
                Log(ex.Message + "----" + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + "----" + ex.StackTrace, "GetTestedVersion");
            }
            finally
            {
                reader2.Close();   
            }

			return version;
		}

		/// <summary>
		///  Check Valid build quality
		/// </summary>
		/// <param name="path">X:\CI\Etricc 5\Etricc - CI_20100223.3</param>
		/// <returns></returns>
		public bool IsValidTestBuildNr(string path)
		{
			// Build nr start type: Release, Epia, Etricc, KC
			try
			{
                Log("check path:" + path);
				//MessageBox.Show(path, "path");
				string buildnr = BuildUtilities.getBuildnr(path);
				string testApp = m_Settings.BuildApplication;
                Log("check buildnr:" + buildnr);
                Log("checktestApp:" + testApp);
				Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, TestTools.TfsUtilities.GetProjectName(testApp), buildnr);
				string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;
				if (quality == null)
				{
					logger.LogMessageToFile(buildnr + " has no quality value:", sLogCount, sLogInterval);
					return false;
				}
				else if (quality.Equals("Rejected")
					|| quality.Equals("Ready for Release")
					|| quality.Equals("Investigation Inconclusive")
					|| quality.Equals("Under Investigation")
					|| quality.Equals("No Tests Needed"))
				{
					logger.LogMessageToFile(buildnr + " has no valid quality:" + quality, sLogCount, sLogInterval);
					return false;
				}
				else
				{
					logger.LogMessageToFile(buildnr + " has valid quality:------------> " + quality, sLogCount, sLogInterval);
					return true;
				}
			}
			catch (Exception ex)
			{
				Log("IsValidTestBuildNr exception:" + path + " - message:" + ex.Message + " --- " + ex.StackTrace);
                logger.LogMessageToFile("IsValidTestBuildNr exception::" + path + " - message:" + ex.Message + " --- " + ex.StackTrace,
					sLogCount, sLogInterval);
				return false;
			}
		}

		public void LoadConfiguration()
		{
			m_Settings = Settings.GetSettings();
			Log("Configuration Loaded");
		}

		public void LoadConfiguration(string configpath)
		{
			m_Settings = Settings.GetSettings(configpath);
			Log("Configuration Loaded");
		}

		//---------------------------------------------------------------------
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

		public void SaveConfiguration()
		{
			Settings.SaveSettings(m_Settings);
			Log("Configuration Saved");
			logger.LogMessageToFile("Configuration Saved",0,0);
		}

		protected static void ThreadProc()
		{
			System.Windows.MessageBox.Show(m_BuildNumber, "start new build deployment");
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
				System.Windows.MessageBox.Show(" <-----> Window not found ", stepMsg);
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
						System.Windows.MessageBox.Show(" <-----> Button not found ;" + buttonName, stepMsg);
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

		public void InstallEtriccSetupByStep(string WindowName, string stepMsg, int step)
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
				System.Windows.MessageBox.Show(" <-----> Window not found ", stepMsg);
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
					WaitUntilMyButtonFoundInThisWindow(WindowName, "Close", 600);
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
						System.Windows.MessageBox.Show(" <-----> Button not found ;" + buttonName, stepMsg);
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
		
		public void InstallSetupByStep(string WindowName, string stepMsg, int step)
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
				System.Windows.MessageBox.Show(" <-----> Window not found ", stepMsg);  
			}
			else
			{
				if (step == 2)
				{
					AutomationElement aeIAgreeRadioButton
						= AUIUtilities.FindElementByName("I Agree", aeWindow);
					SelectionItemPattern itemRadioPattern = aeIAgreeRadioButton.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
					itemRadioPattern.Select();
					logger.LogMessageToFile("<---> <I Agree> selected ... ", sLogCount, sLogInterval);
					Thread.Sleep(3000);
				}

				//Select Installation Folder
				if (step == 5)
				{
                    Thread.Sleep(2000);
					/*AutomationElement aeInstallationFolder
						= AUIUtilities.FindElementByType(ControlType.Edit, aeWindow);

					TextPattern tp = (TextPattern)aeInstallationFolder.GetCurrentPattern(TextPattern.Pattern);
					Thread.Sleep(1000);
					sEtricc5InstallationFolder = tp.DocumentRange.GetText(-1).Trim();*/
                    //sEtricc5InstallationFolder = @"C:\Program Files\Egemin\Etricc Server\";

                    sEtricc5InstallationFolder = 
                        System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5InstallationFolder");
					//System.Windows.MessageBox.Show(sInstallationFolder, "installation folder ");
					Thread.Sleep(3000);
				}

				if (step == 7 || step == 8 || step == 9)
				{
                    AutomationElement aeTitleBar =
                         AUIUtilities.FindElementByID("TitleBar", aeWindow);

                    Point pt1 = new Point(
                        (aeTitleBar.Current.BoundingRectangle.Left + aeTitleBar.Current.BoundingRectangle.Right) / 2,
                        (aeTitleBar.Current.BoundingRectangle.Bottom + aeTitleBar.Current.BoundingRectangle.Top) / 2);

                    Point newPt1 = new Point(200, 100);
                    Input.MoveTo(pt1);

                    System.Threading.Thread.Sleep(1000);
                    Input.SendMouseInput(pt1.X, pt1.Y, 0, SendMouseInputFlags.LeftDown | SendMouseInputFlags.Absolute);
                    System.Threading.Thread.Sleep(1000);
                    Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.Move | SendMouseInputFlags.Absolute);
                    System.Threading.Thread.Sleep(1000);
                    Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);
                    System.Threading.Thread.Sleep(3000);
					Thread.Sleep(20000);

					AutomationElement aeForm = null;
					string FormLauncherFuncID = "FrmLauncherFunctionality";
					string FormSecurityFuncID = "FrmSecurityFunctionality";
					string FormID = FormLauncherFuncID;

					StartTime = DateTime.Now;
					Time = DateTime.Now - StartTime;
					int wt = 0;
					while (aeForm == null && wt < 300)
					{                        
						if (step == 7 || step == 8)
						{
							FormID = FormLauncherFuncID;
							logger.LogMessageToFile("<----->wait for form Launcher window : " + wt, sLogCount, sLogInterval);
							aeForm = AUIUtilities.FindElementByID(FormID, AutomationElement.RootElement);
						}
						else
						{
							FormID = FormSecurityFuncID;
							logger.LogMessageToFile("<----->wait for form Security window : " + wt, sLogCount, sLogInterval);
							aeForm = AUIUtilities.FindElementByID(FormID, AutomationElement.RootElement);
						}
						Thread.Sleep(1000);
						
						Time = DateTime.Now - StartTime;
						Thread.Sleep(2000);
						wt = wt + 2;
					}

					if (aeForm == null)
						logger.LogMessageToFile("<----->aeForm not found : ", sLogCount, sLogInterval);
					else
					{
						logger.LogMessageToFile("<----->aeForm found : " + aeForm.Current.AutomationId, sLogCount, sLogInterval);
						//WindowPattern wpForm = (WindowPattern)aeForm.GetCurrentPattern(WindowPattern.Pattern);
						//if (!wpForm.Current.IsTopmost)
						//{
						//	logger.LogMessageToFile("aeForm : " + FormID + " is not topMost", sLogCount, sLogInterval);
						//	throw new Exception("aeForm : " + FormID + " is not topMost");
						//}
					}
					if (step == 8 || step == 9)
					{
						AutomationElement aeNextButton = FindSpecificButtonByName(aeForm, "Next >");
						System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
						Thread.Sleep(1000);
						Input.MoveTo(OptionPt);
						Thread.Sleep(2000);
						logger.LogMessageToFile("<---> Next > button clicking ... ", sLogCount, sLogInterval);
						Input.ClickAtPoint(OptionPt);
						Thread.Sleep(5000);
					}

					if (step == 9)
					{
						logger.LogMessageToFile("<----->wait extra 15 second  : ", sLogCount, sLogInterval);
						WaitUntilMyButtonFoundInThisWindow(WindowName, "Close", 600);
						Thread.Sleep(3000);
					}

				}
				else
				{
					string buttonName = (step == 10) ? "Close" : "Next >";
					AutomationElement aeNextButton = null;
					StartTime = DateTime.Now;
					Time = DateTime.Now - StartTime;
					int wt = 0;
					while (aeNextButton == null && wt < 500)
					{
						aeNextButton = FindSpecificButtonByName(aeWindow, (step == 10) ? "Close" : "Next >");
						Time = DateTime.Now - StartTime;
						Thread.Sleep(2000);
						wt = wt + 2;
					}

					if (aeNextButton == null)
					{
						System.Windows.MessageBox.Show(" <-----> Button not found ;" + buttonName, stepMsg);  
					}
					else
					{
						System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
						Thread.Sleep(1000);
						Input.MoveTo(OptionPt);
						Thread.Sleep(2000);
						logger.LogMessageToFile("<---> "+buttonName+" button clicking ... ", sLogCount, sLogInterval);
						Input.ClickAtPoint(OptionPt);
						Thread.Sleep(2000);
					}
				}			   
			}
		}

		public void InstallKCSetupByStep(string WindowName, string stepMsg, int step)
		{
			logger.LogMessageToFile("<-----> InstallKCSetupByStep: " + step, sLogCount, sLogInterval);
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
				System.Windows.MessageBox.Show(" <-----> Window not found ", stepMsg);
			}
			else
			{
				if (step == 2)
				{
					AutomationElement aeEpiaServerExtCkb
						= AUIUtilities.FindElementByName("E'pia Server Extensions (Resource & Security Files)", aeWindow);
					TogglePattern tg = aeEpiaServerExtCkb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
					ToggleState tgState = tg.Current.ToggleState;
					if (tgState == ToggleState.Off)
						tg.Toggle();

					AutomationElement aeShellckb
						= AUIUtilities.FindElementByName("E'pia Shell Extensions (Shell Module & Config)", aeWindow);
					TogglePattern tg2 = aeShellckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
					ToggleState tg2State = tg2.Current.ToggleState;
					if (tg2State == ToggleState.Off)
						tg2.Toggle();

					AutomationElement aeEwcsServerckb
						= AUIUtilities.FindElementByName("E'wcs Server Components", aeWindow);
					TogglePattern tg3 = aeEwcsServerckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
					ToggleState tg3State = tg3.Current.ToggleState;
					if (tg3State == ToggleState.Off)
						tg3.Toggle();

					logger.LogMessageToFile("<---> Components Sever and shell checkbox state set on  ... ", sLogCount, sLogInterval);
					Thread.Sleep(3000);
				}


				if (step == 6)
				{
					logger.LogMessageToFile("<----->wait extra 15 second  : ", sLogCount, sLogInterval);
					WaitUntilMyButtonFoundInThisWindow(WindowName, "Close", 600);
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
						System.Windows.MessageBox.Show(" <-----> Button not found ;" + buttonName, stepMsg);
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

		public void InstallKCTestProgramSetupByStep(string WindowName, string stepMsg, int step)
		{
			logger.LogMessageToFile("<-----> InstallKCTestProgramSetupByStep: " + step, sLogCount, sLogInterval);
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
				System.Windows.MessageBox.Show(" <-----> Window not found ", stepMsg);
			}
			else
			{
				if (step == 2)
				{
					AutomationElement aeEpiaServerExtCkb
						= AUIUtilities.FindElementByName("E'pia Server Extensions (Resource & Security Files)", aeWindow);
					TogglePattern tg = aeEpiaServerExtCkb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
					ToggleState tgState = tg.Current.ToggleState;
					if (tgState == ToggleState.Off)
						tg.Toggle();

					AutomationElement aeShellckb
						= AUIUtilities.FindElementByName("E'pia Shell Extensions (Shell Module & Config)", aeWindow);
					TogglePattern tg2 = aeShellckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
					ToggleState tg2State = tg2.Current.ToggleState;
					if (tg2State == ToggleState.Off)
						tg2.Toggle();

					AutomationElement aeEwcsServerckb
						= AUIUtilities.FindElementByName("E'wcs Server Test Programs", aeWindow);
					TogglePattern tg3 = aeEwcsServerckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
					ToggleState tg3State = tg3.Current.ToggleState;
					if (tg3State == ToggleState.Off)
						tg3.Toggle();

					AutomationElement aeEtriccServerckb
						= AUIUtilities.FindElementByName("E'tricc Server Test Programs", aeWindow);
					TogglePattern tg4 = aeEtriccServerckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
					ToggleState tg4State = tg4.Current.ToggleState;
					if (tg4State == ToggleState.Off)
						tg4.Toggle();

					logger.LogMessageToFile("<---> Components Sever and shell checkbox state set on  ... ", sLogCount, sLogInterval);
					Thread.Sleep(3000);
				}


				if (step == 6)
				{
					logger.LogMessageToFile("<----->wait extra 15 second  : ", sLogCount, sLogInterval);
					WaitUntilMyButtonFoundInThisWindow(WindowName, "Close", 600);
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
						System.Windows.MessageBox.Show(" <-----> Button not found ;" + buttonName, stepMsg);
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


		
		public void StartTestWorker(string setupPath, string projectPath, string projectName)
		{
			// startup Launcher
			//if (Configuration.Launcher)
			//{
				//MessageBox.Show( setupPath);
				procLauncher = new System.Diagnostics.Process();
				procLauncher.EnableRaisingEvents = false;
				procLauncher.StartInfo.FileName = "Epia.Launcher.exe";
				procLauncher.StartInfo.Arguments = "/objecttype Egemin.EPIA.WCS.Core.Project  /uri gtcp://localhost:50000/Project /Startup Overview";
				procLauncher.StartInfo.WorkingDirectory = setupPath;
				procLauncher.Start();
			//}


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

				procExplorer = new System.Diagnostics.Process();
				procExplorer.EnableRaisingEvents = false;
				procExplorer.StartInfo.FileName = "Epia.Explorer.exe";
				string explorerInput = "";
				string path = '"'+xmlfile+'"';
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
				procExplorer.StartInfo.WorkingDirectory = setupPath;
				//procExplorer.StartInfo.WorkingDirectory = @"C:\Program Files\Egemin\Epia 1.9.12";

				logger.LogMessageToFile(" explorer args :" + explorerInput, 0, 0);
				logger.LogMessageToFile(" explorer path :" + setupPath, 0, 0);
			   
				procExplorer.Start();
			//}
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
	   
		private void WriteInstallationToReg(string AppName, string FilePath, string MsiName)
		{
			logger.LogMessageToFile("------------WriteInstallationToReg: "+AppName+"InstallationPath"
				+ System.Environment.NewLine + "FilePath: " + FilePath + " & " + "MsiName: " + MsiName, 0, 0);
			RegistryKey key = Registry.CurrentUser.CreateSubKey(REGKEY);
			key.SetValue(AppName+"InstallationPath", FilePath);
			key.SetValue(AppName+"InstallationName", MsiName);
		}
		
		private void AutoDeploymentOutputLog(string InstalledApp, string appDeployedPath)
		{
			// After Epia installed, write the log record to output file: Format
			//2007-4-27 17:47:36: , Installed, C:\Program Files\Egemin\Epia 1.9.10, msi path info, rootpath                                                                            // 0 Time 
			string msg = ", Installed:" + InstalledApp                              // 1 Installed App
				+ " , " + appDeployedPath                                           // 2 Deployed Path
				+ " , " + m_testDef                                                 // 3 Definition: CI, Nightly... 
				+ " , " + m_testApp                                                 // 4 APP
				+ " , " + TESTTOOL_VERSION                                          // 5 TestTool Version
				+ " , " + m_Settings.SelectedProjectFile.Substring(0,
							m_Settings.SelectedProjectFile.LastIndexOf("."))        // 6 Default Project File
				+ " , " + m_installScriptDir                                        // 7 Build install path
				+ " , " + m_SystemDrive                                            // 8 Build base path
				+ " , " + sDemonstration.ToString().ToLower()                       // 9 demo          
				+ " , " + m_Settings.Mail.ToString().ToLower()                      // 10 send mail
				+ " , " + m_Settings.ExcelVisible                                   // 11 Excel Visible
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

				string DotnetVersionPath = m_SystemDrive + @"WINDOWS\Microsoft.NET\Framework\v3.5";
				string exePath = Path.Combine(DotnetVersionPath, "csc.exe");
				// Run recompile Process
				output = Utilities.RunProcessAndGetOutput(exePath, arg);
				if (output.IndexOf("error") >= 0)
					throw new Exception(output);

				Log("TestRun Recompiled ");

				Thread.Sleep(2000);
				#endregion
			}
			catch (Exception exRecomp)
			{
				string msg = exRecomp.ToString() + System.Environment.NewLine + exRecomp.StackTrace;
				System.Windows.MessageBox.Show(msg, "RecompileTestRuns");
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
                System.Windows.MessageBox.Show("InstallEtriccUISetup exception : " +ex.Message + " -- " +ex.StackTrace, "InstallEtriccUISetup(1) with filepath:"+FilePath);
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

				AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 5);
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
								System.Windows.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
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

					aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 10);

					#endregion
				}

				if (aeNextButton == null)
					System.Windows.MessageBox.Show("next button not found", "After remove, Reinstall App");
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
						InstallEtriccSetupByStep(SetupWindowName, SetupStepDescriptions[i], i);
					}

					installed = true;
				}
				#endregion
			}
			catch (Exception ex)
			{
				string msg = ex.ToString() + System.Environment.NewLine + ex.StackTrace;
				System.Windows.MessageBox.Show(msg, "InstallSetup");
				throw new Exception(msg + "  during <InstallSetup>");
				// WIP log exception and return 
			}

			//Log("End Install Setup " + msiName + " at " + FilePath);


			if (installed)
			{
				//save the name and the path in the registry to remove the setup
				WriteInstallationToReg(Constants.ETRICCUI, FilePath, msiName);

				// save deployment log
				AutoDeploymentOutputLog(m_testApp, m_SystemDrive + @"Program Files\Egemin\Epia " + mTestedVersion);
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

		private bool InstallEpiaSetup(string FilePath, string TestedVersion)
		{
			//MessageBox.Show("InstallEpiaSetup FilePath: " + FilePath);
			logger.LogMessageToFile("::: InstallEpiaSetup : " + FilePath, sLogCount, sLogInterval);

			//find the msi in the filepath
			string msiName = string.Empty;
			DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
            FileInfo[] files = DirInfo.GetFiles("*" + sEpia4InstallerName);
			bool installed = false;

			TestedVersion = string.Empty;
			try
			{
				if (files[0] != null)
					msiName = files[0].Name;
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show("cscript exception: " + ex.Message + " -- " + ex.StackTrace, "InstallSetup");
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
                        logger.LogMessageToFile( "New aeForm name is: " + aeForm.Current.Name, sLogCount, sLogInterval);
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

               

				AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 5);
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
								System.Windows.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
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

					aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 10);

					#endregion
				}

				if (aeNextButton == null)
					System.Windows.MessageBox.Show("next button not found", "After remove, Reinstall App");
				else
				{
					System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
					Thread.Sleep(2000);
					Input.MoveTo(OptionPt);
					Thread.Sleep(2000);
					logger.LogMessageToFile("<--->Next Button clicking ... ", sLogCount, sLogInterval);
					Input.ClickAtPoint(OptionPt);
					Thread.Sleep(2000);

					for (int i = 2; i < 9; i++)
					{
						Thread.Sleep(4000);
						logger.LogMessageToFile(" (" + i + ") " + SetupStepDescriptions[i], sLogCount, sLogInterval);
						InstallEpiaSetupByStep(SetupWindowName, SetupStepDescriptions[i], i);
					}

					installed = true;
				}
				#endregion 
                
                #region Install Epia  New Version
                /* 
                Console.WriteLine("Searching for main installer window");
                System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
                AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                AutomationElement btnNext = null;
 
                DateTime startTime = DateTime.Now;
                TimeSpan xTime = DateTime.Now - startTime;
                while (appElement == null && xTime.TotalMilliseconds < 60000)
                {
                    Thread.Sleep(2000);
                    appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                    xTime = DateTime.Now - startTime;
                    if (xTime.TotalMilliseconds > 60000)
                    {
                        System.Windows.Forms.MessageBox.Show("After one minute no Installer Window Form found");
                        return false;
                    }
                }

                if (appElement != null)
                {   // (1) Welcom Main window
                    Console.WriteLine("Welcom Main window opend ");
                    Console.WriteLine("Searching next button...");
                    btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                    if (btnNext != null)
                    {   // (2) Components
                        AUIUtilities.ClickElement(btnNext);
                        appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                        Console.WriteLine("Componts window opend");
                        Console.WriteLine("Searching checkbox...");
                        AutomationElement epiaServerCheckbox = AUIUtilities.GetElementByNameProperty(appElement, "E'pia Server");
                        if (epiaServerCheckbox != null)
                            AUIUtilities.ClickElement(epiaServerCheckbox);

                        AutomationElement epiaShellCheckbox = AUIUtilities.GetElementByNameProperty(appElement, "E'pia Shell");
                        if (epiaShellCheckbox != null)
                            AUIUtilities.ClickElement(epiaShellCheckbox);

                        Thread.Sleep(2000);
                        Console.WriteLine("Searching next button...");
                        btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                        if (btnNext != null)
                        {
                            AUIUtilities.ClickElement(btnNext);

                            appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                            if (appElement != null)
                                Console.WriteLine("Installation Folders window is opend");
                            // in the future maybe will edit installation Folder
                            //WaitUntilInstallationComplete(appElement);
                            // wait until isContent close button found
                            Console.WriteLine("wait until Content Close button found");
                            System.Windows.Automation.Condition c2 = new AndCondition(
                                    new PropertyCondition(AutomationElement.NameProperty, "Close"),
                                    new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                                   );

                            AutomationElement aeBtnClose = null;
                            while (aeBtnClose == null)
                            {
                                Console.WriteLine("Wait until Close button found...");
                                appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                                if (btnNext != null)
                                {
                                    Console.WriteLine("Next > button found first --> Click Next > button");
                                    if (btnNext.Current.IsKeyboardFocusable)
                                        AUIUtilities.ClickElement(btnNext);
                                    else
                                        Console.WriteLine("Next > button IsKeyboardFocusable --> false");
                                }
                                else
                                {
                                    aeBtnClose = appElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                                }
                                Thread.Sleep(5000);
                            }
                            Console.WriteLine("Close button found... ---> Close Installer Window");
                            AUIUtilities.ClickElement(aeBtnClose);
                            Console.WriteLine("---------- Epia Install Successful ---------");
                        }
                    }
                }
                 */
                #endregion
                
			}
			catch (Exception ex)
			{
				string msg = ex.ToString() +"----"+ System.Environment.NewLine + ex.StackTrace;
				System.Windows.MessageBox.Show(msg, "InstallSetup");
				throw new Exception(msg + "  during <InstallSetup>");
				// WIP log exception and return 
			}

			//Log("End Install Setup " + msiName + " at " + FilePath);


			if (installed)
			{
				//save the name and the path in the registry to remove the setup
				WriteInstallationToReg(Constants.EPIA, FilePath, msiName);

				// save deployment log  
				AutoDeploymentOutputLog(m_testApp, m_SystemDrive + @"Program Files\Egemin\Epia Server");
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

				logger.LogMessageToFile(" **************** ( END "+ "*" + sEpia4InstallerName + " Deployment )************************** ", 0, 0);
			}

			return installed;

		}

		private bool InstallEtricc5Setup(string FilePath, string TestedVersion)
		{
			logger.LogMessageToFile("::: InstallEtricc5Setup : " + FilePath, sLogCount, sLogInterval);

			//find the msi in the filepath
			string msiName = string.Empty;
			DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
			FileInfo[] files = DirInfo.GetFiles("Etricc 5.msi");
			bool installed = false;

			TestedVersion = string.Empty;
			try
			{
				if (files[0] != null)
					msiName = files[0].Name;
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show("cscript exception: " + ex.Message + " -- " + ex.StackTrace, "InstallSetup");
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

				AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 5);
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
								System.Windows.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
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

					aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 10);
					
					#endregion
				}

				if (aeNextButton == null)
					System.Windows.MessageBox.Show("next button not found", "After remove, Reinstall App");
				else
				{
					System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeNextButton);
					Thread.Sleep(2000);
					Input.MoveTo(OptionPt);
					Thread.Sleep(2000);
					logger.LogMessageToFile("<--->Next Button clicking ... ", sLogCount, sLogInterval);
					Input.ClickAtPoint(OptionPt);
					Thread.Sleep(2000);

					for (int i = 1; i < 11; i++)
					{
						Thread.Sleep(4000);
						logger.LogMessageToFile(" (" + i + ") " + SetupStepDescriptions[i], sLogCount, sLogInterval);
						InstallSetupByStep(SetupWindowName, SetupStepDescriptions[i], i);
					}

					installed = true;
				}
				#endregion
			}
			catch (Exception ex)
			{
				string msg = ex.ToString() + System.Environment.NewLine + ex.StackTrace;
				System.Windows.MessageBox.Show(msg, "InstallSetup ");
				throw new Exception(msg + "  during <InstallSetup>");
				// WIP log exception and return 
			}

			//Log("End Install Setup " + msiName + " at " + FilePath);

		   
			if (installed)
			{
				//save the name and the path in the registry to remove the setup
				WriteInstallationToReg(Constants.ETRICC5, FilePath, msiName);

				// save deployment log
				AutoDeploymentOutputLog(m_testApp, m_SystemDrive + @"Program Files\Egemin\Epia " + mTestedVersion);
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

				logger.LogMessageToFile(" **************** ( END Etricc 5.msi Deployment )************************** ", 0, 0);
			}

			return installed;
		   
		}

		private bool InstallKCSetup(string FilePath, string TestedVersion)
		{
			logger.LogMessageToFile("::: InstallKCSetup : " + FilePath, sLogCount, sLogInterval);

			//find the msi in the filepath
			string msiName = string.Empty;
			DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
			FileInfo[] files = DirInfo.GetFiles("Ewcs KimberlyClark.msi");
			bool installed = false;

			TestedVersion = string.Empty;
			try
			{
				if (files[0] != null)
					msiName = files[0].Name;
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show("cscript exception: " + ex.Message + " -- " + ex.StackTrace, "InstallSetup");
			}

			try
			{
				#region Install KC

				string[] SetupStepDescriptions = new string[100];
				SetupStepDescriptions[0] = "Welcome";
				SetupStepDescriptions[1] = "Welcome to the E'wcs *";
				SetupStepDescriptions[2] = "Components";
				SetupStepDescriptions[3] = "Installation Folders";
				SetupStepDescriptions[4] = "Confirm Installation";
				SetupStepDescriptions[5] = "Installing E'wcs *";
				SetupStepDescriptions[6] = "Installation Complete";

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

				AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 5);
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
								System.Windows.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
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

					// install KC   UIAutomation
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
					System.Windows.MessageBox.Show("next button not found", "After remove, Reinstall App");
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
						InstallKCSetupByStep(SetupWindowName, SetupStepDescriptions[i], i);
					}

					installed = true;
				}
				#endregion
			}
			catch (Exception ex)
			{
				string msg = ex.ToString() + System.Environment.NewLine + ex.StackTrace;
				System.Windows.MessageBox.Show(msg, "InstallSetup");
				throw new Exception(msg + "  during <InstallSetup>");
				// WIP log exception and return 
			}

			if (installed)
			{
				//save the name and the path in the registry to remove the setup
				WriteInstallationToReg(Constants.KC, FilePath, msiName);

				// save deployment log
				AutoDeploymentOutputLog(m_testApp, m_SystemDrive + @"Program Files\Egemin\Ewcs Server");
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

				logger.LogMessageToFile(" **************** ( END Ewcs KimberlyClark.msi Deployment )************************** ", 0, 0);
			}

			return installed;

		}

		private bool InstallKCTestProgramSetup(string FilePath, string TestedVersion)
		{
			logger.LogMessageToFile("::: InstallKCTestProgramSetup : " + FilePath, sLogCount, sLogInterval);

			//find the msi in the filepath
			string msiName = string.Empty;
			DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
			FileInfo[] files = DirInfo.GetFiles("Ewcs KimberlyClark TestPrograms.msi");
			bool installed = false;

			TestedVersion = string.Empty;
			try
			{
				if (files[0] != null)
					msiName = files[0].Name;
			}
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show("cscript exception: " + ex.Message + " -- " + ex.StackTrace, "InstallSetup");
			}

			try
			{
				#region Install KC

				string[] SetupStepDescriptions = new string[100];
				SetupStepDescriptions[0] = "Welcome";
				SetupStepDescriptions[1] = "Welcome to the E'wcs *";
				SetupStepDescriptions[2] = "Components";
				SetupStepDescriptions[3] = "Installation Folders";
				SetupStepDescriptions[4] = "Confirm Installation";
				SetupStepDescriptions[5] = "Installing E'wcs *";
				SetupStepDescriptions[6] = "Installation Complete";

				string SetupWindowName = string.Empty;
				string sErrorMessage = string.Empty;
				AutomationElement aeForm = null;

				// install KC Test Program   UIAutomation
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

				AutomationElement aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 5);
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
								System.Windows.MessageBox.Show(" <-----> Window not found ", SetupStepDescriptions[0]);
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

					aeNextButton = WaitUntilMyButtonFoundInThisWindow(SetupWindowName, "Next >", 10);

					#endregion
				}

				if (aeNextButton == null)
					System.Windows.MessageBox.Show("next button not found", "After remove, Reinstall App");
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
						InstallKCTestProgramSetupByStep(SetupWindowName, SetupStepDescriptions[i], i);  
					}

					installed = true;
				}
				#endregion
			}
			catch (Exception ex)
			{
				string msg = ex.ToString() + System.Environment.NewLine + ex.StackTrace;
				System.Windows.MessageBox.Show(msg, "InstallSetup");
				throw new Exception(msg + "  during <InstallSetup>");
				// WIP log exception and return 
			}

			if (installed)
			{
				//save the name and the path in the registry to remove the setup
				WriteInstallationToReg(Constants.KC+"TestProgram", FilePath, msiName);

				logger.LogMessageToFile(" **************** ( END Ewcs KimberlyClark TestPrograms.msi Deployment )************************** ", 0, 0);
			}

			return installed;

		}

		private void RemoveSetup(string AppName)
		{
			logger.LogMessageToFile("------------Start Removed Setup AppName: "+AppName, 0, 0);

			string InstallPath = string.Empty;
			string InstallName = string.Empty;
			//find installed msi in reg
			try
			{
				// private const string REGKEY = "Software\\Egemin\\Automatic testing\\";
				RegistryKey key = Registry.CurrentUser.OpenSubKey(REGKEY);
				object keyvalue;

				keyvalue = key.GetValue(AppName+"InstallationPath");
				if (keyvalue != null)
					InstallPath = keyvalue.ToString();

				keyvalue = key.GetValue(AppName+"InstallationName");
				if (keyvalue != null)
					InstallName = keyvalue.ToString();

				logger.LogMessageToFile("(1)Removed Setup:"+AppName+" -- InstallPath " + InstallPath + " and InstallName " + InstallName, 0, 0);

			}
            catch (System.NullReferenceException ex1)
            {
                //MessageBox.Show(ex.ToString() + System.Environment.NewLine + ex.StackTrace,
                //    "DeployTestLogic.Tester  RemoveSetup: find register Key");
                logger.LogMessageToFile(AppName+ " REGKEY not exist:"+ex1.ToString() + System.Environment.NewLine + ex1.StackTrace, 0, 0);

            }
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString() + System.Environment.NewLine + ex.StackTrace,
				    "DeployTestLogic.Tester  RemoveSetup: find register Key");
				logger.LogMessageToFile(ex.ToString() + System.Environment.NewLine + ex.StackTrace, 0, 0);

			}

			if ((InstallPath == string.Empty) || (InstallName == string.Empty))
				return;

			string unattendedXmlFilePath = m_CurrentDrive+@"Epia 3\Testing\Automatic\AutomaticTests\TestData\Etricc5";
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
			Log("Removed Setup: "+AppName+" -- " + InstallName + " at " + InstallPath);
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
				logger.LogMessageToFile(AppName+"------ Test Exception : " + ex.Message + "\r\n" + ex.StackTrace, 0, 0);
			}

			//Remove the install information from the registry
            string RegName = AppName;
            if ( AppName.EndsWith(TestTools.ConstCommon.ETRICC_UI) )
                RegName = Constants.ETRICCUI;
			WriteInstallationToReg(RegName, string.Empty, string.Empty);
		}
		
		#endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

		#region // —— UI Help Methods •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
		public void ClickUiScreenActionToAvoidScreenStandBy()
		{
			System.Windows.Point point = new Point(1,1);

			Input.MoveToAndRightClick(point);

			Thread.Sleep(2000);

			point.Y = point.Y + 300;
			Input.MoveToAndClick(point);

			Thread.Sleep(2000);        
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
                    System.Windows.MessageBox.Show("find button in " + ButtonName + " <-----> Window not found: windows name is: " + WindowName, "WaitUntilMyButtonFoundInThisWindow");
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
                    System.Windows.MessageBox.Show("find button in " + ButtonName + " <-----> Window not found: windows name is: " + WindowName, "WaitUntilMyButtonFoundInThisWindowWithStatusEnable");
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

		public AutomationElement RemovePreviousSetupandInstallCurrentEpia(AutomationElement aeWindow)
		{
			AutomationElement aeMyForm = null;
			AutomationElement aeOKButton = null;
			int OKcnt = 0;
			while (aeOKButton == null && OKcnt <= 100)
			{
				aeOKButton = FindSpecificButtonByName(aeWindow, "OK");
				OKcnt++;
				logger.LogMessageToFile(OKcnt + " time search <-----> OK button not found and try again: ", sLogCount, sLogInterval);
			}

			if (aeOKButton == null)
			{
				System.Windows.MessageBox.Show(" <-----> OK button not found ", "RemovePreviousSetupAndInstallCurrent");
			}
			else
			{
				System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeOKButton);
				Thread.Sleep(1000);
				Input.MoveTo(OptionPt);
				Thread.Sleep(1000);
				Input.ClickAtPoint(OptionPt);
				Thread.Sleep(5000);
			}

			// remove previous setup 
			string PreviousInstallPath = m_SystemDrive + @"Program Files\Egemin\AutomaticTesting\Setup\Epia\Previous";
			string unattendedXmlFilePath = m_CurrentDrive + @"Epia 3\Testing\Automatic\AutomaticTests\TestData";
			string NewInstallPath = m_SystemDrive + @"Program Files\Egemin\AutomaticTesting\Setup\Epia\Current";
			if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))  //  Deployed Testing
			{
				PreviousInstallPath = mPreviousSetupPathEpia;
				unattendedXmlFilePath = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY;
				NewInstallPath = mCurrentSetupPathEpia;
			}

			//remove the setup in silent mode
			//string xmlFile = System.IO.Path.Combine(Application.StartupPath, "UnattendedXmlFile.xml");
			string xmlFile = System.IO.Path.Combine(unattendedXmlFilePath, "UnattendedXmlFile.xml");
			//MessageBox.Show(xmlFile, "1");

            string uninstallParm = Path.Combine(PreviousInstallPath, sEpia4InstallerName);
			string args = "/passive /x " + '"' + uninstallParm + '"' + " " + "UnattendedXmlFile=" + '"' + xmlFile + '"';

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
			proc.StartInfo.WorkingDirectory = PreviousInstallPath;
			proc.Start();
			proc.WaitForExit();

			Thread.Sleep(10000);


			Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						  AutomationElement.RootElement,
						 UIAShellEventHandler);

			logger.LogMessageToFile("<-----> Install Current Epia: ", sLogCount, sLogInterval);

			// install Epia UIAutomation
			Utilities.CloseProcess("msiexec");
			Thread.Sleep(3000);
            System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(NewInstallPath, sEpia4InstallerName, string.Empty);
			Thread.Sleep(2000);

			aeMyForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
			Thread.Sleep(2000);
			if (aeMyForm == null)
			{
				logger.LogMessageToFile("aeMyForm  not found : ", sLogCount, sLogInterval);
				System.Windows.MessageBox.Show("aeMyForm  not found : ");
				return null;
			}
			else
				return aeMyForm;

		}

		public AutomationElement RemovePreviousSetupandInstallCurrentEtricc(AutomationElement aeWindow)
		{
			AutomationElement aeMyForm = null;
			AutomationElement aeOKButton = null;
			int OKcnt = 0;
			while (aeOKButton == null && OKcnt <= 100)
			{
				aeOKButton = FindSpecificButtonByName(aeWindow, "OK");
				OKcnt++;
				logger.LogMessageToFile(OKcnt + " time search <-----> OK button not found and try again: ", sLogCount, sLogInterval);
			}

			if (aeOKButton == null)
			{
				System.Windows.MessageBox.Show(" <-----> OK button not found ", "RemovePreviousSetupAndInstallCurrent");
			}
			else
			{
				System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeOKButton);
				Thread.Sleep(1000);
				Input.MoveTo(OptionPt);
				Thread.Sleep(1000);
				Input.ClickAtPoint(OptionPt);
				Thread.Sleep(5000);
			}

			// remove previous setup 
			string PreviousInstallPath = m_SystemDrive + @"Program Files\Egemin\AutomaticTesting\Setup\Epia\Previous";
			string unattendedXmlFilePath = m_CurrentDrive+@"Epia 3\Testing\Automatic\AutomaticTests\TestData";
			string NewInstallPath = m_SystemDrive + @"Program Files\Egemin\AutomaticTesting\Setup\Epia\Current";
			if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))  //  Deployed Testing
			{
				PreviousInstallPath = mPreviousSetupPathEtricc;
				unattendedXmlFilePath = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY;
				NewInstallPath = mCurrentSetupPathEtricc;
			}

			//remove the setup in silent mode
			//string xmlFile = System.IO.Path.Combine(Application.StartupPath, "UnattendedXmlFile.xml");
			string xmlFile = System.IO.Path.Combine(unattendedXmlFilePath, "UnattendedXmlFile.xml");
			//MessageBox.Show(xmlFile, "1");

            string uninstallParm = Path.Combine(PreviousInstallPath, sEpia4InstallerName);
			string args = "/passive /x " + '"' + uninstallParm + '"' + " " + "UnattendedXmlFile=" + '"' + xmlFile + '"';

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
			proc.StartInfo.WorkingDirectory = PreviousInstallPath;
			proc.Start();
			proc.WaitForExit();

			Thread.Sleep(10000);


			Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						  AutomationElement.RootElement,
						 UIAShellEventHandler);

			logger.LogMessageToFile("<-----> Install Current Etricc: ", sLogCount, sLogInterval);

			// install Epia UIAutomation
			Utilities.CloseProcess("msiexec");
			Thread.Sleep(3000);
			System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(NewInstallPath, "Etricc Shell.msi", string.Empty);
			Thread.Sleep(2000);

			aeMyForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
			Thread.Sleep(2000);
			if (aeMyForm == null)
			{
				logger.LogMessageToFile("aeMyForm  not found : ", sLogCount, sLogInterval);
				System.Windows.MessageBox.Show("aeMyForm  not found : ");
				return null;
			}
			else
				return aeMyForm;

		}

		public AutomationElement RemovePreviousSetupandInstallCurrentEtricc5(AutomationElement aeWindow)
		{
			AutomationElement aeMyForm = null;
			AutomationElement aeOKButton = null;
			int OKcnt = 0;
			while (aeOKButton == null && OKcnt <= 100)
			{
				aeOKButton = FindSpecificButtonByName(aeWindow, "OK");
				OKcnt++;
				logger.LogMessageToFile(OKcnt + " time search <-----> OK button not found and try again: ", sLogCount, sLogInterval);
			}
		   
			if ( aeOKButton == null)
			{
				System.Windows.MessageBox.Show(" <-----> OK button not found ", "RemovePreviousSetupAndInstallCurrent");
			}
			else
			{
				logger.LogMessageToFile("OK button found and Click OK Button: ", sLogCount, sLogInterval);
		  
				System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeOKButton);
				Thread.Sleep(1000);
				Input.MoveTo(OptionPt);
				Thread.Sleep(1000);
				Input.ClickAtPoint(OptionPt);
				Thread.Sleep(5000);
			}

			// remove previous setup 
			string PreviousInstallPath = m_SystemDrive+@"Program Files\Egemin\AutomaticTesting\Setup\Etricc5\Previous";
			string unattendedXmlFilePath = m_CurrentDrive + @"Epia 3\Testing\Automatic\AutomaticTests\TestData";
			string NewInstallPath = m_SystemDrive + @"Program Files\Egemin\AutomaticTesting\Setup\Etricc5\Current";
			if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))  //  Deployed Testing
			{
				PreviousInstallPath = mPreviousSetupPathEtricc5;
				unattendedXmlFilePath = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY;
				NewInstallPath = mCurrentSetupPathEtricc5;
			}

			logger.LogMessageToFile("PreviousInstallPath: " + PreviousInstallPath, sLogCount, sLogInterval);
			logger.LogMessageToFile("NewInstallPath: " + NewInstallPath, sLogCount, sLogInterval);
		  
			//remove the setup in silent mode
			//string xmlFile = System.IO.Path.Combine(Application.StartupPath, "UnattendedXmlFile.xml");
			string xmlFile = System.IO.Path.Combine(unattendedXmlFilePath, "UnattendedXmlFile.xml");
			//MessageBox.Show(xmlFile, "1");

			logger.LogMessageToFile("xmlFile: " + xmlFile, sLogCount, sLogInterval);
			
			string uninstallParm = Path.Combine(PreviousInstallPath, "Etricc 5.msi");
			string args = "/passive /x " + '"' + uninstallParm + '"' + " " + "UnattendedXmlFile=" + '"' + xmlFile + '"';
			logger.LogMessageToFile("args: " + args, sLogCount, sLogInterval);

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
			proc.StartInfo.WorkingDirectory = PreviousInstallPath;
			proc.Start();
			proc.WaitForExit();
		   
			Thread.Sleep(10000);


			Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						  AutomationElement.RootElement,
						 UIAShellEventHandler);

			logger.LogMessageToFile("<-----> Install Current Etricc 5: ", sLogCount, sLogInterval);
		
			// install Etricc 5   UIAutomation
			 Utilities.CloseProcess("msiexec");
			Thread.Sleep(3000);
			System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(NewInstallPath, "Etricc 5.msi", string.Empty);
			Thread.Sleep(2000);

			aeMyForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
			Thread.Sleep(2000);
			if (aeMyForm == null)
			{
				logger.LogMessageToFile("aeMyForm  not found : ", sLogCount, sLogInterval);
				System.Windows.MessageBox.Show("aeMyForm  not found : ");
				return null;
			}
			else
				return aeMyForm;

		}

		public AutomationElement RemovePreviousSetupandInstallCurrentKC(AutomationElement aeWindow)
		{
			AutomationElement aeMyForm = null;
			AutomationElement aeOKButton = null;
			int OKcnt = 0;
			while (aeOKButton == null && OKcnt <= 100)
			{
				aeOKButton = FindSpecificButtonByName(aeWindow, "OK");
				OKcnt++;
				logger.LogMessageToFile(OKcnt + " time search <-----> OK button not found and try again: ", sLogCount, sLogInterval);
			}

			if (aeOKButton == null)
			{
				System.Windows.MessageBox.Show(" <-----> OK button not found ", "RemovePreviousSetupAndInstallCurrent");
			}
			else
			{
				logger.LogMessageToFile("OK button found and Click OK Button: ", sLogCount, sLogInterval);

				System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeOKButton);
				Thread.Sleep(1000);
				Input.MoveTo(OptionPt);
				Thread.Sleep(1000);
				Input.ClickAtPoint(OptionPt);
				Thread.Sleep(5000);
			}

			// remove previous setup 
			string PreviousInstallPath = m_SystemDrive + @"Program Files\Egemin\AutomaticTesting\Setup\Etricc5\Previous";
			string unattendedXmlFilePath = m_CurrentDrive + @"Epia 3\Testing\Automatic\AutomaticTests\TestData";
			string NewInstallPath = m_SystemDrive + @"Program Files\Egemin\AutomaticTesting\Setup\Etricc5\Current";
			if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))  //  Deployed Testing
			{
				PreviousInstallPath = mPreviousSetupPathKC;
				unattendedXmlFilePath = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY;
				NewInstallPath = mCurrentSetupPathKC;
			}

			logger.LogMessageToFile("PreviousInstallPath: " + PreviousInstallPath, sLogCount, sLogInterval);
			logger.LogMessageToFile("NewInstallPath: " + NewInstallPath, sLogCount, sLogInterval);

			//remove the setup in silent mode
			//string xmlFile = System.IO.Path.Combine(Application.StartupPath, "UnattendedXmlFile.xml");
			string xmlFile = System.IO.Path.Combine(unattendedXmlFilePath, "UnattendedXmlFile.xml");
			//MessageBox.Show(xmlFile, "1");

			logger.LogMessageToFile("xmlFile: " + xmlFile, sLogCount, sLogInterval);

			string uninstallParm = Path.Combine(PreviousInstallPath, "Ewcs KimberlyClark.msi");
			string args = "/passive /x " + '"' + uninstallParm + '"' + " " + "UnattendedXmlFile=" + '"' + xmlFile + '"';
			logger.LogMessageToFile("args: " + args, sLogCount, sLogInterval);

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
			proc.StartInfo.WorkingDirectory = PreviousInstallPath;
			proc.Start();
			proc.WaitForExit();

			Thread.Sleep(10000);


			Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						  AutomationElement.RootElement,
						 UIAShellEventHandler);

			logger.LogMessageToFile("<-----> Install Current KC: ", sLogCount, sLogInterval);
			logger.LogMessageToFile("<-----> Install Current KC: NewInstallPath: " + NewInstallPath, sLogCount, sLogInterval);


			// install KC   UIAutomation
			Utilities.CloseProcess("msiexec");
			Thread.Sleep(3000);
			System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(NewInstallPath, "Ewcs KimberlyClark.msi", string.Empty);
			Thread.Sleep(2000);

			aeMyForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
			Thread.Sleep(2000);
			if (aeMyForm == null)
			{
				logger.LogMessageToFile("aeMyForm  not found : ", sLogCount, sLogInterval);
				System.Windows.MessageBox.Show("aeMyForm  not found : ");
				return null;
			}
			else
				return aeMyForm;

		}

		public AutomationElement RemovePreviousSetupandInstallCurrentKCTestProgram(AutomationElement aeWindow)
		{
			AutomationElement aeMyForm = null;
			AutomationElement aeOKButton = null;
			int OKcnt = 0;
			while (aeOKButton == null && OKcnt <= 100)
			{
				aeOKButton = FindSpecificButtonByName(aeWindow, "OK");
				OKcnt++;
				logger.LogMessageToFile(OKcnt + " time search <-----> OK button not found and try again: ", sLogCount, sLogInterval);
			}

			if (aeOKButton == null)
			{
				System.Windows.MessageBox.Show(" <-----> OK button not found ", "RemovePreviousSetupAndInstallCurrentKCTestProgram");
			}
			else
			{
				logger.LogMessageToFile("OK button found and Click OK Button: ", sLogCount, sLogInterval);

				System.Windows.Point OptionPt = AUIUtilities.GetElementCenterPoint(aeOKButton);
				Thread.Sleep(1000);
				Input.MoveTo(OptionPt);
				Thread.Sleep(1000);
				Input.ClickAtPoint(OptionPt);
				Thread.Sleep(5000);
			}

			// remove previous setup 
			string PreviousInstallPath = m_SystemDrive + @"Program Files\Egemin\AutomaticTesting\Setup\Etricc5\Previous";
			string unattendedXmlFilePath = m_CurrentDrive + @"Epia 3\Testing\Automatic\AutomaticTests\TestData";
			string NewInstallPath = m_SystemDrive + @"Program Files\Egemin\AutomaticTesting\Setup\Etricc5\Current";
			if (m_TestWorkingDirectory.Equals(m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY))  //  Deployed Testing
			{
				PreviousInstallPath = mPreviousSetupPathKC;
				unattendedXmlFilePath = m_SystemDrive + ConstCommon.DEPLOY_TESTS_DIRECTORY;
				NewInstallPath = mCurrentSetupPathKC;
			}

			logger.LogMessageToFile("PreviousInstallPath: " + PreviousInstallPath, sLogCount, sLogInterval);
			logger.LogMessageToFile("NewInstallPath: " + NewInstallPath, sLogCount, sLogInterval);

			//remove the setup in silent mode
			//string xmlFile = System.IO.Path.Combine(Application.StartupPath, "UnattendedXmlFile.xml");
			string xmlFile = System.IO.Path.Combine(unattendedXmlFilePath, "UnattendedXmlFile.xml");
			//MessageBox.Show(xmlFile, "1");

			logger.LogMessageToFile("xmlFile: " + xmlFile, sLogCount, sLogInterval);

			string uninstallParm = Path.Combine(PreviousInstallPath, "Ewcs KimberlyClarkTestPrograms.msi");
			string args = "/passive /x " + '"' + uninstallParm + '"' + " " + "UnattendedXmlFile=" + '"' + xmlFile + '"';
			logger.LogMessageToFile("args: " + args, sLogCount, sLogInterval);

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
			proc.StartInfo.WorkingDirectory = PreviousInstallPath;
			proc.Start();
			proc.WaitForExit();

			Thread.Sleep(10000);


			Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						  AutomationElement.RootElement,
						 UIAShellEventHandler);

			logger.LogMessageToFile("<-----> Install Current KCTestPrograms: ", sLogCount, sLogInterval);

			// install Etricc 5   UIAutomation
			Utilities.CloseProcess("msiexec");
			Thread.Sleep(3000);
			System.Diagnostics.Process SetupProc = Utilities.StartProcessNoWait(NewInstallPath, "Ewcs KimberlyClarkTestPrograms.msi", string.Empty);
			Thread.Sleep(2000);

			aeMyForm = AutomationElement.FromHandle(SetupProc.MainWindowHandle);
			Thread.Sleep(2000);
			if (aeMyForm == null)
			{
				logger.LogMessageToFile("aeMyForm  not found : ", sLogCount, sLogInterval);
				System.Windows.MessageBox.Show("aeMyForm  not found : ");
				return null;
			}
			else
				return aeMyForm;

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

			if (aeAllButtons.Count == 1 )
			{
				if (aeAllButtons[0].Current.Name.Equals(buttonName))
				{
					status = true;
				}
				
			}
			
			return status;
		}

		#endregion

		#region // —— HELP Methods •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
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
			return result;
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


		/// <summary>
		/// Cancel a drive mapping to the destination
		/// </summary>
		/// <param name="Destination">Full drive path</param>
		public static int Disconnect(string localpath)
		{
			int result = WNetCancelConnection2A(localpath, 1, 1);
			return result;
		}
		#endregion

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
			logger.LogMessageToFile("OnUIAShellEvent  ==== "+str, 0, 0);
			

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
		
		public enum STATE
		{
			UNDEFINED,
			PENDING,
			INPROGRESS,
			EXCEPTION,
		}

	}
}
