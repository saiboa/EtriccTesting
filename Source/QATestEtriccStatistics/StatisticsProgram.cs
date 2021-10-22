using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Xml;
using System.Xml.Linq;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using TestTools;
using Excel = Microsoft.Office.Interop.Excel;
using TFSQATestTools;

namespace QATestEtriccStatistics
{
	class StatisticsProgram
	{
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Fields of Program (38)
		internal static TestTools.Logger logger;

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
        static bool sUseExcel = true;
		static string sParentProgram = string.Empty;
		static string sTestType = "all";
		static string sTestDefinitionFile = string.Empty;
		static string[] mTestDefinitionTypes;
		static string sInfoFileKey = string.Empty;
        static string sNetworkMap = string.Empty;
        static string sCurrentBuildDefinition = string.Empty;
        
		// PCinfo
		static public string PCName;
		static public string OSName;
		static public string OSVersion;
		static public string UICulture;
		static public string TimeOnPC;
		// Build info
		static string sBuildDropFolder = string.Empty;
		static string sBuildConfig = string.Empty;
		static string sBuildDef = string.Empty;
		static string sBuildNr = string.Empty;
		static string sTestToolsVersion = string.Empty;
		static string m_SystemDrive = string.Empty;
		static string sCurrentProject = "Demo";
		static string sTargetPlatform = string.Empty;
		static string sCurrentPlatform = string.Empty;
		static string sTestResultFolder = string.Empty;

        static DateTime sReleaseVersionDate = DateTime.MinValue;
        static DateTime sCurrentVersionDate = DateTime.MinValue;
		// Testcase not used =================================
		public static string sConfigurationName = string.Empty;
		static string sErrorMessage;
        static string sErrorScreenMessage;
		static bool sEventEnd;
		static string sExcelVisible = string.Empty;
		static bool sAutoTest = true;
        static string sInstallMsiDir = @"C:\LocalTest";
		public static string sLayoutName = string.Empty;
		static string sServerRunAs = "Service";
		static bool sDemo;
		static string sSendMail = "false";
		static string sTFSServer = "http://Team2010App.TeamSystems.Egemin.Be:8080";
        static bool sErrorScreen = false;
		// LOG=================================================================
		public static string slogFilePath = @"C:\";
		static string sOutFilename = "OutFilename";
		static string sOutFilePath = string.Empty;
		static StreamWriter Writer;
		// Build param ========================================================
		static IBuildServer m_BuildSvc;
		static bool TFSConnected = true;
        static bool sThisBuildBeforeTestCaseDate = false;
		// excel 	--------------------------------------------------------
		static Excel.Application xApp;
        static int sHeaderContentsLength = 11;
		
		static AutomationEventHandler UIErrorEventHandler = new AutomationEventHandler(OnErrorUIEvent);
        private static bool sEpiaServerStartupOK = true;
        private static bool sSqlOrReportServiceOK = true;
        private static bool sPaeserConfiguratorConnectToComputerOK = true;
		#endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Methods of Program (1)
		/// <summary>
		/// 
		/// </summary>
		/// <param name="args">Application input
		///     </param>
		[STAThread]
		static void Main(string[] args)
		{
			try  // Get test PC info======================================
			{
				TestTools.HelpUtilities.SavePCInfo("y");
				TestTools.HelpUtilities.GetPCInfo(out PCName, out OSName, out OSVersion, out UICulture, out TimeOnPC);
                Console.WriteLine("<PCName : " + PCName + ">, <OSName : " + OSName + ">, <OSVersion : " + OSVersion + ">");
                Console.WriteLine("<TimeOnPC : " + TimeOnPC + ">, <UICulture : " + UICulture + ">");
				string windir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
				m_SystemDrive = Path.GetPathRoot(windir);

			}
			catch (Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Check PC-info");
			}

			sOnlyUITest = false;
			string x = System.Configuration.ConfigurationManager.AppSettings.Get("OnlyUITest");
			if (x.ToLower().StartsWith("true"))
				sOnlyUITest = true;

			Console.WriteLine("sOnlyUITest : " + sOnlyUITest);


            sUseExcel = true;
            string y = System.Configuration.ConfigurationManager.AppSettings.Get("useExcel");
            if (y.ToLower().StartsWith("true"))
                sUseExcel = true;
            else
                sUseExcel = false;

            Console.WriteLine("useExcel : " + sUseExcel);

			sCurrentProject = System.Configuration.ConfigurationManager.AppSettings.Get("CurrentProject");
			Console.WriteLine("sCurrentProject : " + sCurrentProject);

			if (!sOnlyUITest)
			{
				try
				{
					// validate inputs
					if (args != null)
					{
						for (int i = 0; i <= 19; i++)
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
                        sCurrentBuildDefinition = args[19];

						sTestResultFolder = sBuildDropFolder + "\\TestResults";
						if (!System.IO.Directory.Exists(sTestResultFolder))
							System.IO.Directory.CreateDirectory(sTestResultFolder);

						//Epia3Common.CreateOutputFileInfo(args, PCName, ref sOutFilePath, ref sOutFilename);
						//CreateOutputFileInfo(args, sCurrentPlatform, PCName, ref sOutFilePath, ref sOutFilename);
                        sOutFilename = FileManipulation.CreateOutputInfoFileName(sInfoFileKey, sAutoTest);
						sOutFilePath = Path.GetFullPath(sBuildDropFolder) + "\\TestResults";
						Console.WriteLine("sOutFilePath : " + sOutFilePath);
						Console.WriteLine("sOutFilePath2 : " + Path.GetDirectoryName(sOutFilePath));
						Console.WriteLine("sOutFilename : " + sOutFilename);

						Epia3Common.CreateTestLog(ref slogFilePath, sOutFilePath, sOutFilename, ref Writer);
						logger = new Logger(slogFilePath);     // logger passed to other object

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
						Epia3Common.WriteTestLogMsg(slogFilePath, "16) TestDefinitionFile: " + sTestDefinitionFile, sOnlyUITest);
						Epia3Common.WriteTestLogMsg(slogFilePath, "17) InfoFileKey: " + sInfoFileKey, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "18) NetworkMap: " + sNetworkMap, sOnlyUITest);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "19) Current Build Definition: " + sCurrentBuildDefinition, sOnlyUITest);

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
                        Console.WriteLine("19) Current Build Definition: " + sCurrentBuildDefinition);

						mTestDefinitionTypes = System.IO.File.ReadAllLines(sTestDefinitionFile);

						for (int i = 0; i < mTestDefinitionTypes.Length; i++)
						{
							Console.WriteLine(i + " testdefinition : " + mTestDefinitionTypes[i]);
						}

                        if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc Stat Prog.Production.Release", new DateTime(2012, 12, 19), ref sErrorMessage) == true)
                            sThisBuildBeforeTestCaseDate = true;

                        Epia3Common.WriteTestLogMsg(slogFilePath, "sThisBuildBeforeTestCaseDate: " + sThisBuildBeforeTestCaseDate, sOnlyUITest);

                        sCurrentVersionDate = DeployUtilities.getCurrentBuildCompleteDate(sBuildNr, ref sReleaseVersionDate, "Etricc 5", sCurrentBuildDefinition, ref sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "20) sCurrentVersionDate: " + sCurrentVersionDate.ToShortDateString(), sOnlyUITest);

                        if (sBuildNr.IndexOf("Hotfix") >= 0) //string buildnr = "Hotfix 4.3.2.1 of Epia.Production.Hotfix_20120405.1";
                        {
                            //Release 4.4.4 of Epia.Production.Release_20120731.1
                            string releaseBuildNr = DeployUtilities.GetReleaseBuildNrFromThisHotfixBuild(sBuildNr);
                            sReleaseVersionDate = DeployUtilities.GetDateCompletedOfThisBuild("Etricc 5", "Etricc Stat Prog.Production.Release", releaseBuildNr);
                        }
                        else if (sBuildNr.IndexOf("Production.Release") >= 0)
                        {
                            sReleaseVersionDate = DeployUtilities.GetDateCompletedOfThisBuild("Etricc 5", "Etricc Stat Prog.Production.Release", sBuildNr);
                        }
                        else
                            sReleaseVersionDate = DateTime.MinValue;

                        Epia3Common.WriteTestLogMsg(slogFilePath, "21) sReleaseVersionDate: " + sReleaseVersionDate.ToShortDateString(), sOnlyUITest);
					}
				}
				catch (Exception ex)
				{
					System.Windows.Forms.MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Validate command-line params");
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
                        Uri serverUri = new Uri(serverUrl);
                        System.Net.ICredentials tfsCredentials
                            = new System.Net.NetworkCredential("TfsBuild", "Egemin01", "TeamSystems.Egemin.Be");

                        TfsTeamProjectCollection tfsProjectCollection
                            = new TfsTeamProjectCollection(serverUri, tfsCredentials);
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
                                TestTools.MessageBoxEx.Show("Team Foundation services are not available from server\nWill try to reconnect the Server ...\nException message:"+ex.Message,
                                 kTime++ + " During E'tricc Statistics UI Testing, please not touch the screen, time:" + DateTime.Now.ToLongTimeString(), (uint)Tfs.ReconnectDelay);
                                System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                                conn = false;
                            }
                            catch (Exception ex)
                            {
                                TestTools.MessageBoxEx.Show("TeamFoundation getService Exception:" + ex.Message + " ----- " + ex.StackTrace,
                                     kTime++ + " This is automatic testing, please not touch the screen: exception, time:" + DateTime.Now.ToLongTimeString(), (uint)Tfs.ReconnectDelay);
                                System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                                conn = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Get TFS Server:" + ex.Message, sOnlyUITest);
                        TFSConnected = false;
                    }
                }
                else
                {
                    TFSConnected = false;
                }
            }
            else
            {
                sBuildNr = "Etricc Stat Prog.Main.CI_20130205.2";
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc Stat Prog.Production.Release", new DateTime(2012, 12, 19), ref sErrorMessage) == true)
                    sThisBuildBeforeTestCaseDate = true;

                Console.WriteLine("sThisBuildBeforeTestCaseDate  : " + sThisBuildBeforeTestCaseDate);
                //Console.WriteLine("  sBuildNr:" + sBuildNr);
                //sThisBuildBeforeTestCaseDate = true;
            }

			Console.WriteLine("Test started:");
			Epia3Common.WriteTestLogMsg(slogFilePath, "Test started: ", sOnlyUITest);
			sTestCaseName[0] = UNINSTALL_ETRICC5_IF_ALREADY_INSTALLED;
			sTestCaseName[1] = INSTALL_ETRICC5;
			sTestCaseName[2] = PREPARE_TESTDATA;
			sTestCaseName[3] = PARSERCONFIGURATOR_CONNECT_COMPUTER;
			sTestCaseName[4] = CREATE_XML_SCHEMA_DEFINITION;
			sTestCaseName[5] = CREATE_NEW_PARSE_PROJECT;
			sTestCaseName[6] = CREATE_DATABASE;
			sTestCaseName[7] = SET_DATABASE;
			sTestCaseName[8] = CREATE_ETRICC_LAYOUT;
			sTestCaseName[9] = PARSE_TEST_PROJECT;
			sTestCaseName[10] = MODIFY_CONFIGURATION_FILES;
			sTestCaseName[11] = START_EPIA_SERVER_SHELL;
			sTestCaseName[12] = ReportName.StatusGraphicalView;
			//====ANALYSIS=========================================//
			sTestCaseName[13] = ReportName.ANALYSIS_TransportLookupBySrcDstGroup;
            sTestCaseName[14] = ReportName.ANALYSIS_ProjectActivation;
			sTestCaseName[15] = ReportName.ANALYSIS_TransportLookupBySrcDstLocationOrStation;
			sTestCaseName[16] = ReportName.ANALYSIS_TransportWithJobsAndStatusHistory;
			sTestCaseName[17] = ReportName.ANALYSIS_LoadHistory;
            //====JOBS=========================================//
            sTestCaseName[18] = ReportName.PERFORMANCE_JOBS_CountByLocationInGroupDay;
            sTestCaseName[19] = ReportName.PERFORMANCE_JOBS_CountByLocationInGroupMonth;
            sTestCaseName[20] = ReportName.PERFORMANCE_JOBS_CountByLocationDay;
            sTestCaseName[21] = ReportName.PERFORMANCE_JOBS_CountByLocationMonth;
			//====TRANSPORTS=========================================//
            //------------------------------------
            sTestCaseName[22] = ReportName.PERFORMANCE_TRANSPORTS_DurationDistributionBySrcDstGroup;
            sTestCaseName[23] = ReportName.PERFORMANCE_TRANSPORTS_DurationDistributionSrcDstLocationOrStation;
            //--------------------
			sTestCaseName[24] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupHour;
			sTestCaseName[25] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupDay;
			sTestCaseName[26] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupMonth;
			sTestCaseName[27] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationHour;
			sTestCaseName[28] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationDay;
			sTestCaseName[29] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationMonth;
            sTestCaseName[30] = ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestGroupDay;
            sTestCaseName[31] = ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestLocationOrStationDay;
            sTestCaseName[32] = ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestGroupHour;
            sTestCaseName[33] = ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestLocationOrStationHour;
            //====_VEHICLES=========================================//
            sTestCaseName[34] = ReportName.PERFORMANCE_VEHICLES_ModeOverview;
            sTestCaseName[35] = ReportName.PERFORMANCE_VEHICLES_StateOverview;
            sTestCaseName[36] = ReportName.PERFORMANCE_VEHICLES_StatusOverview;
            sTestCaseName[37] = ReportName.PERFORMANCE_VEHICLES_StatusCountDayTrend;
            sTestCaseName[38] = ReportName.PERFORMANCE_VEHICLES_StatusCountTop;
            sTestCaseName[39] = ReportName.PERFORMANCE_VEHICLES_StatusDurationDayTrend;
            sTestCaseName[40] = ReportName.PERFORMANCE_VEHICLES_StatusDurationTop;
			
			try
			{
				if (!sOnlyUITest)
				{
					ProcessUtilities.CloseProcess( "EXCEL" );
                    ProcessUtilities.CloseProcess("Egemin.Etricc.Statistics.ParserConfigurator");
					ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SERVER );
					ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SHELL );
					Thread.Sleep(1000);
				}

                if (sUseExcel)
                {
                    xApp = new Excel.Application();
                    string[] sHeaderContents = { System.DateTime.Now.ToString("MMMM-dd") + "*" + "ETRICC STATISTICS" +  " UI Test Scenarios",
                                              "Test Machine:" + "*" + PCName,
                                               "Tester::" + "*" + System.Security.Principal.WindowsIdentity.GetCurrent().Name,
                                               "OSName:" + "*" + OSName,
                                               "OS Version:" + "*" + OSVersion,
                                               "UI Culture:" + "*" + UICulture,
                                               "Time On PC:" + "*" + "local time:" + TimeOnPC,
                                               "Test Tool Version:" + "*" +sTestToolsVersion,
                                               "NetworkMap:" + "*" +sNetworkMap,
                                                "Build Location:" + "*" +sInstallMsiDir.Substring(3),
                                                 "Test Project:" + "*" +sCurrentProject,
                                          };

                    sHeaderContentsLength = sHeaderContents.Length;
                    Epia3Common.WriteTestLogMsg(slogFilePath, "sHeaderContents.length: " + sHeaderContentsLength, sOnlyUITest);
                    if ( sUseExcel)
                        FileManipulation.WriteExcelHeader(ref xApp, sExcelVisible, sHeaderContents);
                }

				// start test----------------------------------------------------------
				int sResult = ConstCommon.TEST_UNDEFINED;
				int aantal = 41;
				if (sDemo)
					aantal = 2;

				if (sOnlyUITest)
				{
					sTestType = System.Configuration.ConfigurationManager.AppSettings.Get("TestType");
					if (sTestType.ToLower().StartsWith("all"))
						aantal = 41;
					else
					{
						int thisTest = 0;
						if (sTestType.IndexOf("-") > 0)
						{
							Console.WriteLine("first num: " + (sTestType.Substring(0, sTestType.IndexOf("-"))));
							Console.WriteLine("second num: " + (sTestType.Substring(sTestType.IndexOf("-") + 1)));

							thisTest = Convert.ToInt16(sTestType.Substring(0, sTestType.IndexOf("-")));
							Counter = thisTest-1;
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
					/*Epia3Common.WriteTestLogMsg(slogFilePath, " has build quality", sOnlyUITest);
					if (TFSConnected)
					{
						Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
							TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS), sBuildNr);
						string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

						if (quality.Equals("GUI Tests Failed"))
						{
							Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has build quality: " + quality + " , no update needed", sOnlyUITest);
						}
						else
						{
							TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
								TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS),
								"GUI Tests Started", m_BuildSvc, sDemo ? "true" : "false");
						}

						if (sAutoTest)
						{
							//TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", "EtriccStatistics+" + sCurrentPlatform);
							Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.ETRICCSTATISTICS, sOnlyUITest);
						}
					}*/
				}

               
				while (Counter < aantal)
				{
					sResult = ConstCommon.TEST_UNDEFINED;
					switch (sTestCaseName[Counter])
					{
						case UNINSTALL_ETRICC5_IF_ALREADY_INSTALLED:
							UninstallEtricc5IfAlreadyInstalled(UNINSTALL_ETRICC5_IF_ALREADY_INSTALLED, aeForm, out sResult);
							break;
						case INSTALL_ETRICC5:
							InstallEtricc5Application(INSTALL_ETRICC5, aeForm, out sResult);
							break;
						case PREPARE_TESTDATA:
							PrepareTestData(PREPARE_TESTDATA, aeForm, out sResult);
							break;
						case PARSERCONFIGURATOR_CONNECT_COMPUTER:
							ParserConfiguratorConnectComputer(PARSERCONFIGURATOR_CONNECT_COMPUTER, aeForm, out sResult);
							break;
						case CREATE_XML_SCHEMA_DEFINITION:
							CreateXmlSchemaDefinition(CREATE_XML_SCHEMA_DEFINITION, aeForm, out sResult);
							break;
						case CREATE_NEW_PARSE_PROJECT:
							CreateNewParseProject(CREATE_NEW_PARSE_PROJECT, aeForm, out sResult);
							break;
						case CREATE_DATABASE:
							CreateDatabase(CREATE_DATABASE, aeForm, out sResult);
							break;
						case SET_DATABASE:
							SetDatabase(SET_DATABASE, aeForm, out sResult);
							break;
						case CREATE_ETRICC_LAYOUT:
							CreateEtriccLayout(CREATE_ETRICC_LAYOUT, aeForm, out sResult);
							break;
						case PARSE_TEST_PROJECT:
							ParseTestProject(PARSE_TEST_PROJECT, aeForm, out sResult);
							break;
						case MODIFY_CONFIGURATION_FILES:
							ModifyConfigurationFiles(MODIFY_CONFIGURATION_FILES, aeForm, out sResult);
							break;
						case START_EPIA_SERVER_SHELL:
							StartEpiaServerShell(START_EPIA_SERVER_SHELL, aeForm, out sResult);
							break;
						case ReportName.StatusGraphicalView:
                            OpenThisStatisticsReport(null, null, ReportName.StatusGraphicalView, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_VEHICLES_ModeOverview:
                            OpenThisStatisticsReport("Performance", "Vehicles", ReportName.PERFORMANCE_VEHICLES_ModeOverview, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_VEHICLES_StateOverview:
                            OpenThisStatisticsReport("Performance", "Vehicles", ReportName.PERFORMANCE_VEHICLES_StateOverview, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_VEHICLES_StatusOverview:
                            OpenThisStatisticsReport("Performance", "Vehicles", ReportName.PERFORMANCE_VEHICLES_StatusOverview, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_VEHICLES_StatusCountDayTrend:
                            OpenThisStatisticsReport("Performance", "Vehicles", ReportName.PERFORMANCE_VEHICLES_StatusCountDayTrend, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_VEHICLES_StatusCountTop:
                            OpenThisStatisticsReport("Performance", "Vehicles", ReportName.PERFORMANCE_VEHICLES_StatusCountTop, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_VEHICLES_StatusDurationDayTrend:
                            OpenThisStatisticsReport("Performance", "Vehicles", ReportName.PERFORMANCE_VEHICLES_StatusDurationDayTrend, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_VEHICLES_StatusDurationTop:
                            OpenThisStatisticsReport("Performance", "Vehicles", ReportName.PERFORMANCE_VEHICLES_StatusDurationTop, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupHour:
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Transports", ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupHour, aeForm, out sResult);
                            else
                                OpenThisStatisticsReport("Performance", "Transports", "Hourly", ReportName.PERFORMANCE_TRANSPORTS_HOURLY_CountBySrcDstGroup, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupDay:
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Transports", ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupDay, aeForm, out sResult);
                            else
                                OpenThisStatisticsReport("Performance", "Transports", "Daily", ReportName.PERFORMANCE_TRANSPORTS_HOURLY_CountBySrcDstGroup, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupMonth:
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Transports", ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupMonth, aeForm, out sResult);
                            else
                                OpenThisStatisticsReport("Performance", "Transports", "Monthly", ReportName.PERFORMANCE_TRANSPORTS_HOURLY_CountBySrcDstGroup, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationHour:
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Transports", ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationHour, aeForm, out sResult);
                            else
                                OpenThisStatisticsReport("Performance", "Transports", "Hourly", ReportName.PERFORMANCE_TRANSPORTS_HOURLY_CountBySrcDstLocationOrStation, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationDay:
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Transports", ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationDay, aeForm, out sResult);
                            else
                                OpenThisStatisticsReport("Performance", "Transports", "Daily", ReportName.PERFORMANCE_TRANSPORTS_DAILY_CountBySrcDstLocationOrStation, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationMonth:
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Transports", ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationMonth, aeForm, out sResult);
                            else
                                OpenThisStatisticsReport("Performance", "Transports", "Monthly", ReportName.PERFORMANCE_TRANSPORTS_MONTHLY_CountBySrcDstLocationOrStation, aeForm, out sResult);
							break;
                        case ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestGroupDay:
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Transports", ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestGroupDay, aeForm, out sResult);
                            else
                                OpenThisStatisticsReport("Performance", "Transports", "Daily", ReportName.PERFORMANCE_TRANSPORTS_DAILY_DurationBySrcDestGroup, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestLocationOrStationDay:
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Transports", ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestLocationOrStationDay, aeForm, out sResult);
                            else
                                OpenThisStatisticsReport("Performance", "Transports", "Daily", ReportName.PERFORMANCE_TRANSPORTS_DAILY_DurationBySrcDestLocationOrStation, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestGroupHour:
                            if (sThisBuildBeforeTestCaseDate)
                                NotTestReason(ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestGroupHour, aeForm, out sResult, "build before test case date, not test ");
                            else
                                OpenThisStatisticsReport("Performance", "Transports", "Hourly", ReportName.PERFORMANCE_TRANSPORTS_HOURLY_DurationBySrcDestGroup, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestLocationOrStationHour:
                            if (sThisBuildBeforeTestCaseDate)
                                NotTestReason(ReportName.PERFORMANCE_TRANSPORTS_DurationBySrcDestGroupHour, aeForm, out sResult, "build before test case date, not test ");
                            else
                                OpenThisStatisticsReport("Performance", "Transports", "Hourly", ReportName.PERFORMANCE_TRANSPORTS_HOURLY_DurationBySrcDestLocationOrStation, aeForm, out sResult);
                            break;

						case ReportName.PERFORMANCE_JOBS_CountByLocationInGroupDay:
                            Console.WriteLine("sThisBuildBeforeTestCaseDate  : " + sThisBuildBeforeTestCaseDate);
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Jobs", ReportName.PERFORMANCE_JOBS_CountByLocationInGroupDay, aeForm, out sResult);
                            else // after test date, should check date 2013 2 12
                            {

                                if ( sCurrentVersionDate >= new DateTime(2013, 2, 13) &&  (sReleaseVersionDate == DateTime.MinValue || sReleaseVersionDate >= new DateTime(2013, 2, 13)) ) 
                                {
                                    OpenThisStatisticsReport("Performance", "Jobs", "Daily", ReportName.PERFORMANCE_JOBS_HOURLY_CountByLocationInGroup, aeForm, out sResult);
                                }
                                else // before 2013 2 13 --> release version is hourly and current version is hourly
                                    OpenThisStatisticsReport("Performance", "Jobs", "Hourly", ReportName.PERFORMANCE_JOBS_HOURLY_CountByLocationInGroup, aeForm, out sResult);
                            }
							break;
						case ReportName.PERFORMANCE_JOBS_CountByLocationInGroupMonth:
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Jobs", ReportName.PERFORMANCE_JOBS_CountByLocationInGroupMonth, aeForm, out sResult);
                            else
                                OpenThisStatisticsReport("Performance", "Jobs", "Monthly", ReportName.PERFORMANCE_JOBS_MONTHLY_CountByLocationInGroup, aeForm, out sResult);
							break;
						case ReportName.PERFORMANCE_JOBS_CountByLocationDay:
                            if (sCurrentVersionDate >= new DateTime(2013, 2, 13) )
                                Epia3Common.WriteTestLogMsg(slogFilePath, "sCurrentVersionDate >= new DateTime(2013, 2, 13) )", sOnlyUITest);
                            else
                                Epia3Common.WriteTestLogMsg(slogFilePath, "sCurrentVersionDate < new DateTime(2013, 2, 13) )", sOnlyUITest);

                            if (sReleaseVersionDate == DateTime.MinValue)
                                Epia3Common.WriteTestLogMsg(slogFilePath, "sReleaseVersionDate == DateTime.MinValue", sOnlyUITest);
                            else
                                Epia3Common.WriteTestLogMsg(slogFilePath, "sReleaseVersionDate != DateTime.MinValue", sOnlyUITest);


                            if (sReleaseVersionDate >= new DateTime(2013, 2, 13))
                                Epia3Common.WriteTestLogMsg(slogFilePath, "sReleaseVersionDate >= new DateTime(2013, 2, 13)", sOnlyUITest);
                            else
                                Epia3Common.WriteTestLogMsg(slogFilePath, "sReleaseVersionDate < new DateTime(2013, 2, 13)", sOnlyUITest);


                            if (sThisBuildBeforeTestCaseDate) // 
                                OpenThisStatisticsReport("Performance", "Jobs", ReportName.PERFORMANCE_JOBS_CountByLocationDay, aeForm, out sResult);
                            else
                            {
                                if (sCurrentVersionDate >= new DateTime(2013, 2, 13) && (sReleaseVersionDate == DateTime.MinValue || sReleaseVersionDate >= new DateTime(2013, 2, 13)))
                                {
                                    OpenThisStatisticsReport("Performance", "Jobs", "Daily", ReportName.PERFORMANCE_JOBS_HOURLY_CountByLocation, aeForm, out sResult);
                                }
                                else
                                    OpenThisStatisticsReport("Performance", "Jobs", "Hourly", ReportName.PERFORMANCE_JOBS_HOURLY_CountByLocation, aeForm, out sResult);
                            }
							break;
						case ReportName.PERFORMANCE_JOBS_CountByLocationMonth:
                            if (sThisBuildBeforeTestCaseDate)
                                OpenThisStatisticsReport("Performance", "Jobs", ReportName.PERFORMANCE_JOBS_CountByLocationMonth, aeForm, out sResult);
                            else
                                OpenThisStatisticsReport("Performance", "Jobs", "Monthly", ReportName.PERFORMANCE_JOBS_MONTHLY_CountByLocation, aeForm, out sResult);
                              
                            Thread.Sleep(2000);
                            AutomationElement aeJobs = StatUtilities.RefetchNodeTreeView("MainForm", "m_TreeView", "Jobs", 120, ref sErrorMessage);
                            if (aeJobs == null)
                            {
                                Console.WriteLine("  Jobs NOT FOUND After all Job report are tested");
                                Thread.Sleep(60000);
                            }
                            else
                            {
                                System.Windows.Point JobsPt = AUIUtilities.GetElementCenterPoint(aeJobs);
                                Input.MoveToAndDoubleClick(JobsPt);
                                Thread.Sleep(2000);
                            }                           
							break;
						case ReportName.ANALYSIS_ProjectActivation:
                            OpenThisStatisticsReport("Analysis", null, ReportName.ANALYSIS_ProjectActivation, aeForm, out sResult);
							break;
						case ReportName.ANALYSIS_TransportLookupBySrcDstGroup:
                            OpenThisStatisticsReport("Analysis", null, ReportName.ANALYSIS_TransportLookupBySrcDstGroup, aeForm, out sResult);
							break;
						case ReportName.ANALYSIS_TransportLookupBySrcDstLocationOrStation:
                            OpenThisStatisticsReport("Analysis", null, ReportName.ANALYSIS_TransportLookupBySrcDstLocationOrStation, aeForm, out sResult);
							break;
						case ReportName.ANALYSIS_TransportWithJobsAndStatusHistory:
                            OpenThisStatisticsReport("Analysis", null, ReportName.ANALYSIS_TransportWithJobsAndStatusHistory, aeForm, out sResult);
							break;
						case ReportName.ANALYSIS_LoadHistory:
                            OpenThisStatisticsReport("Analysis", null, ReportName.ANALYSIS_LoadHistory, aeForm, out sResult);
							break;
                        case ReportName.PERFORMANCE_TRANSPORTS_DurationDistributionBySrcDstGroup:
                            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc Stat Prog.Production.Release", new DateTime(2014, 1, 15), ref sErrorMessage) == true)
                                sResult = ConstCommon.TEST_UNDEFINED;
                            else
                                OpenThisStatisticsReport("Performance", "Transports", ReportName.PERFORMANCE_TRANSPORTS_DurationDistributionBySrcDstGroup, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_TRANSPORTS_DurationDistributionSrcDstLocationOrStation:
                            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Etricc 5", "Etricc Stat Prog.Production.Release", new DateTime(2014, 1, 15), ref sErrorMessage) == true)
                                sResult = ConstCommon.TEST_UNDEFINED;
                            else
                                OpenThisStatisticsReport("Performance", "Transports", ReportName.PERFORMANCE_TRANSPORTS_DurationDistributionSrcDstLocationOrStation, aeForm, out sResult);
                            break;
						default:
							break;
					}

					// write result to Excel
					int k = 0;
					bool wr = false;
					while (wr == false)
					{
                        if (sUseExcel)
                        {
                            try
                            {
                                FileManipulation.WriteExcelTestCaseResult(xApp, sResult, sHeaderContentsLength, Counter, sTestCaseName[Counter], sErrorMessage);
                                wr = true;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("write excel result exception: " + ex.Message + "-----" + ex.StackTrace);
                                Thread.Sleep(5000);
                                if (k++ < 10)
                                    wr = false;
                                else
                                    wr = true;
                            }
                        }
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
				}

				Epia3Common.WriteTestLogMsg(slogFilePath, "Total Tests: " + Counter, sOnlyUITest);
				Epia3Common.WriteTestLogMsg(slogFilePath, "Total Passed: " + sTotalPassed, sOnlyUITest);
				Epia3Common.WriteTestLogMsg(slogFilePath, "Total Failed: " + sTotalFailed, sOnlyUITest);

                if (sUseExcel)
                    FileManipulation.WriteExcelFoot(xApp, sHeaderContentsLength, Counter, sTotalCounter, sTotalPassed, sTotalFailed);
               
				if (!sOnlyUITest)
				{
					string msgX = "update etricc statistics build quality test status to Passed if needed";
					Console.WriteLine(msgX);
					TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX);
					while (TFSConnected == false)
					{
						TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
								"update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay );
						System.Threading.Thread.Sleep( Tfs.ReconnectDelay );
						TFSConnected = TfsUtilities.CheckTFSConnection( ref msgX );
					}

					if (TFSConnected)
					{
						// added check sTestResultFolder exist; some time during testing this build can be completely deleted by WVB
						if (Directory.Exists(sTestResultFolder))
						{
							#region  // update testinfo file first and then update build quality
							string testout = "-->" + sOutFilename + ".xls";
							if (sAutoTest)
							{
								Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
								TfsUtilities.GetProjectName(ConstCommon.ETRICCSTATISTICS), sBuildNr);
								string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

								if (sTotalFailed == 0)
								{
									Epia3Common.WriteTestLogMsg(slogFilePath, "sTotalFailed == 0: UpdateStatusInTestInfoFile :sInfoFileKey=" + sInfoFileKey, sOnlyUITest);
									TestListUtilities.UpdateStatusInTestInfoFile( sTestResultFolder, "GUI Tests Passed", "Tests OK", sInfoFileKey );
									Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Passed:" + ConstCommon.ETRICCSTATISTICS, sOnlyUITest);

									Console.WriteLine(" Update build quality:  quality: " + quality);
									if (quality.Equals("GUI Tests Failed"))
									{
										Console.WriteLine("Quality is GUI Tests Failed : No Update Needed ");
										Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
									}
									else if ( TestListUtilities.IsAllTestDefinitionsTested( mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage ) == false )
									{
										Console.WriteLine("NOT All Test definitions tested " + sErrorMessage);
									}
									else
									{
										if ( TestListUtilities.IsAllTestStatusPassed( mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage ) == true )
										{
											// update quality to GUI Tests Passed
											TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(ConstCommon.ETRICCSTATISTICS),
											"GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");

											Console.WriteLine("update quality to true -----  ");
											Thread.Sleep(1000);
										}
									}
								}
								else
								{
									TestListUtilities.UpdateStatusInTestInfoFile( sTestResultFolder, "GUI Tests Failed", "--->" + sOutFilename + ".log", sInfoFileKey );
									Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Failed:" + ConstCommon.ETRICCSTATISTICS, sOnlyUITest);

									Console.WriteLine(" Update build quality:  quality: " + quality);
									if (quality.Equals("GUI Tests Failed"))
									{
										Console.WriteLine("Quality is GUI Tests Failed : No Update Needed ");
										Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
									}
									else
									{
										// update quality to GUI Tests Passed
										TfsUtilities.UpdateBuildQualityStatus(logger, uri,
										TfsUtilities.GetProjectName(ConstCommon.ETRICCSTATISTICS),
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

				// save Excel to Local machine
                string sXLSPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), sOutFilename + ".xls");
                Epia3Common.WriteTestLogMsg(slogFilePath, "Save to local machine : " + sXLSPath, sOnlyUITest);
                if (sUseExcel)
                {
                    if (FileManipulation.SaveExcel(xApp, sXLSPath, ref sErrorMessage) == false)
                    {
                        string sTXTPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), sOutFilename + ".txt");
                        StreamWriter write = File.CreateText(sTXTPath);
                        write.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        write.Close();
                    }
                }
                // Save to remote machine
                if (System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToUpper().StartsWith("TEAMSYSTEMS\\JIEMINSHI"))
                {
                    Console.WriteLine("\n   not write to remote machine");
                }
                else
                {
                    string sXLSPath2 = System.IO.Path.Combine(sOutFilePath, sOutFilename + ".xls");
                    Console.WriteLine("Save2 : " + sXLSPath2);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "sXLSPath2 =: " + sXLSPath2, sOnlyUITest);
                    {
                        if (FileManipulation.SaveExcel(xApp, sXLSPath2, ref sErrorMessage) == false)
                        {
                            string sTXTPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), sOutFilename + ".txt");
                            StreamWriter write = File.CreateText(sTXTPath);
                            write.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                            write.Close();
                        }
                    }
                }

				// quit Excel.
				xApp.Quit();

				// Send Result via Email
				if (!sOnlyUITest)
					SendEmail(sXLSPath);
				
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
				ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SHELL );
				Thread.Sleep(10000);
				ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SERVER );
				ProcessUtilities.CloseProcess( "cmd" );
				Console.WriteLine("\nEnd test run\n");
				//Console.ReadLine();
			}
			catch (Exception ex)
			{
				Console.WriteLine("main Fatal error: " + ex.Message + "----: " + ex.StackTrace  +"---" + ex.ToString());
				Thread.Sleep(2000);
				if (sAutoTest)
				{
					#region // test exception : update infofile and build quality
					TestListUtilities.UpdateStatusInTestInfoFile( sTestResultFolder, "GUI Tests Exception", " -->" + sOutFilename + ".log", sInfoFileKey );
					Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Exception -->" + sOutFilename + ".log:" + ConstCommon.ETRICCSTATISTICS, sOnlyUITest);

					Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace, sOnlyUITest);

					ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SHELL );
					ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SERVER );
					ProcessUtilities.CloseProcess( "cmd" );

					string msgX = "epia exception build quality test status to Failed if needed";
					TFSConnected = TfsUtilities.CheckTFSConnection( ref msgX );
					while (TFSConnected == false)
					{
						TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server after 1 minutes",
								"update build quality Deployment Failed2", 60000);
						System.Threading.Thread.Sleep(60000);
						TFSConnected = TfsUtilities.CheckTFSConnection( ref msgX );
					}

					if (TFSConnected)
					{
						Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
							TfsUtilities.GetProjectName(ConstCommon.ETRICCSTATISTICS), sBuildNr);
						string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

						if (quality.Equals("GUI Tests Failed"))
						{
							Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
						}
						else
						{
							TfsUtilities.UpdateBuildQualityStatus(logger, uri,
								TfsUtilities.GetProjectName(ConstCommon.ETRICCSTATISTICS),
								"GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
						}
					}
					#endregion
				   
				}
			}
		}
		#endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
        public static void NotTestReason(string testname, AutomationElement root, out int result, string reason)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            TestCheck = ConstCommon.TEST_PASS;
            result = ConstCommon.TEST_UNDEFINED;
           
            sErrorMessage = reason;
            return;
        }
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region TestCase Name
		private const string UNINSTALL_ETRICC5_IF_ALREADY_INSTALLED = "UninstallEtricc5IfAlreadyInstalled";
		private const string INSTALL_ETRICC5 = "InstallEtricc5";
		private const string PARSERCONFIGURATOR_CONNECT_COMPUTER = "ParserConfiguratorConnectComputer";
		private const string CREATE_XML_SCHEMA_DEFINITION = "CretateXmlSchemaDefinition";
		private const string CREATE_NEW_PARSE_PROJECT = "CreateNewParseProject";
		private const string PREPARE_TESTDATA = "PrepareTestData";
		private const string CREATE_DATABASE = "CreateDatabase";
		private const string SET_DATABASE = "SetDatabase";
		private const string CREATE_ETRICC_LAYOUT = "CreateEtriccLayout";
		private const string PARSE_TEST_PROJECT = "ParseTestProject";
		private const string MODIFY_CONFIGURATION_FILES = "UpdateConfigurationFiles";
		private const string START_EPIA_SERVER_SHELL = "StartEpiaServerShell";
		//private const string OPEN_GRAPHICAL_VIEW = "OpenGraphicalView";
		//private const string OPEN_PERFORMANCE_VEHICLES_MODE_OVERVIEW = "OpenPerformanceVehiclesModeOverview";
		//private const string OPEN_PERFORMANCE_VEHICLES_STATE_OVERVIEW = "OpenPerformanceVehiclesStateOverview";
		//private const string OPEN_PERFORMANCE_VEHICLES_STATUS_OVERVIEW = "OpenPerformanceVehiclesStatusOverview";
		//private const string OPEN_PERFORMANCE_VEHICLES_STATUS_COUNT_DAY_TREND = "OpenPerformanceVehiclesStatusCountDayTrend";
		//private const string OPEN_PERFORMANCE_VEHICLES_STATUS_COUNT_TOP = "OpenPerformanceVehiclesStatusCountTop";
		//private const string OPEN_PERFORMANCE_VEHICLES_STATUS_DURATION_DAY_TREND = "OpenPerformanceVehiclesStatusDurationDayTrend";

		//private const string OPEN_PERFORMANCE_TRANSPORTS_COUNT_BY_SRC_DST_GROUP_HOUR = "Count by src/dst group (hour)";
		//private const string OPEN_PERFORMANCE_TRANSPORTS_COUNT_BY_SRC_DST_GROUP_DAY = "Count by src/dst group (day)";
		//private const string OPEN_PERFORMANCE_TRANSPORTS_COUNT_BY_SRC_DST_GROUP_MONTH = "Count by src/dst group (month)";
		//private const string OPEN_PERFORMANCE_TRANSPORTS_COUNT_BY_SRC_DST_LOCATION_OR_STATION_HOUR = "Count by src/dst location or station (hour)";
		//private const string OPEN_PERFORMANCE_TRANSPORTS_COUNT_BY_SRC_DST_LOCATION_OR_STATION_DAY = "Count by src/dst location or station (daty)";
		//private const string OPEN_PERFORMANCE_TRANSPORTS_COUNT_BY_SRC_DST_LOCATION_OR_STATION_MONTH = "Count by src/dst location or station (month)";
		//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

		#endregion TestCase Name
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Test Cases -------------------------------------------------------------------------------------------
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region UninstallEtricc5IfAlreadyInstalled
		public static void UninstallEtricc5IfAlreadyInstalled(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
			result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
           
			try
			{
                // uninstall etricc core if already installed:
                if (ProjAppInstall.UninstallApplication(EgeminApplication.ETRICC, ref sErrorMessage) == false)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    sErrorMessage = EgeminApplication.ETRICC + " Uninstall failed:" + sErrorMessage;
                    Console.WriteLine(sErrorMessage);
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
		#region InstallEtricc5Application
		public static void InstallEtricc5Application(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			AutomationEventHandler UIALayoutXPosEventHandler = new AutomationEventHandler(OnInstallEtricc5UIEvent);

			try
			{
				// unzip test data first
				if (System.IO.Directory.Exists(System.IO.Directory.GetCurrentDirectory() + @"\EtriccStatistics"))
				{
					Console.WriteLine("EtriccStatistics folder exist, delete first: ");
					System.IO.Directory.Delete(System.IO.Directory.GetCurrentDirectory() + @"\EtriccStatistics", true);
				}

				try
				{
					// unzip project file
					//string zipFile = EtriccStatistics.zip;
					string zipFile = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "EtriccStatistics.zip");
					FastZip fz = new FastZip();
					fz.ExtractZip(zipFile, System.IO.Directory.GetCurrentDirectory(), "");
				}
				catch (Exception ex)
				{
					System.Windows.Forms.MessageBox.Show("EtriccStatistics.zip unzip exception: " + ex.Message + "--------------" + ex.StackTrace);

				}
				Thread.Sleep(5000);

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
				Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (System.IO.File.Exists(System.IO.Directory.GetCurrentDirectory() + @"\EtriccStatistics\LeenBakker\Data\Xml\LeenBakker.Xml") == false 
					&& mTime.Seconds <= 600)
				{
					Thread.Sleep(5000);
					mTime = DateTime.Now - mStartTime;
					Console.WriteLine("Testdata not unzipped yet :");
				}

				if (System.IO.Directory.Exists(System.IO.Directory.GetCurrentDirectory() + @"\EtriccStatistics"))
				{
					TestCheck = ConstCommon.TEST_PASS;
				}
				else
				{
					TestCheck = ConstCommon.TEST_FAIL;
					sErrorMessage = "Test data not existed !!!";

				}

				// Add Open window Event Handler
				Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					AutomationElement.RootElement, TreeScope.Descendants, UIALayoutXPosEventHandler);

				System.Threading.Thread.Sleep(15000);
                string pa = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE");
                string systemBits = "" + ((String.IsNullOrEmpty(pa) || String.Compare(pa, 0, "x86", 0, 3, true) == 0) ? 32 : 64);

                string InstallerSource = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory() + @"\EtriccStatistics\", "Etricc32.msi");
                if (systemBits.StartsWith("64"))
                    InstallerSource = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory() + @"\EtriccStatistics\", "Etricc64.msi");

				Console.WriteLine("start:" + InstallerSource);
				Process Proc = new System.Diagnostics.Process();
				Proc.StartInfo.FileName = InstallerSource;
				Proc.StartInfo.CreateNoWindow = false;
				Proc.Start();
				Console.WriteLine("started:" + InstallerSource);

				sEventEnd = false;
				mStartTime = DateTime.Now;
				mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);
				while (sEventEnd == false && mTime.TotalMinutes <= 5)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
					Console.WriteLine("wait time is :" + mTime.TotalSeconds + "   --------- sEventEnd:" + sEventEnd);

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
					Console.WriteLine("\nInstall EtriccCore.: Pass");
					result = ConstCommon.TEST_PASS;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
				   
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
		#region PrepareTestData
		public static void PrepareTestData(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			TestCheck = ConstCommon.TEST_PASS;
			result = ConstCommon.TEST_UNDEFINED;

			try
			{
                //========================  SQL SERVER =================================================
                if (ProjServerOrShellStartup.CheckThisServiceIsStartedUp("MSSQLSERVER", ref sErrorMessage, slogFilePath, sOnlyUITest) == false)
                {
                    sEpiaServerStartupOK = false;
                    sSqlOrReportServiceOK = false;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                //========================  Report SERVER =================================================
                if (TestCheck == ConstCommon.TEST_PASS)
				{
                    if (ProjServerOrShellStartup.CheckThisServiceIsStartedUp("ReportServer", ref sErrorMessage, slogFilePath, sOnlyUITest) == false)
                    {
                        sEpiaServerStartupOK = false;
                        sSqlOrReportServiceOK = false;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    // unzip test data first
                    if (System.IO.Directory.Exists(System.IO.Directory.GetCurrentDirectory() + @"\EtriccStatistics"))
                    {
                        Console.WriteLine("EtriccStatistics folder exist, do nothing: ");
                    }
                    else
                    {
                        Console.WriteLine("EtriccStatistics folder not exist, unzip test data: ");
                        try
                        {
                            // unzip project file
                            //string zipFile = EtriccStatistics.zip;
                            string zipFile = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "EtriccStatistics.zip");
                            FastZip fz = new FastZip();
                            fz.ExtractZip(zipFile, System.IO.Directory.GetCurrentDirectory(), "");
                        }
                        catch (Exception ex)
                        {
                            sErrorMessage = "EtriccStatistics.zip unzip exception: " + ex.Message + "--------------" + ex.StackTrace;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                            //System.Windows.Forms.MessageBox.Show("EtriccStatistics.zip unzip exception: " + ex.Message + "--------------" + ex.StackTrace);
                        }
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(5000);
                    string batFile = Path.Combine(Directory.GetCurrentDirectory(), "DropDatabaseEtriccStatistics_" + sCurrentProject + ".bat");
                    string sqlFile = Path.Combine(Directory.GetCurrentDirectory(), "DropDatabaseEtriccStatistics_" + sCurrentProject + ".sql");

                    //string batLine1 = "C:\Program Files\Microsoft SQL Server\100\Tools\Binn\sqlcmd"  -Stcp:%computername%,1433 -E -i DropDatabaseEtriccStatistics_Demo.sql
                    // sqlcmd are in same folder for  both 32bit or 64 bit SQL server 
                    //string batLine1 = '"' + @"C:\Program Files\Microsoft SQL Server\100\Tools\Binn\sqlcmd" + '"' + "\t" 
                    // search sqlcmd.exe fullpath
                    string sqlCMDpath = DeployUtilities.getFullPathFromFilename("SQLCMD.EXE", @"C:\Program Files\Microsoft SQL Server", string.Empty);
                    if (sqlCMDpath == null || sqlCMDpath == "")
                        sqlCMDpath = DeployUtilities.getFullPathFromFilename("SQLCMD.EXE", @"C:\Program Files (x86)\Microsoft SQL Server", string.Empty);

                    if (sqlCMDpath == null || sqlCMDpath.Length < 10)
                        System.Windows.Forms.MessageBox.Show("SQLCMD.EXE not found!!! please install SQL Server sqlCMDpath:" + sqlCMDpath);

                    string batLine1 = '"' + sqlCMDpath + '"' + "\t"
                        + "-Stcp:%computername%,1433 -E -i DropDatabaseEtriccStatistics_" + sCurrentProject + ".sql";


                    //System.Windows.Forms.MessageBox.Show("batLine1:" + batLine1);


                    string batLine2 = System.Environment.NewLine;
                    string batLine3 = "pause";
                    //if (!File.Exists(batFile))
                    //{
                        StreamWriter writeBat = File.CreateText(batFile);
                        writeBat.WriteLine(batLine1);
                        writeBat.WriteLine(batLine2);
                        writeBat.WriteLine(batLine3);
                        writeBat.Close();
                    //}

                    //IF EXISTS(SELECT 1 FROM sys.databases WHERE name = 'EtriccStatistics_Demo' )
                    //DROP database [EtriccStatistics_Demo]
                    string sqlLine1 = "IF EXISTS(SELECT 1 FROM sys.databases WHERE name = 'EtriccStatistics_" + sCurrentProject + "' )";
                    string sqlLine2 = "DROP database [EtriccStatistics_" + sCurrentProject + "]";
                    if (!File.Exists(sqlFile))
                    {
                        StreamWriter writeSQL = File.CreateText(sqlFile);
                        writeSQL.WriteLine(sqlLine1);
                        writeSQL.WriteLine(sqlLine2);
                        writeSQL.Close();
                    }

                    // drop database if exist
                    ProcessUtilities.StartProcessNoWait(System.IO.Directory.GetCurrentDirectory(),
                       "DropDatabaseEtriccStatistics_" + sCurrentProject + ".bat", string.Empty);
                    Thread.Sleep(10000);
                }

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
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region ParserConfiguratorConnectComputer
		public static void ParserConfiguratorConnectComputer(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;
			
			AutomationEventHandler UIEventHandler = new AutomationEventHandler(OnParserConfiguratorConnectComputerEvent);

			try
			{
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

				// Add Open MyLayoutScreen window Event Handler
				Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
				   AutomationElement.RootElement, TreeScope.Descendants, UIEventHandler);

				//Thread.Sleep(10000);
				Console.WriteLine("start :" + ConstCommon.PARSERCONFIGURATOR_EXE);
				string InstallerSource = System.IO.Path.Combine(TestTools.OSVersionInfoClass.ProgramFilesx86()
					+ConstCommon.PARSERCONFIGURATOR_ROOT, ConstCommon.PARSERCONFIGURATOR_EXE);
				Console.WriteLine("InstallerSource:" + InstallerSource);
				Process Proc = new System.Diagnostics.Process();
				Proc.StartInfo.FileName = InstallerSource;
				Proc.StartInfo.CreateNoWindow = false;
				Proc.Start();
				Console.WriteLine("started:" + ConstCommon.PARSERCONFIGURATOR_EXE);
                AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
                if (aeWindow == null)
                {
                    sErrorMessage = "MainForm not found after one minute";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {

                    sEventEnd = false;
                    DateTime mStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - mStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalSeconds);
                    while (sEventEnd == false && mTime.TotalMinutes <= 3)
                    {
                        Thread.Sleep(10000);
                        mTime = DateTime.Now - mStartTime;
                        Console.WriteLine("wait ParserConfiguratorConnectComputer time is :" + mTime.TotalSeconds + "   --------- sEventEnd:" + sEventEnd);

                    }
                }

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					 AutomationElement.RootElement,
					UIEventHandler);

				Console.WriteLine("Searching ParserConfiguratorConnectComputer windows ..................");
				Thread.Sleep(4000);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
                    if (aeWindow == null)
                    {
                        sErrorMessage = "MainForm not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        string treeViewId = "m_TreeView";
                        AutomationElement aeTreeView = null;
                        DateTime sTime = DateTime.Now;
                        AUIUtilities.WaitUntilElementByIDFound(aeWindow, ref aeTreeView, treeViewId, sTime, 60);

                        if (aeTreeView == null)
                        {
                            sErrorMessage = "aeTreeView not found name";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            AutomationElement aeNodeLink = null;
                            TreeWalker walker = TreeWalker.ControlViewWalker;
                            AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
                            while (elementNode != null)
                            {
                                Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
                                Console.WriteLine("PCName name is: " + PCName);
                                if (elementNode.Current.Name.ToLower().StartsWith(PCName.ToLower()))
                                {
                                    //Input.MoveTo(elementNode);
                                    aeNodeLink = elementNode;
                                    TestCheck = ConstCommon.TEST_PASS;
                                    break;
                                }
                                Thread.Sleep(3000);
                                elementNode = walker.GetNextSibling(elementNode);
                            }

                            if (aeNodeLink == null)
                            {
                                sErrorMessage = "Computeer name not found name";
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                    }
                }

				if (TestCheck == ConstCommon.TEST_FAIL)
				{
                    sPaeserConfiguratorConnectToComputerOK = false; ;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					result = ConstCommon.TEST_FAIL;
				}

				if (TestCheck == ConstCommon.TEST_PASS)
				{   
					Console.WriteLine(PCName + " is now connected");
					Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
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
					  UIEventHandler);
			}
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region CreateXmlSchemaDefinition
		public static void CreateXmlSchemaDefinition(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;
			try
			{
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

                if (sPaeserConfiguratorConnectToComputerOK == false)
                {
                    sErrorMessage = "parser configurator failed to connect to the computer , this testcase cannot be tested";
                    return;
                }
				Console.WriteLine("Searching main windows ..................");
				AutomationElement aeTreeView = null;
				AutomationElement aeComputerNameNode = null;
				AutomationElement aeXSDsNode = null;

                AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
				if (aeWindow == null)
				{
					sErrorMessage = "MainForm not found";
					Console.WriteLine(sErrorMessage);
					TestCheck = ConstCommon.TEST_FAIL;
				}
				else
				{
					aeWindow.SetFocus();
					string treeViewId = "m_TreeView";
					DateTime sTime = DateTime.Now;
					AUIUtilities.WaitUntilElementByIDFound(aeWindow, ref aeTreeView, treeViewId, sTime, 60);
					if (aeTreeView == null)
					{
						sErrorMessage = "aeTreeView not found name";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
                    {
                        #region // find computer from treeview
                        TreeWalker walker = TreeWalker.ControlViewWalker;
						AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
						while (elementNode != null)
						{
							Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
							if (elementNode.Current.Name.ToLower().Equals(PCName.ToLower()))
							{
								aeComputerNameNode = elementNode;
								Console.WriteLine("Computer node name found , it is: " + aeComputerNameNode.Current.Name);
								break;
							}
							Thread.Sleep(3000);
							elementNode = walker.GetNextSibling(elementNode);
						}

                        if (aeComputerNameNode == null)
                        {
                            sErrorMessage = "Computer node not found, ";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            try
                            {
                                ExpandCollapsePattern ep = (ExpandCollapsePattern)aeComputerNameNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                                ep.Expand();
                                Thread.Sleep(300);
                            }
                            catch (Exception ex)
                            {   // aeComputerNameNode should can be Expanded
                                sErrorMessage = "aeComputerNameNode can not Expanded: " + aeComputerNameNode.Current.Name+ "---"+ex.Message;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                        #endregion
                    }
                }

				if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // find xsd tree item from computernode
                    Console.WriteLine("\n=== Find XSDs node ===");
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					aeXSDsNode = TestTools.AUICommon.WalkEnabledElements(aeTreeView, treeNode, "XSDs");
                    if (aeXSDsNode == null)  // aeComputerNameNode Expanded, aeXSDsNode should always be found
					{
                        sErrorMessage = "=== XSDs node NOT Exist ===";
						Console.WriteLine("\n=== XSDs node NOT Exist ===");
                        Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Console.WriteLine("\n=== XSDs node Exist ===");
                        try
                        {
                            ExpandCollapsePattern ep = (ExpandCollapsePattern)aeXSDsNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                            ep.Expand();
                            Thread.Sleep(1000);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("XsdNode is empty, can not expaned: " + aeXSDsNode.Current.Name+ "----"+ex.Message);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "XsdNode is empty, can not expaned: " + aeXSDsNode.Current.Name, sOnlyUITest); ;
                        }

                        // delete all existed xsd node
                        TreeWalker walkerProj = TreeWalker.ControlViewWalker;
                        AutomationElement aeMyXsdNode = walkerProj.GetFirstChild(aeXSDsNode);
                        while (aeMyXsdNode != null)
                        {
                            Console.WriteLine("aeMyProjNode name is: " + aeMyXsdNode.Current.Name);
                            StatUtilities.DeleteSelectedXsd(aeWindow, aeMyXsdNode);
                            Thread.Sleep(1000);
                            aeMyXsdNode = walkerProj.GetFirstChild(aeXSDsNode);
                        }
                    }
                    #endregion
                }

                // Create new xsd file
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // XSDs --> Actions--> Create Xsd...
                    Thread.Sleep(500);
                    sEventEnd = false;
                    Point XSDsNodePnt = AUIUtilities.GetElementCenterPoint(aeXSDsNode);
                    Input.MoveToAndRightClick(XSDsNodePnt);
                    Thread.Sleep(1000);

                    AutomationElement aeMenuItemActions = ProjBasicUI.GetWindowMenuItemActionElement("MainForm", "Actions", ref sErrorMessage);
                    if (aeMenuItemActions != null)
                    {
                        Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeMenuItemActions));
                        Thread.Sleep(100);
                        AutomationElement aeMenuItemCreateXsd = ProjBasicUI.GetMenuItemFromElement(aeMenuItemActions, "Create Xsd...", "name", 30, ref sErrorMessage);
                        if (aeMenuItemCreateXsd != null)
                        {
                            Point CreateXsdPnt = AUIUtilities.GetElementCenterPoint(aeMenuItemCreateXsd);
                            Input.MoveTo(CreateXsdPnt);
                            Thread.Sleep(200);
                            Input.ClickAtPoint(CreateXsdPnt);
                            Thread.Sleep(500);
                        }
                        else
                        {
                            sErrorMessage = "aeMenuItemCreateXsd not found";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        sErrorMessage = "aeMenuItemActions not found";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    #endregion
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // select Egemin.EPIA.WCS.dll file
                    AutomationElement aeRemoteFileBrowserWindow = ProjBasicUI.GetPopupDialogFromMainWindow("MainForm", "RemoteBrowserForm", "id", ref sErrorMessage);
                    DateTime xStartTime = DateTime.Now;
                    TimeSpan xTime = DateTime.Now - xStartTime;
                    Console.WriteLine(" time is :" + xTime.TotalSeconds);
                    while (aeRemoteFileBrowserWindow == null && xTime.TotalMinutes <= 5)
                    {
                        Thread.Sleep(3000);
                        aeRemoteFileBrowserWindow = ProjBasicUI.GetPopupDialogFromMainWindow("MainForm", "RemoteBrowserForm", "id", ref sErrorMessage);
                        xTime = DateTime.Now - xStartTime;
                        Console.WriteLine("wait time is :" + xTime.TotalSeconds + "   --------- sEventEnd:" + sEventEnd);
                    }

                    if (aeRemoteFileBrowserWindow == null)
                    {
                        sErrorMessage = "aeRemoteFileBrowserWindow not found";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        string treeViewID = "m_TreeView";
                        List<string> FolderList = new List<string>();
                        FolderList.Insert(0, System.Environment.MachineName);
                        FolderList.Insert(1, "C: (Local Disk )");
                        FolderList.Insert(2, OSVersionInfoClass.ProgramFilesx86FolderName());
                        FolderList.Insert(3, "Egemin");
                        FolderList.Insert(4, "Etricc Server");
                        string listViewId = "m_ListView";
                        string SelectedFilename = "Egemin.EPIA.WCS.dll";
                        //string BtnConnectId = "m_BtnSelect";
                        if (StatUtilities.SelectFileInFileBrowserWindow(aeRemoteFileBrowserWindow, treeViewID, FolderList,
                            listViewId, SelectedFilename, "m_BtnSelect", ref sErrorMessage) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
#endregion
                }
				
				Console.WriteLine("checkoutput screen ..................");
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.CheckParserConfiguratorOutput("MainForm", ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                Thread.Sleep(3000);
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
				sErrorMessage = ex.Message;
				Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
				Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
			}
			finally
			{
				Thread.Sleep(3000);
			}
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region CreateNewParseProject
		public static void CreateNewParseProject(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			AutomationEventHandler UIEventHandler = new AutomationEventHandler(OnCreateNewParseProjectUIEvent);

			try
			{
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

                if (sPaeserConfiguratorConnectToComputerOK == false)
                {
                    sErrorMessage = "parser configurator failed to connect to the computer , this testcase cannot be tested";
                    return;
                }

				Console.WriteLine("Searching main windows ..................");
				AutomationElement aeTreeView = null;
				AutomationElement aeComputerNameNode = null;
				AutomationElement aeProjectsNode = null;
				#region //Search test project and delte this project if it exist
                AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
				if (aeWindow == null)
				{
					sErrorMessage = "MainForm not found";
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					TestCheck = ConstCommon.TEST_FAIL;
				}
				else
				{
					Console.WriteLine("----- set focus -----");
					aeWindow.SetFocus();
					string treeViewId = "m_TreeView";
					DateTime sTime = DateTime.Now;
					AUIUtilities.WaitUntilElementByIDFound(aeWindow, ref aeTreeView, treeViewId, sTime, 60);
					if (aeTreeView == null)
					{
						sErrorMessage = "aeTreeView not found name";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						TreeWalker walker = TreeWalker.ControlViewWalker;
						AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
						while (elementNode != null)
						{
							Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
							if (elementNode.Current.Name.ToLower().Equals(PCName.ToLower()))
							{
								aeComputerNameNode = elementNode;
								Console.WriteLine("Computer node name found , it is: " + aeComputerNameNode.Current.Name);
								break;
							}
							Thread.Sleep(3000);
							elementNode = walker.GetNextSibling(elementNode);
						}

						if (aeComputerNameNode == null)
						{
							sErrorMessage = "Computer node not found, ";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							try
							{
								ExpandCollapsePattern ep = (ExpandCollapsePattern)aeComputerNameNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
								ep.Expand();
							}
							catch (Exception ex)
							{
								sErrorMessage = "aeComputerNameNode can not expaned " + aeComputerNameNode.Current.Name +" --- "+ex.Message;
								Console.WriteLine(sErrorMessage);
								TestCheck = ConstCommon.TEST_FAIL;
							}
						}
					}
				}

				Thread.Sleep(5000);
				 // find Projects tree item from computernode
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find Projects node ===");
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					aeProjectsNode = TestTools.AUICommon.WalkEnabledElements(aeComputerNameNode, treeNode, "Projects");
					if (aeProjectsNode == null)
					{
						sErrorMessage = "Projects node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						try
						{
							ExpandCollapsePattern ep = (ExpandCollapsePattern)aeProjectsNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
							ep.Expand();
						}
						catch (Exception ex)
						{
                            sErrorMessage = "aeProjectsNode can not expaned " + aeProjectsNode.Current.Name + " --- " + ex.Message;
							Console.WriteLine(sErrorMessage);
						}
					}
				}

				Thread.Sleep(5000);
				// find Test Project tree item from Project Node
				//AutomationElement aeTestNode = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find Test project node ===" + sCurrentProject);
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();

					Condition condition1 = new PropertyCondition(AutomationElement.IsControlElementProperty, true);
					Condition condition2 = new PropertyCondition(AutomationElement.IsEnabledProperty, true);
					TreeWalker walker = new TreeWalker(new AndCondition(condition1, condition2));
					AutomationElement elementNode = walker.GetFirstChild(aeProjectsNode);
					AutomationElement elementNextNode = null;
					while (elementNode != null)
					{
						string nodeName = elementNode.Current.Name;
						elementNextNode = walker.GetNextSibling(elementNode);

						Console.WriteLine("delete this node, name is: " + nodeName);
						if (StatUtilities.DeleteSelectedProject(aeWindow, elementNode, ref sErrorMessage))
						{
							Console.WriteLine("\n=== Delete Test project node OK  " + nodeName);
							//System.Windows.Forms.TreeNode childTreeNode = treeNode.Nodes.Add(elementNode.Current.ControlType.LocalizedControlType);
							//ele = AUICommon.WalkEnabledElements(elementNode, childTreeNode, nodeName);
							//elementNode = walker.GetNextSibling(elementNode);
							if (elementNextNode != null)
							{
								Console.WriteLine("\n=== Next Test project node is  " + elementNextNode.Current.Name);
								elementNode = elementNextNode;
							}
							else
							{
								Console.WriteLine("\n=== NO Next Test project node  ");
								elementNode = null;
							}
						}
						else
						{
							Console.WriteLine("\n=== Delete Test project node failed: " + nodeName);
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
							elementNode = null;
						}
					}
				  
				}
				#endregion
			  
				// Add Open Wwindow Event Handler
				Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					AutomationElement.RootElement, TreeScope.Descendants, UIEventHandler);

				// Find and click menu item new Project... on Project node
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== FindTest project node and create a new project ===");
					if (!StatUtilities.FindAndClickMenuItemOnThisNode(aeWindow, aeProjectsNode, "New project...", ref sErrorMessage))
					{
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}   
				}

				Thread.Sleep(5000);

				sEventEnd = false;
				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
				Console.WriteLine(" time is :" + mTime.TotalSeconds);
				while (sEventEnd == false && mTime.TotalMinutes <= 5)
				{
					Thread.Sleep(10000);
					mTime = DateTime.Now - mStartTime;
					Console.WriteLine("wait time is :" + mTime.TotalSeconds + "   --------- sEventEnd:" + sEventEnd);
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						AutomationElement.RootElement,
					   UIEventHandler);


				Thread.Sleep(10000);

				// validate result
				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
					if (aeWindow == null)
					{
						sErrorMessage = "MainForm not found";
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						// Check Name row 0 name is value is Test
						TestCheck = ConstCommon.TEST_PASS;
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
				sErrorMessage = ex.Message;
				Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
				Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
			}
			finally
			{
				Thread.Sleep(3000);
				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					   AutomationElement.RootElement,
					  UIEventHandler);
			}
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region CreateDatabase
		public static void CreateDatabase(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			AutomationElement aeTreeView = null;
			AutomationElement aeComputerNameNode = null;
			AutomationElement aeProjectsNode = null;
			try
			{
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

                if (sPaeserConfiguratorConnectToComputerOK == false)
                {
                    sErrorMessage = "parser configurator failed to connect to the computer , this testcase cannot be tested";
                    return;
                }

				#region searching Test project
                AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
				if (aeWindow == null)
				{
					sErrorMessage = "MainForm not found";
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					TestCheck = ConstCommon.TEST_FAIL;
				}
				else
				{
					Console.WriteLine("----- set focus -----");
					aeWindow.SetFocus();
					string treeViewId = "m_TreeView";
					DateTime sTime = DateTime.Now;
					AUIUtilities.WaitUntilElementByIDFound(aeWindow, ref aeTreeView, treeViewId, sTime, 60);
					if (aeTreeView == null)
					{
						sErrorMessage = "aeTreeView not found name";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						TreeWalker walker = TreeWalker.ControlViewWalker;
						AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
						while (elementNode != null)
						{
							Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
							if (elementNode.Current.Name.ToLower().Equals(PCName.ToLower()))
							{
								aeComputerNameNode = elementNode;
								Console.WriteLine("Computer node name found , it is: " + aeComputerNameNode.Current.Name);
								break;
							}
							Thread.Sleep(3000);
							elementNode = walker.GetNextSibling(elementNode);
						}

						if (aeComputerNameNode == null)
						{
							sErrorMessage = "Computer node not found, ";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{ 
							try
							{
								ExpandCollapsePattern ep = (ExpandCollapsePattern)aeComputerNameNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
								ep.Expand();
							}
							catch (Exception ex)
							{
								sErrorMessage = "aeComputerNameNode can not expaned " + aeComputerNameNode.Current.Name+" --- " + ex.Message;
								Console.WriteLine(sErrorMessage);
								TestCheck = ConstCommon.TEST_FAIL;
							}
						}
					}
				}

				// find Projects tree item from computernode
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find Projects node ===");
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					aeProjectsNode = TestTools.AUICommon.WalkEnabledElements(aeComputerNameNode, treeNode, "Projects");
					if (aeProjectsNode == null)
					{
						sErrorMessage = "Projects node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						try
						{
							ExpandCollapsePattern ep = (ExpandCollapsePattern)aeProjectsNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
							ep.Expand();
						}
						catch (Exception ex)
						{
                            sErrorMessage = "aeProjectsNode can not expaned " + aeProjectsNode.Current.Name + " --- " + ex.Message;
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
					}
				}

				Thread.Sleep(5000);
				// find Test Project tree item from Project Node
				AutomationElement aeTestNode = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find" + sCurrentProject +" project node ===");
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					aeTestNode = TestTools.AUICommon.WalkEnabledElements(aeProjectsNode, treeNode, sCurrentProject);
					if (aeTestNode == null)
					{
						sErrorMessage = sCurrentProject+ " projrct Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point TestNodePnt = AUIUtilities.GetElementCenterPoint(aeTestNode);
						Input.MoveToAndRightClick(TestNodePnt);
						Thread.Sleep(2000);
					}
				}
				#endregion

				#region // Find Actions Menuitem --> Find Create database menuitem --> Open database form
				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    AutomationElement aeActions = ProjBasicUI.GetMenuItemFromElement(aeWindow, "Actions", "name", 120, ref sErrorMessage);
					if (aeActions == null)
					{
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Input.MoveTo(aeActions);
						Thread.Sleep(2000);
					}
				}

				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    AutomationElement aeCreateDB = ProjBasicUI.GetMenuItemFromElement(aeWindow, "Create database...", "name", 120, ref sErrorMessage);
					if (aeCreateDB == null)
					{
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						//Input.MoveTo(AUIUtilities.GetElementCenterPoint( aeCreateDB));
						Thread.Sleep(2000);
						Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeCreateDB));
					}
				}
				#endregion

				Thread.Sleep(500);
				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
				#region // work with DatabaseFporm
				AutomationElement aeCreateOrSetDatabaseForm = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					while (aeCreateOrSetDatabaseForm == null && mTime.TotalSeconds <= 120)
					{
						aeCreateOrSetDatabaseForm = 
							StatUtilities.GetElementByIdFromParserConfigurationMainWindow("MainForm", "CreateOrSetDatabaseForm", 120, ref sErrorMessage);
						mTime = DateTime.Now - mStartTime;
						Thread.Sleep(2000);
					}
					
					if (aeCreateOrSetDatabaseForm == null)
					{
						sErrorMessage = "Create New Statistic database form not found after 2 min";
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					
				}

				mStartTime = DateTime.Now;
				mTime = DateTime.Now - mStartTime;
				AutomationElement aeBtnBrowse = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					while (aeBtnBrowse == null && mTime.TotalSeconds <= 120)
					{
						aeBtnBrowse = AUIUtilities.FindElementByID("m_BtnBrowseSqlServerInstance", aeCreateOrSetDatabaseForm);
						mTime = DateTime.Now - mStartTime;
						Thread.Sleep(2000);
					}

					if (aeBtnBrowse == null)
					{
						sErrorMessage = "aeBtnBrowse not found after 2 min";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Input.MoveToAndClick(aeBtnBrowse);
						Thread.Sleep(1000);
						Input.MoveToAndClick(aeBtnBrowse);
					}
				}
				#endregion

				// work with Remote Sql 
				AutomationElement aeBRemoteSqlServerDbBrowserForm = null;
				string RemoteSqlServerDbBrowserFormID = "RemoteSqlServerDbBrowserForm";
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeBRemoteSqlServerDbBrowserForm =
						StatUtilities.GetElementByIdFromParserConfigurationMainWindow("MainForm", RemoteSqlServerDbBrowserFormID, 120, ref sErrorMessage);
					if (aeBRemoteSqlServerDbBrowserForm == null)
					{
						sErrorMessage = "aeBRemoteSqlServerDbBrowserForm form not found after 2 min";
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
				}

				string InstanceRefreshButtonId = "m_BtnRefreshInstance";  // "ControlType.Button"
				string InstanceComboBoxId = "m_ComboBoxInstance";  // "ControlType.Combo
				string SelectButtonId = "m_BtnSelect";  // "ControlType.Button"
				string CancelButtonId = "m_BtnCancel";  // "ControlType.Button"

				// (1) Click Instance Refresh button first 
				AutomationElement aeInstanceRefreshButton = null;
				AutomationElement aeInstanceComboBox = null;
				AutomationElement aeSelectButton = null;
				AutomationElement aeCancelButton = null;
				mStartTime = DateTime.Now;
				mTime = DateTime.Now - mStartTime;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeInstanceRefreshButton = AUIUtilities.FindElementByID(InstanceRefreshButtonId, aeBRemoteSqlServerDbBrowserForm);
					while (aeInstanceRefreshButton == null && mTime.TotalSeconds <= 120)
					{
						Console.WriteLine("aeInstanceRefreshButton is not found yet ....");
						aeInstanceRefreshButton = AUIUtilities.FindElementByID(InstanceRefreshButtonId, aeBRemoteSqlServerDbBrowserForm);
						Thread.Sleep(10000);
						mTime = DateTime.Now - mStartTime;
					}

					if (aeInstanceRefreshButton == null)
					{
						sErrorMessage = "aeInstanceRefreshButton not found after 2 minutes";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Console.WriteLine("aeInstanceRefreshButton is found and click ....");
						Thread.Sleep(1000);
						Input.MoveToAndClick(aeInstanceRefreshButton);
						Thread.Sleep(4000);
					}
				}

				// Wait until ComboBoxInstance received Server Instances
				Console.WriteLine("Wait until ComboBoxInstance received Server Instances ....");
				Condition cCombo = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem);
				mStartTime = DateTime.Now;
				mTime = DateTime.Now - mStartTime;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					int k = 0;
					while (aeInstanceComboBox == null && mTime.TotalSeconds <= 300)
					{
						Console.WriteLine("aeInstanceComboBox is not found yet ...."+k);
						k++;
						aeInstanceComboBox = AUIUtilities.FindElementByID(InstanceComboBoxId, aeBRemoteSqlServerDbBrowserForm);
						if (aeInstanceComboBox == null)
						{
							Console.WriteLine("aeInstanceComboBox is still not found yet ....");
							Thread.Sleep(10000);
						}
						else
						{
							//Get the List child control inside the combo box
							AutomationElement comboboxList = aeInstanceComboBox.FindFirst(TreeScope.Children,
								  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List));

							//Get the all the listitems in List control
							AutomationElementCollection comboboxItem =
							   comboboxList.FindAll(TreeScope.Children,
								new PropertyCondition(AutomationElement.ControlTypeProperty,ControlType.ListItem));

							Console.WriteLine("aeInstanceComboBox is found, with item count ...." + comboboxItem.Count);
							if (comboboxItem.Count == 0)
							{
								Console.WriteLine("aeInstanceComboBox is empty ...." + comboboxItem.Count);
								Thread.Sleep(2000);
								aeInstanceComboBox = null;
							}
							else
							{
								SelectionPattern selectPattern =
								   aeInstanceComboBox.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

								AutomationElement item = AUIUtilities.FindElementByName(PCName, aeInstanceComboBox);
									//= AUIUtilities.FindElementByName("PCC20090125-AM", aeInstanceComboBox);
								if (item != null)
								{
									Console.WriteLine("Server instance item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
									Thread.Sleep(5000);
									Console.WriteLine("Select this Server instance item at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
									SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
									itemPattern.Select();
									#region // CLICK SELECT BUTTON
									aeSelectButton = AUIUtilities.FindElementByID(SelectButtonId, aeBRemoteSqlServerDbBrowserForm);
									mStartTime = DateTime.Now;
									mTime = DateTime.Now - mStartTime;
									while (aeSelectButton.Current.IsEnabled == false && mTime.TotalSeconds <= 300)
									{
										Console.WriteLine("aeSelectButton is not enabled yet ....");
										aeSelectButton = AUIUtilities.FindElementByID(SelectButtonId, aeBRemoteSqlServerDbBrowserForm);
										Thread.Sleep(10000);
										mTime = DateTime.Now - mStartTime;
									}

									if (aeSelectButton != null)
									{
										Thread.Sleep(1000);
										Input.MoveToAndClick(aeSelectButton);
										Thread.Sleep(1000);

									}
									else
									{
										sErrorMessage = "aeSelectButtonn not found name";
										Console.WriteLine(sErrorMessage);
										TestCheck = ConstCommon.TEST_FAIL;
									}
									#endregion
								}
								else
								{
									Console.WriteLine("Finding server instance item failed");
									sErrorMessage = "Finding server instance item failed";
									#region // CLICK Cancel BUTTON
									mStartTime = DateTime.Now;
									mTime = DateTime.Now - mStartTime;
									while (aeCancelButton == null && mTime.TotalSeconds <= 300)
									{
										Console.WriteLine("aeCancelButton is not found yet ....");
										aeCancelButton = AUIUtilities.FindElementByID(CancelButtonId, aeBRemoteSqlServerDbBrowserForm);
										Thread.Sleep(10000);
										mTime = DateTime.Now - mStartTime;
									}

									if (aeCancelButton != null)
									{
										Console.WriteLine("aeCancelButton is found yet ...." + aeCancelButton.Current.AutomationId);
										Point CancelPtn = AUIUtilities.GetElementCenterPoint(aeCancelButton);
										Thread.Sleep(1000);
										Input.MoveToAndClick(CancelPtn);
										Thread.Sleep(1000);
										Input.MoveToAndClick(CancelPtn);
										Thread.Sleep(1000);
										Input.MoveToAndClick(CancelPtn);
										Thread.Sleep(1000);
										Input.MoveToAndClick(CancelPtn);

									}
									else
									{
										sErrorMessage = "aeCancelButton not found";
										Console.WriteLine(sErrorMessage);
										TestCheck = ConstCommon.TEST_FAIL;
									}
									#endregion
								}
							}
						}
						mTime = DateTime.Now - mStartTime;
					}
				}
				
				//Valated database form instance is slected
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeCreateOrSetDatabaseForm =
						StatUtilities.GetElementByIdFromParserConfigurationMainWindow("MainForm", "CreateOrSetDatabaseForm", 120, ref sErrorMessage);
					if (aeCreateOrSetDatabaseForm == null)
					{
						sErrorMessage = "Create New Statistic database form not found";
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
				}

				mStartTime = DateTime.Now;
				mTime = DateTime.Now - mStartTime;
				AutomationElement aeSqlServerInstance = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					while (aeSqlServerInstance == null && mTime.TotalSeconds <= 300)
					{
						Console.WriteLine("aeSqlServerInstance is not found yet ....");
						aeSqlServerInstance = AUIUtilities.FindElementByID("m_TextBoxSqlServerInstance", aeCreateOrSetDatabaseForm);
						if (aeSqlServerInstance == null)
						{
							sErrorMessage = "aeSqlServerInstance form not found";
							Console.WriteLine(sErrorMessage);
							Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
							Thread.Sleep(5000);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							Console.WriteLine("aeSqlServerInstance is found, now check server instance ....");
							ValuePattern valuePattern =
							   aeSqlServerInstance.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;

							string value = valuePattern.Current.Value;
							Console.WriteLine("aeSqlServerInstance is found with value :"+value);
							if (value.ToLower().Equals(PCName.ToLower()))
							{
								Console.WriteLine("Server instance is selected: " + System.DateTime.Now.ToString("HH:mm:ss"));
								Thread.Sleep(5000);
								#region // CLICK OK BUTTON
								AutomationElement aeOKButton = AUIUtilities.FindElementByID("m_BtnOK", aeCreateOrSetDatabaseForm);
								mStartTime = DateTime.Now;
								mTime = DateTime.Now - mStartTime;
								while (aeOKButton == null && mTime.TotalSeconds <= 300)
								{
									Console.WriteLine(" aeOKButton is not found yet ....");
									aeOKButton = AUIUtilities.FindElementByID("m_BtnOK", aeCreateOrSetDatabaseForm);
									Thread.Sleep(5000);
									mTime = DateTime.Now - mStartTime;
								}

								if (aeOKButton != null)
								{
									Thread.Sleep(1000);
									Input.MoveToAndClick(aeOKButton);
									Thread.Sleep(1000);

								}
								else
								{
									sErrorMessage = "aeOKButton not found";
									Console.WriteLine(sErrorMessage);
									TestCheck = ConstCommon.TEST_FAIL;
								}
								#endregion
							}
							else
							{
								Console.WriteLine("select server instance item failed");
								sErrorMessage = "select server instance item failed";
								TestCheck = ConstCommon.TEST_FAIL;
								#region // CLICK Cancel BUTTON
								mStartTime = DateTime.Now;
								mTime = DateTime.Now - mStartTime;
								while (aeCancelButton == null && mTime.TotalSeconds <= 300)
								{
									Console.WriteLine("aeCancelButton is not found yet ....");
									aeCancelButton = AUIUtilities.FindElementByID(CancelButtonId, aeCreateOrSetDatabaseForm);
									Thread.Sleep(10000);
									mTime = DateTime.Now - mStartTime;
								}

								if (aeCancelButton != null)
								{
									Thread.Sleep(1000);
									Input.MoveToAndClick(aeCancelButton);
									Thread.Sleep(1000);

								}
								else
								{
									sErrorMessage = "aeCancelButton not found";
									Console.WriteLine(sErrorMessage);
									TestCheck = ConstCommon.TEST_FAIL;
								}
								#endregion
							}
						}
						mTime = DateTime.Now - mStartTime;
					}
				}

                Console.WriteLine("checkoutput screen ..................");
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.CheckParserConfiguratorOutput("MainForm", ref sErrorMessage) == false)
                    {
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
					result = ConstCommon.TEST_PASS;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
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
			}
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region SetDatabase
		public static void SetDatabase(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			AutomationElement aeTreeView = null;
			AutomationElement aeComputerNameNode = null;
			AutomationElement aeProjectsNode = null;
			try
			{
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

                if (sPaeserConfiguratorConnectToComputerOK == false)
                {
                    sErrorMessage = "parser configurator failed to connect to the computer , this testcase cannot be tested";
                    return;
                }

				#region searching Test project
                AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
				if (aeWindow == null)
				{
					sErrorMessage = "MainForm not found";
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					TestCheck = ConstCommon.TEST_FAIL;
				}
				else
				{
					Console.WriteLine("----- set focus -----");
					aeWindow.SetFocus();
					string treeViewId = "m_TreeView";
					DateTime sTime = DateTime.Now;
					AUIUtilities.WaitUntilElementByIDFound(aeWindow, ref aeTreeView, treeViewId, sTime, 60);
					if (aeTreeView == null)
					{
						sErrorMessage = "aeTreeView not found name";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						TreeWalker walker = TreeWalker.ControlViewWalker;
						AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
						while (elementNode != null)
						{
							Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
							if (elementNode.Current.Name.ToLower().Equals(PCName.ToLower()))
							{
								aeComputerNameNode = elementNode;
								Console.WriteLine("Computer node name found , it is: " + aeComputerNameNode.Current.Name);
								break;
							}
							Thread.Sleep(3000);
							elementNode = walker.GetNextSibling(elementNode);
						}

						if (aeComputerNameNode == null)
						{
							sErrorMessage = "Computer node not found, ";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							try
							{
								ExpandCollapsePattern ep = (ExpandCollapsePattern)aeComputerNameNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
								ep.Expand();
							}
							catch (Exception ex)
							{
                                sErrorMessage = "aeComputerNameNode can not expaned " + aeComputerNameNode.Current.Name + " --- " + ex.Message;
								Console.WriteLine(sErrorMessage);
								TestCheck = ConstCommon.TEST_FAIL;
							}
						}
					}
				}

				// find Projects tree item from computernode
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find Projects node ===");
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					aeProjectsNode = TestTools.AUICommon.WalkEnabledElements(aeComputerNameNode, treeNode, "Projects");
					if (aeProjectsNode == null)
					{
						sErrorMessage = "Projects node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						try
						{
							ExpandCollapsePattern ep = (ExpandCollapsePattern)aeProjectsNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
							ep.Expand();
						}
						catch (Exception ex)
						{
                            sErrorMessage = "aeProjectsNode can not expaned " + aeProjectsNode.Current.Name + " --- " + ex.Message;
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
					}
				}

				Thread.Sleep(5000);
				// find Test Project tree item from Project Node
				AutomationElement aeTestNode = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find" + sCurrentProject + " project node ===");
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					aeTestNode = TestTools.AUICommon.WalkEnabledElements(aeProjectsNode, treeNode, sCurrentProject);
					if (aeTestNode == null)
					{
						sErrorMessage = sCurrentProject + " projrct Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point TestNodePnt = AUIUtilities.GetElementCenterPoint(aeTestNode);
						Input.MoveToAndRightClick(TestNodePnt);
						Thread.Sleep(2000);
					}
				}
				#endregion

				#region // Find Actions Menuitem --> Find Create database menuitem --> Open database form
				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    AutomationElement aeActions = ProjBasicUI.GetMenuItemFromElement(aeWindow, "Actions", "name", 120, ref sErrorMessage);
					if (aeActions == null)
					{
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Input.MoveTo(aeActions);
						Thread.Sleep(2000);
					}
				}

				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    AutomationElement aeCreateDB = ProjBasicUI.GetMenuItemFromElement(aeWindow, "Set database...", "name", 120, ref sErrorMessage);
					if (aeCreateDB == null)
					{
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						//Input.MoveTo(AUIUtilities.GetElementCenterPoint( aeCreateDB));
						Thread.Sleep(2000);
						Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeCreateDB));
					}
				}
				#endregion

				Thread.Sleep(5000);
				#region // work with DatabaseFporm
				AutomationElement aeCreateOrSetDatabaseForm = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeCreateOrSetDatabaseForm =
						StatUtilities.GetElementByIdFromParserConfigurationMainWindow("MainForm", "CreateOrSetDatabaseForm", 120, ref sErrorMessage);
					if (aeCreateOrSetDatabaseForm == null)
					{
						sErrorMessage = "Create New Statistic database form not found";
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
				AutomationElement aeBtnOK = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					while (aeBtnOK == null && mTime.TotalSeconds <= 120)
					{
						aeBtnOK = AUIUtilities.FindElementByID("m_BtnOK", aeCreateOrSetDatabaseForm);
						mTime = DateTime.Now - mStartTime;
					}

					if (aeBtnOK == null)
					{
						sErrorMessage = "aeBtnOK not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point pt = AUIUtilities.GetElementCenterPoint(aeBtnOK);
						Thread.Sleep(2000);
						Input.MoveToAndClick(pt);
						Thread.Sleep(2000);
					}
				}
				#endregion

                Console.WriteLine("checkoutput screen ..................");
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.CheckParserConfiguratorOutput("MainForm", ref sErrorMessage) == false)
                    {
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
					result = ConstCommon.TEST_PASS;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
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
			}
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region CreateEtriccLayout
		public static void CreateEtriccLayout(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			AutomationElement aeTreeView = null;
			AutomationElement aeComputerNameNode = null;
			AutomationElement aeProjectsNode = null;
			try
			{
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

                if (sPaeserConfiguratorConnectToComputerOK == false)
                {
                    sErrorMessage = "parser configurator failed to connect to the computer , this testcase cannot be tested";
                    return;
                }

				#region searching Test project
                AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
				if (aeWindow == null)
				{
					sErrorMessage = "MainForm not found";
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					TestCheck = ConstCommon.TEST_FAIL;
				}
				else
				{
					Console.WriteLine("----- set focus -----");
					aeWindow.SetFocus();
					string treeViewId = "m_TreeView";
					DateTime sTime = DateTime.Now;
					AUIUtilities.WaitUntilElementByIDFound(aeWindow, ref aeTreeView, treeViewId, sTime, 60);
					if (aeTreeView == null)
					{
						sErrorMessage = "aeTreeView not found name";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						TreeWalker walker = TreeWalker.ControlViewWalker;
						AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
						while (elementNode != null)
						{
							Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
							if (elementNode.Current.Name.ToLower().Equals(PCName.ToLower()))
							{
								aeComputerNameNode = elementNode;
								Console.WriteLine("Computer node name found , it is: " + aeComputerNameNode.Current.Name);
								break;
							}
							Thread.Sleep(3000);
							elementNode = walker.GetNextSibling(elementNode);
						}

						if (aeComputerNameNode == null)
						{
							sErrorMessage = "Computer node not found, ";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							try
							{
								ExpandCollapsePattern ep = (ExpandCollapsePattern)aeComputerNameNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
								ep.Expand();
							}
							catch (Exception ex)
							{
                                sErrorMessage = "aeComputerNameNode can not expaned " + aeComputerNameNode.Current.Name + " --- " + ex.Message;
								Console.WriteLine(sErrorMessage);
								TestCheck = ConstCommon.TEST_FAIL;
							}
						}
					}
				}

				// find Projects tree item from computernode
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find Projects node ===");
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					aeProjectsNode = TestTools.AUICommon.WalkEnabledElements(aeComputerNameNode, treeNode, "Projects");
					if (aeProjectsNode == null)
					{
						sErrorMessage = "Projects node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						try
						{
							ExpandCollapsePattern ep = (ExpandCollapsePattern)aeProjectsNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
							ep.Expand();
						}
						catch (Exception ex)
						{
                            sErrorMessage = "aeProjectsNode can not expaned " + aeProjectsNode.Current.Name + " --- " + ex.Message;
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
					}
				}

				Thread.Sleep(5000);
				// find Test Project tree item from Project Node
				AutomationElement aeTestNode = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find "+ sCurrentProject + " project node ===");
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					aeTestNode = TestTools.AUICommon.WalkEnabledElements(aeProjectsNode, treeNode,  sCurrentProject);
					if (aeTestNode == null)
					{
						sErrorMessage = sCurrentProject + " projrct Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point TestNodePnt = AUIUtilities.GetElementCenterPoint(aeTestNode);
						Input.MoveToAndRightClick(TestNodePnt);
						Thread.Sleep(2000);
					}
				}
				#endregion

				#region // Find Actions Menuitem --> Find Create Etricc layout... menuitem --> Open window form
				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    AutomationElement aeActions = ProjBasicUI.GetMenuItemFromElement(aeWindow, "Actions", "name", 120, ref sErrorMessage);
					if (aeActions == null)
					{
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Input.MoveTo(aeActions);
						Thread.Sleep(2000);
					}
				}

				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    AutomationElement aeCreateEtricc = ProjBasicUI.GetMenuItemFromElement(aeWindow, "Create Etricc layout...", "name", 120, ref sErrorMessage);
					if (aeCreateEtricc == null)
					{
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point pt = AUIUtilities.GetElementCenterPoint(aeCreateEtricc);
						Input.MoveTo(pt);
						Thread.Sleep(2000);
						Input.MoveToAndClick(pt);
						Thread.Sleep(2000);
					}
				}

				AutomationElement aeSecondWindow = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
					if (aeWindow == null)
					{
						sErrorMessage = "MainForm not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						aeSecondWindow = TestTools.AUIUtilities.FindElementByName(aeWindow.Current.Name, aeWindow);
						if (aeSecondWindow == null)
						{
							sErrorMessage = "aeSecondWindow not found";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
					}
				}
				#endregion

				Thread.Sleep(5000);
				#region // work with second window
				AutomationElement aeYesButton = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeYesButton = AUIUtilities.FindElementByName("Yes", aeSecondWindow);
					if (aeYesButton == null)
					{
						sErrorMessage = "aeYesButton not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point pt = AUIUtilities.GetElementCenterPoint(aeYesButton);
						Thread.Sleep(2000);
						Input.MoveTo(pt);
						Thread.Sleep(2000);
						Input.ClickAtPoint(pt);
					
					}
				}
				#endregion

                Console.WriteLine("checkoutput screen ..................");
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.CheckParserConfiguratorOutput("MainForm", ref sErrorMessage) == false)
                    {
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
					result = ConstCommon.TEST_PASS;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
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
			}

		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region ParseTestProject
		public static void ParseTestProject(string testname, AutomationElement root, out int result)
		{
			// delete backup folder
			// check stat data
			// after activate check zip file
			Console.WriteLine("=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;

			TestCheck = ConstCommon.TEST_PASS;

			AutomationElement aeTreeView = null;
			AutomationElement aeComputerNameNode = null;
			AutomationElement aeProjectsNode = null;
			try
			{
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

                if (sPaeserConfiguratorConnectToComputerOK == false)
                {
                    sErrorMessage = "parser configurator failed to connect to the computer , this testcase cannot be tested";
                    return;
                }

				#region searching Test project
                AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
				if (aeWindow == null)
				{
					sErrorMessage = "MainForm not found";
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					TestCheck = ConstCommon.TEST_FAIL;
				}
				else
				{
					Console.WriteLine("----- set focus -----");
					aeWindow.SetFocus();
					string treeViewId = "m_TreeView";
					DateTime sTime = DateTime.Now;
					AUIUtilities.WaitUntilElementByIDFound(aeWindow, ref aeTreeView, treeViewId, sTime, 60);
					if (aeTreeView == null)
					{
						sErrorMessage = "aeTreeView not found name";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						TreeWalker walker = TreeWalker.ControlViewWalker;
						AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
						while (elementNode != null)
						{
							Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
							if (elementNode.Current.Name.ToLower().Equals(PCName.ToLower()))
							{
								aeComputerNameNode = elementNode;
								Console.WriteLine("Computer node name found , it is: " + aeComputerNameNode.Current.Name);
								break;
							}
							Thread.Sleep(3000);
							elementNode = walker.GetNextSibling(elementNode);
						}

						if (aeComputerNameNode == null)
						{
							sErrorMessage = "Computer node not found, ";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							try
							{
								ExpandCollapsePattern ep = (ExpandCollapsePattern)aeComputerNameNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
								ep.Expand();
							}
							catch (Exception ex)
							{
                                sErrorMessage = "aeComputerNameNode can not expaned " + aeComputerNameNode.Current.Name + " --- " + ex.Message;
								Console.WriteLine(sErrorMessage);
								TestCheck = ConstCommon.TEST_FAIL;
							}
						}
					}
				}

				// find Projects tree item from computernode
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find Projects node ===");
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					aeProjectsNode = TestTools.AUICommon.WalkEnabledElements(aeComputerNameNode, treeNode, "Projects");
					if (aeProjectsNode == null)
					{
						sErrorMessage = "Projects node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						try
						{
							ExpandCollapsePattern ep = (ExpandCollapsePattern)aeProjectsNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
							ep.Expand();
						}
						catch (Exception ex)
						{
                            sErrorMessage = "aeProjectsNode can not expaned " + aeProjectsNode.Current.Name + " --- " + ex.Message;
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
					}
				}

				Thread.Sleep(5000);
				// find Test Project tree item from Project Node
				AutomationElement aeTestNode = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find"+ sCurrentProject + " project node ===");
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					aeTestNode = TestTools.AUICommon.WalkEnabledElements(aeProjectsNode, treeNode, sCurrentProject);
					if (aeTestNode == null)
					{
						sErrorMessage = sCurrentProject + " projrct Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point TestNodePnt = AUIUtilities.GetElementCenterPoint(aeTestNode);
						Input.MoveToAndRightClick(TestNodePnt);
						Thread.Sleep(2000);
					}
				}
				#endregion

				#region // Find Actions Menuitem --> Find Create Etricc layout... menuitem --> Open window form
				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    AutomationElement aeActivate = ProjBasicUI.GetMenuItemFromElement(aeWindow, "Activate", "name", 120, ref sErrorMessage);
					if (aeActivate == null)
					{
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point pt = AUIUtilities.GetElementCenterPoint(aeActivate);
						Input.MoveTo(pt);
						Thread.Sleep(2000);
						Input.ClickAtPoint(pt);
					}
				}
				#endregion

				Thread.Sleep(60000);
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					// wip check backup folder
					#region // validate result
					/*
					Console.WriteLine("check output ....");
					aeWindow = StatUtilities.GetParserConfigurationMainWindow();
					if (aeWindow == null)
					{
						sErrorMessage = "MainForm not found";
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						string textBoxPanelId = "m_TextBoxContainerPanel";
						AutomationElement aeTextBoxPanel = null;
						DateTime sTime = DateTime.Now;
						AUIUtilities.WaitUntilElementByIDFound(aeWindow, ref aeTextBoxPanel, textBoxPanelId, sTime, 60);

						if (aeTextBoxPanel == null)
						{
							sErrorMessage = "aeTextBoxPanel not found name";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							AutomationElement aeDocument = null;
							AutomationElementCollection aeAllDocuments = null;
							// find ducument text
							System.Windows.Automation.Condition cDocs = new PropertyCondition(
								AutomationElement.ControlTypeProperty, ControlType.Document);

							int k = 0;
							DateTime mStartTime = DateTime.Now;
							TimeSpan mTime = DateTime.Now - mStartTime;
							while (aeDocument == null && mTime.TotalSeconds <= 120)
							{
								Console.WriteLine("aeDocument[k]=" + k);
								k++;
								aeAllDocuments = aeTextBoxPanel.FindAll(TreeScope.Children, cDocs);
								Thread.Sleep(3000);
								for (int i = 0; i < aeAllDocuments.Count; i++)
								{
									Console.WriteLine("aeDocument[" + i + "] name length=" + aeAllDocuments[i].Current.Name.Length);
									if (aeAllDocuments[i].Current.Name.Length > 20)
									{
										aeDocument = aeAllDocuments[i];
										Console.WriteLine("aeDocument[" + i + "]=" + aeDocument.Current.Name);
										Epia3Common.WriteTestLogFail(slogFilePath, "aeDocument found", sOnlyUITest);
										break;
									}
								}
								mTime = DateTime.Now - mStartTime;
							}

							if (aeDocument == null)
							{
								sErrorMessage = "aeDocument not found name";
								Console.WriteLine(sErrorMessage);
								TestCheck = ConstCommon.TEST_FAIL;
							}
							else
							{
								if (aeDocument.Current.Name.IndexOf("=== Create of Etricc layout finished successfully") > 0)
								{
									TestCheck = ConstCommon.TEST_PASS;

								}
								else
								{
									sErrorMessage = "=== Create of Etricc layout finished with error(s)=========:";
									Console.WriteLine(sErrorMessage + aeDocument.Current.Name);
									TestCheck = ConstCommon.TEST_FAIL;
								}
							}
						}
					}
					 *  */
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
					ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_ETRICC_STATISTICS_PARSERCONFIGURATOR );
					result = ConstCommon.TEST_PASS;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
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
			}
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region ModifyConfigurationFiles
		public static void ModifyConfigurationFiles(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_PASS;
			TestCheck = ConstCommon.TEST_PASS;
			 /* 
	replace <DataSource>localhost</DataSource>   localhost --> epiatestpc2
	<InitialCatalog>EtriccStatistics_PROJECTNAME</InitialCatalog>
	EtriccStatistics_PROJECTNAME --> EtriccStatistics_Test

<ReportServerUrl>http://localhost/ReportServer</ReportServerUrl>

localhost --> epiatestpc2
			*/

			//xPathNav.ReplaceSelf("<ReportServerUrl>http://" + PCName + "/ReportServer</ReportServerUrl>");
			// should be <ReportServerUrl>http://ETRICCSTATAUTOTEST1.Teamsystems.egemin.be/ReportServer</ReportServerUrl>
			// getFQDN()
			try
			{
				#region // Edit C:\Program Files\Egemin\Epia Server\Data\SqlRptServices\Etricc.Default.xml
				var etriccDefaultXml = new XmlDocument();
				string path = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Epia Server\Data\SqlRptServices";
				etriccDefaultXml.Load(System.IO.Path.Combine(path, "Etricc.Default.xml"));
				var xPathNav = etriccDefaultXml.CreateNavigator();
				Console.WriteLine("=== CreateNavigator " + xPathNav.LocalName);
				Thread.Sleep(2000);
				xPathNav.MoveToFirstChild();
				Console.WriteLine("=== MoveToFirstChild " + xPathNav.LocalName);
				Thread.Sleep(2000);
				xPathNav.MoveToFirstChild();
				Console.WriteLine("=== MoveToFirstChild " + xPathNav.LocalName);
				Thread.Sleep(2000);
				while (xPathNav.MoveToNext() == true)
				{
					Console.WriteLine("=== while MoveToFirstChild " + xPathNav.LocalName);
					Thread.Sleep(2000);
					if (xPathNav.LocalName.StartsWith("DataSource"))
					{
						Console.WriteLine("=== replace " + xPathNav.LocalName);
						xPathNav.ReplaceSelf("<DataSource>"+ PCName+"</DataSource>");
					}
					else if (xPathNav.LocalName.StartsWith("InitialCatalog"))
						xPathNav.ReplaceSelf("<InitialCatalog>EtriccStatistics_"+ sCurrentProject +"</InitialCatalog>");
					else if (xPathNav.LocalName.StartsWith("ReportServerUrl"))
						xPathNav.ReplaceSelf("<ReportServerUrl>http://"+ StatUtilities.getFQDN() +"/ReportServer</ReportServerUrl>");
					//else if (xPathNav.LocalName.StartsWith("ReportTimeoutInMs")) // end node
					 //   break;
				}
				etriccDefaultXml.Save(System.IO.Path.Combine(path, "Etricc.Default.xml"));
				#endregion

				#region // Edit C:\Program Files\Egemin\Epia Server\Egemin.Epia.Server.exe.config
				// <systemOverviewManagerConfiguration dataSource="ETRICCSTATAUTOT" initialCatalog="EtriccStatistics_Test"/>
				var epiaServerExeConfig = new XmlDocument();
				string EpiaPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Epia Server";
				epiaServerExeConfig.Load(System.IO.Path.Combine(EpiaPath, "Egemin.Epia.Server.exe.config"));
				var epiaPathNav = epiaServerExeConfig.CreateNavigator();
				Console.WriteLine("=== CreateNavigator " + epiaPathNav.LocalName);
				Thread.Sleep(2000);
				epiaPathNav.MoveToFirstChild();
				Console.WriteLine("=== MoveToFirstChild " + epiaPathNav.LocalName);
				Thread.Sleep(2000);
				epiaPathNav.MoveToFirstChild();
				Console.WriteLine("=== MoveToFirstChild " + epiaPathNav.LocalName);
				Thread.Sleep(2000);
				while (epiaPathNav.MoveToNext() == true)
				{
					Console.WriteLine("=== while MoveToFirstChild " + epiaPathNav.LocalName);
					Thread.Sleep(2000);
					if (epiaPathNav.LocalName.StartsWith("systemOverviewManagerConfiguration"))
					{
						Console.WriteLine("=== replace " + epiaPathNav.LocalName);
						string replacestr = "<systemOverviewManagerConfiguration dataSource=\""+ PCName +"\" initialCatalog=\"EtriccStatistics_" +sCurrentProject +"\"/>";
						Console.WriteLine("=== replace " + replacestr);
						epiaPathNav.ReplaceSelf("<systemOverviewManagerConfiguration dataSource=\"" + PCName + "\" initialCatalog=\"EtriccStatistics_" + sCurrentProject + "\"/>");
					}
				}
				epiaServerExeConfig.Save(System.IO.Path.Combine(EpiaPath, "Egemin.Epia.Server.exe.config"));
				#endregion

				if (TestCheck == ConstCommon.TEST_FAIL)
				{
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					result = ConstCommon.TEST_FAIL;
				}

				if (TestCheck == ConstCommon.TEST_PASS)
				{
					if (result == ConstCommon.TEST_PASS)
					{
						Console.WriteLine("\nTest scenario update configuration files: Pass");
						Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
						result = ConstCommon.TEST_PASS;
					}
					else
					{
						Console.WriteLine("\nTest scenario  update configuration files: *FAIL*");
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
			   
			}
		}
		#endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ModifyConfigurationFilesLinqToXML
        public static void ModifyConfigurationFilesLinqToXML(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_PASS;
            TestCheck = ConstCommon.TEST_PASS;
            /* 
   replace <DataSource>localhost</DataSource>   localhost --> epiatestpc2
   <InitialCatalog>EtriccStatistics_PROJECTNAME</InitialCatalog>
   EtriccStatistics_PROJECTNAME --> EtriccStatistics_Test

<ReportServerUrl>http://localhost/ReportServer</ReportServerUrl>

localhost --> epiatestpc2
           */

            //xPathNav.ReplaceSelf("<ReportServerUrl>http://" + PCName + "/ReportServer</ReportServerUrl>");
            // should be <ReportServerUrl>http://ETRICCSTATAUTOTEST1.Teamsystems.egemin.be/ReportServer</ReportServerUrl>
            // getFQDN()
            try
            {
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

                if (sPaeserConfiguratorConnectToComputerOK == false)
                {
                    sErrorMessage = "parser configurator failed to connect to the computer , this testcase cannot be tested";
                    return;
                }

                #region // Edit C:\Program Files\Egemin\Epia Server\Data\SqlRptServices\Etricc.Default.xml
                string path = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Epia Server\Data\SqlRptServices";
                System.Xml.Linq.XDocument etriccDefaultXml = System.Xml.Linq.XDocument.Load(System.IO.Path.Combine(path, "Etricc.Default.xml"));
              
                XElement root1 = etriccDefaultXml.Root;  

                // Update Lisa's job to florist  
                /*root1.Elements("DataSource")..Where(e => e.Element("Name").Value.Equals("Lisa")).Select(e => e.Element("Job")).Single().SetValue("Florist");  
                  //update the comment  
                document.Nodes().OfType<XComment>().Single().Value = "My new, updated comment!";  
                  document.Save("People.xml"); 
                */
                  // Find the colors for a given make.
                /*var dataSourceInfo = from car in etriccDefaultXml.Descendants("Car")
                           where (string)car.Element("Make") == make
                           select car.Element("Color").Value;

            // Build a string representing each color. 
            string data = string.Empty;
            foreach (var item in makeInfo.Distinct())
            {
                data += string.Format("- {0}\n", item);
            }*/


                //etriccDefaultXml.Load(System.IO.Path.Combine(path, "Etricc.Default.xml"));
                /*var xPathNav = etriccDefaultXml.CreateNavigator();
                Console.WriteLine("=== CreateNavigator " + xPathNav.LocalName);
                Thread.Sleep(2000);
                xPathNav.MoveToFirstChild();
                Console.WriteLine("=== MoveToFirstChild " + xPathNav.LocalName);
                Thread.Sleep(2000);
                xPathNav.MoveToFirstChild();
                Console.WriteLine("=== MoveToFirstChild " + xPathNav.LocalName);
                Thread.Sleep(2000);
                while (xPathNav.MoveToNext() == true)
                {
                    Console.WriteLine("=== while MoveToFirstChild " + xPathNav.LocalName);
                    Thread.Sleep(2000);
                    if (xPathNav.LocalName.StartsWith("DataSource"))
                    {
                        Console.WriteLine("=== replace " + xPathNav.LocalName);
                        xPathNav.ReplaceSelf("<DataSource>" + PCName + "</DataSource>");
                    }
                    else if (xPathNav.LocalName.StartsWith("InitialCatalog"))
                        xPathNav.ReplaceSelf("<InitialCatalog>EtriccStatistics_" + sCurrentProject + "</InitialCatalog>");
                    else if (xPathNav.LocalName.StartsWith("ReportServerUrl"))
                        xPathNav.ReplaceSelf("<ReportServerUrl>http://" + StatUtilities.getFQDN() + "/ReportServer</ReportServerUrl>");
                    //else if (xPathNav.LocalName.StartsWith("ReportTimeoutInMs")) // end node
                    //   break;
                }
                etriccDefaultXml.Save(System.IO.Path.Combine(path, "Etricc.Default.xml"));*/
                #endregion

                #region // Edit C:\Program Files\Egemin\Epia Server\Egemin.Epia.Server.exe.config
                // <systemOverviewManagerConfiguration dataSource="ETRICCSTATAUTOT" initialCatalog="EtriccStatistics_Test"/>
                var epiaServerExeConfig = new XmlDocument();
                string EpiaPath = @"C:\" + OSVersionInfoClass.ProgramFilesx86FolderName() + @"\Egemin\Epia Server";
                epiaServerExeConfig.Load(System.IO.Path.Combine(EpiaPath, "Egemin.Epia.Server.exe.config"));
                var epiaPathNav = epiaServerExeConfig.CreateNavigator();
                Console.WriteLine("=== CreateNavigator " + epiaPathNav.LocalName);
                Thread.Sleep(2000);
                epiaPathNav.MoveToFirstChild();
                Console.WriteLine("=== MoveToFirstChild " + epiaPathNav.LocalName);
                Thread.Sleep(2000);
                epiaPathNav.MoveToFirstChild();
                Console.WriteLine("=== MoveToFirstChild " + epiaPathNav.LocalName);
                Thread.Sleep(2000);
                while (epiaPathNav.MoveToNext() == true)
                {
                    Console.WriteLine("=== while MoveToFirstChild " + epiaPathNav.LocalName);
                    Thread.Sleep(2000);
                    if (epiaPathNav.LocalName.StartsWith("systemOverviewManagerConfiguration"))
                    {
                        Console.WriteLine("=== replace " + epiaPathNav.LocalName);
                        string replacestr = "<systemOverviewManagerConfiguration dataSource=\"" + PCName + "\" initialCatalog=\"EtriccStatistics_" + sCurrentProject + "\"/>";
                        Console.WriteLine("=== replace " + replacestr);
                        epiaPathNav.ReplaceSelf("<systemOverviewManagerConfiguration dataSource=\"" + PCName + "\" initialCatalog=\"EtriccStatistics_" + sCurrentProject + "\"/>");
                    }
                }
                epiaServerExeConfig.Save(System.IO.Path.Combine(EpiaPath, "Egemin.Epia.Server.exe.config"));
                #endregion

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (result == ConstCommon.TEST_PASS)
                    {
                        Console.WriteLine("\nTest scenario update configuration files: Pass");
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        Console.WriteLine("\nTest scenario  update configuration files: *FAIL*");
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

            }
        }
        #endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region StartEpiaServerShell
		public static void StartEpiaServerShell(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
			try
			{
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

                if (sPaeserConfiguratorConnectToComputerOK == false)
                {
                    sErrorMessage = "parser configurator failed to connect to the computer , this testcase cannot be tested";
                    return;
                }

				//========================   SERVER =================================================
                if (ProjServerOrShellStartup.ServerStartup("Epia Server", sServerRunAs, ref sErrorMessage, slogFilePath, sOnlyUITest) == false)
                {
                    sEpiaServerStartupOK = false;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
				
				//========================   SHELL =================================================
				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    Thread.Sleep(5000);
					Console.WriteLine("EPIA SERVER Service Started : ");
					Thread.Sleep(2000);

					// Add Open window Event Handler
					  Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);
					sEventEnd = false;
					#region  Shell
					ProcessUtilities.StartProcessNoWait( m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() +
						@"\Egemin\Epia Shell",
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

				if (TestCheck == ConstCommon.TEST_PASS)
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
					{
						Console.WriteLine("Application maeForm name : " + aeForm.Current.Name);
                        // resize shell 
                        TransformPattern tranform = aeForm.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                        if (tranform != null)
                        {
                            tranform.Resize(System.Windows.Forms.SystemInformation.VirtualScreen.Width - System.Windows.Forms.SystemInformation.VirtualScreen.Width * 0.1,
                                System.Windows.Forms.SystemInformation.VirtualScreen.Height - System.Windows.Forms.SystemInformation.VirtualScreen.Height * 0.2);
                            Thread.Sleep(1000);
                            tranform.Move(0, 0);
                        }

						TestCheck = ConstCommon.TEST_PASS;
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
					Console.WriteLine( testname+": Pass");
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
					result = ConstCommon.TEST_PASS;
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
				//Input.SendKeyboardInput(System.Windows.Input.Key.Space, true);
			}
		}
		#endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OpenThisStatisticsReport
        public static void OpenThisStatisticsReport(string reportType, string reportGroup, string reportName, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + reportName + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, reportName, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sErrorScreen = false;

            AutomationElement aeTreeView = null;
            AutomationElement aeEtriccNode = null;
            //string reportType = "Performance";
            //string reportGroup = "Vehicles";
            //string reportName = ReportName.PERFORMANCE_VEHICLES_ModeOverview;
            string fromDate = "11/25/2010";
            string toDate = "11/27/2010";
            string validateData = "02:01:13";

            try
            {
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

                if (sPaeserConfiguratorConnectToComputerOK == false)
                {
                    sErrorMessage = "parser configurator failed to connect to the computer , this testcase cannot be tested";
                    return;
                }
                
                if (sEpiaServerStartupOK == false )
                {
                    sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                    return;
                }

                StatUtilities.GetThisReportTestData(reportName, "Demo",ref fromDate, ref toDate, ref validateData);
                AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                if (aeWindow == null)
                {
                    sErrorMessage = "Main Window not opend";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    AUICommon.ClearDisplayedScreens(aeWindow, 2);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    //Epia3Common.WriteTestLogMsg(slogFilePath, reportName + " -- before get report TreeView:" + sErrorMessage, sOnlyUITest);
                    #region// get report TreeView and clik to get report, will optimalised later
                    aeTreeView = StatUtilities.GetReportsTrieView("MainForm", "m_TreeView", 120, ref  sErrorMessage);
                    if (aeTreeView == null)
                    {
                        Console.WriteLine("aeTreeView not found : ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("aeTreeView found name : " + aeTreeView.Current.Name);
                        TreeWalker walker = TreeWalker.ControlViewWalker;
                        aeEtriccNode = walker.GetFirstChild(aeTreeView);
                        if (aeEtriccNode != null)
                        {
                            Console.WriteLine("aeTreeView node name is: " + aeEtriccNode.Current.Name);
                            Thread.Sleep(500);
                        }
                        else
                        {
                            sErrorMessage = "aeEtriccNode not found : ";
                            Console.WriteLine("aeEtriccNode not found : ");
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    #endregion
                }

                // Add Open ErrorScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIErrorEventHandler);

                sErrorScreenMessage = string.Empty;
                #region// Open report screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    //Epia3Common.WriteTestLogMsg(slogFilePath, reportName + " -- before FindPerformanceFinalReport:" + sErrorMessage, sOnlyUITest);
                    if (StatUtilities.FindPerformanceFinalReport(aeTreeView, reportType, reportGroup, reportName, ref sErrorMessage))
                    {
                        // FindPerformanceFinalReport find report name in navigator panel, if it is collasp, sErrorMessage will show not find and
                        // by expand it again, to find report name. if founf,  sErrorMessage = report name is found;
                        Console.WriteLine("aeReport is openning  : ");
                    }
                    else
                    {
                        Console.WriteLine("open aeReport is failed: ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                #endregion

                int ky = 0;
                AutomationElement aeOverview = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (reportName.Equals(ReportName.StatusGraphicalView))  // WIP: for system overview multiple close and open screen 
                    {
                        Console.WriteLine("reportName: " + reportName);
                        reportName = "System overview";
                        Console.WriteLine("New reportName: " + reportName);
                        Thread.Sleep(1000);
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(4000);
                    //Epia3Common.WriteTestLogMsg(slogFilePath, reportName + " -- before GetSelectedOverviewWindow:" + sErrorMessage, sOnlyUITest);
                    aeOverview = ProjBasicUI.GetSelectedOverviewWindow(reportName, ref sErrorMessage);
                    while (aeOverview == null && ky < 10)
                    {
                        Console.WriteLine("wait until selected " + reportName + " window open :" + ky++);
                        aeOverview = ProjBasicUI.GetSelectedOverviewWindow(reportName, ref sErrorMessage);
                        Thread.Sleep(1000);
                    }

                    //Epia3Common.WriteTestLogMsg(slogFilePath, reportName + " -- after GetSelectedOverviewWindow:" + sErrorMessage, sOnlyUITest);
                    if (sErrorScreen == true)
                    {
                        sErrorMessage = sErrorScreenMessage;
                        Console.WriteLine("open aeReport error sctreen: ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("sErrorScreen == false ");
                        if (aeOverview == null)
                        {
                            Console.WriteLine("Window not opened :" + sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                #region // validate report screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    //Epia3Common.WriteTestLogMsg(slogFilePath, reportName + " -- before validate report screen:" + sErrorMessage, sOnlyUITest);
                    if (StatUtilities.ValidateReportPerformanceVehiclesReport2(aeOverview, fromDate, toDate, validateData, ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                #endregion

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, reportName + ":" + sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    result = ConstCommon.TEST_PASS;
                    string ms = reportName + " window validation OK at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                    Epia3Common.WriteTestLogPass(slogFilePath, reportName, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(reportName + " Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement,
                   UIErrorEventHandler);
                Thread.Sleep(3000);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OpenThisStatisticsReport
        public static void OpenThisStatisticsReport(string reportType, string reportGroup, string monthlyOrHourly, string reportName, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + reportName + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, reportName, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sErrorScreen = false;

            AutomationElement aeTreeView = null;
            AutomationElement aeEtriccNode = null;
            //string reportType = "Performance";
            //string reportGroup = "Vehicles";
            //string reportName = ReportName.PERFORMANCE_VEHICLES_ModeOverview;
            string fromDate = "11/25/2010";
            string toDate = "11/27/2010";
            string validateData = "02:01:13";

            try
            {
                if (sSqlOrReportServiceOK == false)
                {
                    sErrorMessage = "sql or report service not running , this testcase cannot be tested";
                    return;
                }

                if (sPaeserConfiguratorConnectToComputerOK == false)
                {
                    sErrorMessage = "parser configurator failed to connect to the computer , this testcase cannot be tested";
                    return;
                }

                if (sEpiaServerStartupOK == false)
                {
                    sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                    return;
                }


                StatUtilities.GetThisReportTestData(reportName, "Demo", ref fromDate, ref toDate, ref validateData);
                AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                if (aeWindow == null)
                {
                    sErrorMessage = "Main Window not opend";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    AUICommon.ClearDisplayedScreens(aeWindow, 2);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region// get report TreeView and clik to get report, will optimalised later
                    aeTreeView = StatUtilities.GetReportsTrieView("MainForm", "m_TreeView", 120, ref  sErrorMessage);
                    if (aeTreeView == null)
                    {
                        Console.WriteLine("aeTreeView not found : ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("aeTreeView found name : " + aeTreeView.Current.Name);
                        TreeWalker walker = TreeWalker.ControlViewWalker;
                        aeEtriccNode = walker.GetFirstChild(aeTreeView);
                        if (aeEtriccNode != null)
                        {
                            Console.WriteLine("aeTreeView node name is: " + aeEtriccNode.Current.Name);
                            Thread.Sleep(500);
                        }
                        else
                        {
                            sErrorMessage = "aeEtriccNode not found : ";
                            Console.WriteLine("aeEtriccNode not found : ");
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    #endregion
                }

                // Add Open ErrorScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIErrorEventHandler);

                sErrorScreenMessage = string.Empty;
                #region// Open report screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.FindPerformanceFinalReport(aeTreeView, reportType, reportGroup, monthlyOrHourly, reportName, ref sErrorMessage))
                    {
                        Console.WriteLine("aeReport is openning  : ");
                    }
                    else
                    {
                        Console.WriteLine("open aeReport is failed: ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                #endregion


                int ky = 0;
                AutomationElement aeOverview = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (reportName.Equals(ReportName.StatusGraphicalView))  // WIP: for system overview multiple close and open screen 
                    {
                        Console.WriteLine("reportName: " + reportName);
                        reportName = "System overview";
                        Console.WriteLine("New reportName: " + reportName);
                        Thread.Sleep(1000);
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("------------ wait until selected " + reportName + " window open :" + ky++);
                    Thread.Sleep(4000);
                    aeOverview = ProjBasicUI.GetSelectedOverviewWindow(reportName, ref sErrorMessage);
                    while (aeOverview == null && ky < 10)
                    {
                        Console.WriteLine("wait until selected " + reportName + " window open :" + ky++);
                        aeOverview = ProjBasicUI.GetSelectedOverviewWindow(reportName, ref sErrorMessage);
                        Thread.Sleep(1000);
                    }

                    if (sErrorScreen == true)
                    {
                        sErrorMessage = sErrorScreenMessage;
                        Console.WriteLine("open aeReport error sctreen: ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("sErrorScreen == false ");
                        if (aeOverview == null)
                        {
                            Console.WriteLine("Window not opened :" + sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                #region // validate report screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.ValidateReportPerformanceVehiclesReport2(aeOverview, fromDate, toDate, validateData, ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                #endregion

                Thread.Sleep(2000);
                AutomationElement aeHourlyOrDailyOrMonthly = StatUtilities.RefetchNodeTreeView("MainForm", "m_TreeView", monthlyOrHourly, 120, ref sErrorMessage);
                if (aeHourlyOrDailyOrMonthly == null)
                {
                    Console.WriteLine(monthlyOrHourly + " NOT FOUND After aeGroup node is expanded");
                }
                else
                {
                    System.Windows.Point HourlyOrDailyOrMonthlyPt = AUIUtilities.GetElementCenterPoint(aeHourlyOrDailyOrMonthly);
                    Input.MoveToAndDoubleClick(HourlyOrDailyOrMonthlyPt);
                    Thread.Sleep(2000);
                }                               

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, reportName + ":" + sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sErrorMessage = string.Empty;
                    result = ConstCommon.TEST_PASS;
                    string ms = reportName + " window validation OK at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                    Epia3Common.WriteTestLogPass(slogFilePath, reportName, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(reportName + " Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement,
                   UIErrorEventHandler);
                Thread.Sleep(3000);
            }
        }
        #endregion
		#endregion Test Cases ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Event ------------------------------------------------------------------------------------------------
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
			else if (name.Equals("Egemin.Epia.Server"))
			{
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
				Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
				System.Windows.Automation.Condition c = new AndCondition(
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
					System.Windows.Automation.Condition c2 = new AndCondition(
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
							TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS), sBuildNr);
						string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

						if (quality.Equals("GUI Tests Failed"))
						{
							//
							// check testinfo file line count if Line count = 2, and if sTotalFailed == 0 ---> update status to QUI Test Passed
							// this means I have deleted testinfo file due to the previous test failure; and rerun this test again. and if OK I should
							// change this satus to 'Pass'
							//_---
							string path = sBuildDropFolder + "\\TestResults";
							string testPC = System.Environment.MachineName;

							// read allline from test info file
							int count = 0;
							StreamReader reader = File.OpenText(Path.Combine(path, ConstCommon.TESTINFO_FILENAME));
							string infoline = reader.ReadLine();
							Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " read line :" + infoline, sOnlyUITest);

							while (infoline != null && infoline.Length > 0)
							{
								Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " read line while :" + infoline, sOnlyUITest);
								count++;
								infoline = reader.ReadLine();
							}
							reader.Close();

							Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " number of line in testinfo.txt file :" + count, sOnlyUITest);


							/*if (count == 2)
							{    
								if (sTotalFailed == 0)
									TestTools.TfsUtilities.UpdateBuildQualityStatusEvenHasFailedStatus(logger, uri,
									TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS),
									"GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
							}*/
							//---------
							Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
						}
						else
						{
							//if (sTotalFailed == 0)
							//    TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
							//    TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS),
							//    "GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
							//else
							TfsUtilities.UpdateBuildQualityStatus(logger, uri,
							TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS), "GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
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
                            TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Failed", testout, sInfoFileKey);
							Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Failed:" + ConstCommon.EPIA, sOnlyUITest);
							//}
						}
					}

                    if (sAutoTest)
                    {
                        if (sUseExcel)
                            TestTools.FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
                    }
				}

				#endregion
				System.Environment.Exit(1);
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
		static string uninstallWindowName = "Programs and Features";
		static string sgrid = "Folder View";
		static private bool IsApplicationInstalled(string ApplicationType)
		{
			bool applicationInstalled = false;

			System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName);
			AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
			if (appElement != null)
			{   // (1) Programs and Features main window
				Console.WriteLine("Programs and Features main window opend");
				Wait(1);
				Console.WriteLine("Searching programs item button...");
				AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(appElement, sgrid);
				if (aeGridView != null)
					Console.WriteLine("Gridview found...");
				Wait(1);
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
								Wait(5);
								break;
							}
							break;
						case "EtriccCore":
							if (aeProgram[i].Current.Name.StartsWith("E'tricc")
								   && aeProgram[i].Current.Name.IndexOf("Statistics") < 0
									&& aeProgram[i].Current.Name.IndexOf("Shell") < 0
								 && aeProgram[i].Current.Name.IndexOf("Playback") < 0
								&& aeProgram[i].Current.Name.IndexOf("HostTest") < 0)
							{
								applicationInstalled = true;
								Console.WriteLine(aeProgram[i].Current.Name + " is installed");
								Wait(5);
								break;
							}
							break;
						case "EtriccShell":
							if (aeProgram[i].Current.Name.StartsWith("E'tricc")
									&& aeProgram[i].Current.Name.IndexOf("Shell") > 0)
							{
								applicationInstalled = true;
								Console.WriteLine(aeProgram[i].Current.Name + " is installed");
								Wait(5);
								break;
							}
							break;
						case "Ewcs":
							if (aeProgram[i].Current.Name.StartsWith("E'wcs")
								&& aeProgram[i].Current.Name.IndexOf("TestPrograms") < 0)
							{
								applicationInstalled = true;
								Console.WriteLine(aeProgram[i].Current.Name + " is installed");
								Wait(5);
								break;
							}
							break;
						case "EwcsTestProgram":
							if (aeProgram[i].Current.Name.StartsWith("E'wcs")
								&& aeProgram[i].Current.Name.IndexOf("TestPrograms") > 0)
							{
								applicationInstalled = true;
								Console.WriteLine(aeProgram[i].Current.Name + " is installed");
								Wait(5);
								break;
							}
							break;
					}
				}
			}
			return applicationInstalled;
		}

		static private void Wait(int seconds)
		{
			System.Threading.Thread.Sleep(seconds * 1000);
		}
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region OnInstallEtricc5UIEvent
		public static void OnInstallEtricc5UIEvent(object src, AutomationEventArgs args)
		{
			Console.WriteLine("OnInstallEtricc5UIEvent");
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
			string str = string.Format("OnInstallEtricc5UIEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
			Console.WriteLine(str);

			Thread.Sleep(5000);
			if (name.Equals("Error"))
			{
				TestCheck = ConstCommon.TEST_FAIL;
				Console.WriteLine("Name is ------------:" + name);
				AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
				sEventEnd = true;
			}
			else if (name.Equals("<NoName>"))
			{
				Console.WriteLine("Do nothing Name is ------------:" + name);
				sEventEnd = false;
			}
			else if (name.StartsWith("E'tricc"))
			{
				AutomationElement rootElement = AutomationElement.RootElement;
				#region Install EtriccCore
				Console.WriteLine("Searching for main window");

				PropertyCondition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
				AutomationElement appElement = rootElement.FindFirst(TreeScope.Children, condition);

				DateTime startTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - startTime;
                while (appElement == null && mTime.TotalSeconds < 60)
				{
					Wait(2);
					appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
					mTime = DateTime.Now - startTime;
                    if (mTime.TotalSeconds > 60)
					{
						System.Windows.Forms.MessageBox.Show("After one minute no Installer Window Form found");
						return;
					}
				}

				// (1) Welcom Main window
				Console.WriteLine("EtriccCore Main Form found ...");
				Console.WriteLine("Searching next button...");
				AutomationElement btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
				if (btnNext != null)
				{   // (2) Components
					AUIUtilities.ClickElement(btnNext);
					appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
					Console.WriteLine("Welcom Etricc Core window opend...");
					Console.WriteLine("Searching next button...");
					btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
					if (btnNext != null)
					{
						AUIUtilities.ClickElement(btnNext);
						Wait(3);
						appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
						Console.WriteLine("Welcom Etricc Core window opend...");
						Console.WriteLine("Searching next button...");

						Wait(3);
						Console.WriteLine("License agreement window opend...");
						Console.WriteLine("Searching I Agree button...");
						AutomationElement aeBtnAgree = AUIUtilities.GetElementByNameProperty(appElement, "I Agree");

						if (aeBtnAgree == null)
							Console.WriteLine("Agree button not found...");

						AUIUtilities.ClickElement(aeBtnAgree);
						Wait(2);
						btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");

						AUIUtilities.ClickElement(btnNext);
						Wait(3);

						appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
						Console.WriteLine("Select componts to install window opend...");
						Console.WriteLine("Searching Next button...");

						btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
						if (btnNext != null)
						{
							AUIUtilities.ClickElement(btnNext);
							Wait(3);
							appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
							Console.WriteLine("Environment Configuration window opend...");
							Console.WriteLine("Searching Next button...");
							btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
							if (btnNext != null)
							{
								AUIUtilities.ClickElement(btnNext);
								Wait(3);
								appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
								Console.WriteLine("Select Installation Folder window opend...");
								Console.WriteLine("Searching Next button...");
								btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
								if (btnNext != null)
								{
									AUIUtilities.ClickElement(btnNext);
									Wait(3);
									appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
									Console.WriteLine("Confirm Installation window opend...");
									Console.WriteLine("Searching Next button...");
									btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
									if (btnNext != null)
									{
										AUIUtilities.ClickElement(btnNext);
										Wait(3);
										appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
										Console.WriteLine("Installing Etricc  window opend...");
										Console.WriteLine("Installing Etricc  window move to left...");

										AutomationElement aeTitleBar =
											AUIUtilities.FindElementByID("TitleBar", appElement);

										Point pt1 = new Point(
											(aeTitleBar.Current.BoundingRectangle.Left + aeTitleBar.Current.BoundingRectangle.Right) / 2,
											(aeTitleBar.Current.BoundingRectangle.Bottom + aeTitleBar.Current.BoundingRectangle.Top) / 2);

										Point newPt1 = new Point(200, 100);
										Input.MoveTo(pt1);

										Wait(1);
										Input.SendMouseInput(pt1.X, pt1.Y, 0, SendMouseInputFlags.LeftDown | SendMouseInputFlags.Absolute);
										Wait(1);
										Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.Move | SendMouseInputFlags.Absolute);
										Wait(1);
										Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);
										Wait(3);

										Console.WriteLine("try to find  FrmLauncherFunctionality window...");

										Condition conditionFrm = new PropertyCondition(AutomationElement.AutomationIdProperty, "FrmLauncherFunctionality");
										Condition conditionFrm2 = new PropertyCondition(AutomationElement.AutomationIdProperty, "FrmSecurityFunctionality");

										AutomationElement frmElement = null;
										while (frmElement == null)
										{
											Wait(5);
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
												AUIUtilities.ClickElement(aeBtnNext);
											}
										}


										AutomationElement frmElement2 = null;
										while (frmElement2 == null)
										{
											Wait(5);
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
												AUIUtilities.ClickElement(aeBtnNext);
											}
										}

										Console.WriteLine("Installation complete  window opend...");
										Console.WriteLine("Searching close button...");

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
											Wait(5);
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
						}
					}
				}

				#endregion
				sEventEnd = true;
			}
			else
			{
				TestCheck = ConstCommon.TEST_FAIL;
				Console.WriteLine("Name is ------------:" + name);
				AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
				sEventEnd = false;
			}
			
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region OnParserConfiguratorConnectComputerEvent
		public static void OnParserConfiguratorConnectComputerEvent(object src, AutomationEventArgs args)
		{
			Console.WriteLine("OnParserConfiguratorConnectComputerEvent");
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
			string str = string.Format("OnParserConfiguratorConnectComputerEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
			Console.WriteLine(str);

			Thread.Sleep(5000);
			if (name.Equals("Error"))
			{
				TestCheck = ConstCommon.TEST_FAIL;
				Console.WriteLine("Name is ------------:" + name);
				AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
				sEventEnd = true;
			}
			else if (name.Equals("<NoName>"))
			{
				Console.WriteLine("Do nothing Name is ------------:" + name);
				sEventEnd = false;
			}
			else if (name.Equals("Connect to"))
			{
				Console.WriteLine("Fill in Computer name : " + System.DateTime.Now);
				AutomationElement aeComboBoxComputerNameEdit = null;
				Thread.Sleep(3000);
				//DateTime mAppTime = DateTime.Now;
				//TimeSpan mTime = DateTime.Now - mAppTime;
				//while (aeComboBoxComputerNameEdit == null && mTime.Minutes < 2)
				//{
					Console.WriteLine("Find Application aeComboBoxComputerNameEdit : " + System.DateTime.Now);
					aeComboBoxComputerNameEdit = AUIUtilities.FindElementByType(ControlType.Edit, element);
					Console.WriteLine("Application aeComboBoxComputerNameEdit name : " + System.DateTime.Now);
				 //   mTime = DateTime.Now - mAppTime;
					Thread.Sleep(2000);
				 //   Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
			   // }

				if (aeComboBoxComputerNameEdit == null)
				{
					//sErrorMessage = "aeComboBoxComputerNameEdit not found";
					//TestCheck = ConstCommon.TEST_FAIL;
					sEventEnd = true;
					return;
				}
				else
				{
					Thread.Sleep(500);
					Point pnt = TestTools.AUIUtilities.GetElementCenterPoint(aeComboBoxComputerNameEdit);
					Input.MoveTo(pnt);
					Thread.Sleep(1000);
					ValuePattern vp = (ValuePattern)aeComboBoxComputerNameEdit.GetCurrentPattern(ValuePattern.Pattern);
					Thread.Sleep(1000);
					string getValue = vp.Current.Value;
					Thread.Sleep(2000);
					vp.SetValue(PCName);
				}


				AutomationElement aeConnectButton = null;
				string BtnConnectId = "m_BtnConnect";
				Thread.Sleep(3000);
				//mAppTime = DateTime.Now;
				//mTime = DateTime.Now - mAppTime;
				//while (aeConnectButton == null && mTime.Minutes < 5)
				//{
					Console.WriteLine("xFind Application aeConnectButton : " + System.DateTime.Now);
					aeConnectButton = AUIUtilities.FindElementByID(BtnConnectId, element);
					Console.WriteLine("xApplication aeConnectButton : " + System.DateTime.Now);
			   //     mTime = DateTime.Now - mAppTime;
					Thread.Sleep(20000);
			   //     Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
			   // }

				if (aeConnectButton == null)
				{
					sErrorMessage = "aeConnectButton not found";
					TestCheck = ConstCommon.TEST_FAIL;
					sEventEnd = true;
					return;
				}
				else
				{
					Thread.Sleep(500);
					TestTools.AUIUtilities.ClickElement(aeConnectButton);
					Thread.Sleep(2000);
					sEventEnd = true;
				}
			}
			else
			{
				Console.WriteLine("xxxxx        Name is ------------:" + name);
				sEventEnd = false;
			}
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region OnCreateNewParseProjectUIEvent
		public static void OnCreateNewParseProjectUIEvent(object src, AutomationEventArgs args)
		{
			Console.WriteLine("OnCreateNewParseProjectUIEvent");
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
			string str = string.Format("OnCreateNewParseProjectUIEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
			Console.WriteLine(str);

			Thread.Sleep(5000);
			if (name.Equals("Error"))
			{
				TestCheck = ConstCommon.TEST_FAIL;
				Console.WriteLine("Name is ------------:" + name);
				AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
				sEventEnd = true;
			}
			else if (name.Equals("<NoName>"))
			{
				Console.WriteLine("Do nothing Name is ------------:" + name);
				sEventEnd = false;
			}
			else if (element.Current.AutomationId.Equals("CreateOrEditParseProjectForm"))
			{
				Console.WriteLine("Create new Parse project panel is opend -------------- : " + System.DateTime.Now);
				string prijectNameId = "m_TextBoxName";     // "ControlType.Edit"
				AutomationElement aeProjectName = null;
				
				string EtriccXmlFilename = @"C:\" +  TestTools.OSVersionInfoClass.ProgramFilesx86FolderName()
								+ @"\Egemin\AutomaticTesting\EtriccStatistics\"+sCurrentProject+@"\Data\Xml\" + sCurrentProject + ".xml";
				string EtriccXmlFilenameId = "m_TextBoxEtriccXmlFilename";  //"ControlType.Edit"
				AutomationElement aeEtriccXmlFilename = null;

				string EtriccStatLogFolder =  @"C:\" +  TestTools.OSVersionInfoClass.ProgramFilesx86FolderName()
								+ @"\Egemin\AutomaticTesting\EtriccStatistics\" + sCurrentProject + @"\Data\Xml\Stat";
				string EtriccStatLogFolderId = "m_TextBoxEtriccStatLogFolder";  // "ControlType.Edit"
				AutomationElement aeEtriccStatLogFolder = null;

				string OKButtonId = "m_BtnOK";  // "ControlType.Button"
				AutomationElement aeOKButton = null;

				aeProjectName = AUIUtilities.FindElementByID(prijectNameId, element);
				if (aeProjectName != null)
				{
					Thread.Sleep(1000);
					ValuePattern vp = aeProjectName.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
					vp.SetValue(sCurrentProject);
				}
				else
				{
					sErrorMessage = "aeProjectName not found name";
					Console.WriteLine(sErrorMessage);
					TestCheck = ConstCommon.TEST_FAIL;
					sEventEnd = true;
					return;
				}

				aeEtriccXmlFilename = AUIUtilities.FindElementByID(EtriccXmlFilenameId, element);
				if (aeEtriccXmlFilename != null)
				{
					Thread.Sleep(1000);
					ValuePattern vp = aeEtriccXmlFilename.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
					vp.SetValue(EtriccXmlFilename);
				}
				else
				{
					sErrorMessage = "aeEtriccXmlFilename not found name";
					Console.WriteLine(sErrorMessage);
					TestCheck = ConstCommon.TEST_FAIL;
					sEventEnd = true;
					return;
				}

				aeEtriccStatLogFolder = AUIUtilities.FindElementByID(EtriccStatLogFolderId, element);
				if (aeEtriccStatLogFolder != null)
				{
					Thread.Sleep(1000);
					ValuePattern vp = aeEtriccStatLogFolder.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
					vp.SetValue(EtriccStatLogFolder);
				}
				else
				{
					sErrorMessage = "aeEtriccStatLogFolder not found name";
					Console.WriteLine(sErrorMessage);
					TestCheck = ConstCommon.TEST_FAIL;
					sEventEnd = true;
					return;
				}

				aeOKButton = AUIUtilities.FindElementByID(OKButtonId, element);
				if (aeOKButton != null)
				{
					Thread.Sleep(1000);
					Input.MoveToAndClick(aeOKButton);
					Thread.Sleep(1000);
					sEventEnd = true;
				}
				else
				{
					sErrorMessage = "aeOKButton not found name";
					Console.WriteLine(sErrorMessage);
					TestCheck = ConstCommon.TEST_FAIL;
					sEventEnd = true;
					return;
				}
			}
			else
			{
				//TestCheck = ConstCommon.TEST_FAIL;
				Console.WriteLine("xxxxx        Name is ------------:" + name);
				//AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
				sEventEnd = false;
			}
		   
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region OnErrorUIEvent
		public static void OnErrorUIEvent(object src, AutomationEventArgs args)
		{
			Console.WriteLine("OnErrorUIAEvent");
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
			string str = string.Format("Error :={0} : {1}", name, args.EventId.ProgrammaticName);
			Console.WriteLine(str);

			Thread.Sleep(5000);

			if (element.Current.AutomationId.Equals("ErrorScreen"))
			{
				TestCheck = ConstCommon.TEST_FAIL;
				Console.WriteLine("Name is ------------:" + name);
				StatUtilities.ErrorWindowHandling(element, ref sErrorMessage);
                sErrorScreenMessage = sErrorMessage;
				TestCheck = ConstCommon.TEST_FAIL;
				Console.WriteLine("OnErrorUIAEvent ------------ testcheck:" + TestCheck);
				sEventEnd = true;
				sErrorScreen = true;
				return;
				//Thread.Sleep(5000);
			}
			else if (element.Current.Name.StartsWith("Microsoft .NET"))
			{
				TestCheck = ConstCommon.TEST_FAIL;
				Console.WriteLine("Name is ------------:" + name);
				StatUtilities.ErrorMicrosoftNETFrameWorkWindowHandling(element, ref sErrorMessage);
				TestCheck = ConstCommon.TEST_FAIL;
				Console.WriteLine("OnErrorUIAEvent ------------ testcheck:" + TestCheck);
				sEventEnd = true;
				sErrorScreen = true;
				return;
				//Thread.Sleep(5000);
			}
			else
			{
				Console.WriteLine("Do nothing --------------------------------------------------- ------------:" + name);
				return;
			}
		}
		#endregion

		//&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
		#endregion Event +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region TestLog ----------------------------------------------------------------------------------------------
		public static void SendEmail(string resultFile)
		{
            string str1 = DeployUtilities.GetTestReportContentString(sTotalCounter, sTotalPassed, sTotalFailed, sTotalException, sTotalUntested,
               sCurrentPlatform, sInstallMsiDir); // AnyCPU 


			ProcessUtilities.SendTestResultToDevelopers( resultFile, sTestApp, sBuildDef, logger, sTotalFailed,
				sBuildNr/*used for email title*/, str1/*content*/, sSendMail);
		}
		#endregion TestLog +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	}

}
