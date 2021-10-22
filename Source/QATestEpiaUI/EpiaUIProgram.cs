using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Automation;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using TestTools;
using TFSQATestTools;
using Excel = Microsoft.Office.Interop.Excel;

namespace QATestEpiaUI
{
	class EpiaUIProgram
	{
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Fields of Program (38)
		internal static TestTools.Logger logger;

		static string[] EpiaServerDlls
		   = { "C5.dll",
				"Egemin.Epia.Components.dll", 
				"Egemin.Epia.Foundation.Alarming.Interfaces.dll",
				"Egemin.Epia.Foundation.Communication.dll",
				"Egemin.Epia.Foundation.Communication.Interfaces.dll",
				"Egemin.Epia.Foundation.ComponentManagement.Configuration.dll",
				"Egemin.Epia.Foundation.ComponentManagement.Development.dll",
				"Egemin.Epia.Foundation.dll",
				"Egemin.Epia.Foundation.Globalization.Interfaces.dll",
				"Egemin.Epia.Foundation.License.Interfaces.dll",
				"Egemin.Epia.Foundation.Security.Interfaces.dll",
				"Egemin.Epia.Foundation.SqlRptServices.Discovery.Interfaces.dll",
				"Egemin.Epia.Foundation.VersionInformation.Interfaces.dll",
				"Egemin.Epia.Presentation.Configuration.Interfaces.dll",
				"LinqTransaction.dll",
				"Microsoft.Practices.EnterpriseLibrary.Common.dll",
				"Microsoft.Practices.EnterpriseLibrary.Logging.dll",
				"Microsoft.Practices.EnterpriseLibrary.Validation.dll",
				"Microsoft.Practices.ObjectBuilder2.dll",
				"Plossum.dll",
											   };

		static string[] EpiaShellDlls
		   = { "C5.dll",
				"Egemin.Epia.CompositeUI.dll", 
				"Egemin.Epia.Foundation.Alarming.Interfaces.dll",
				"Egemin.Epia.Foundation.ComponentManagement.Configuration.dll",
				"Egemin.Epia.Foundation.ComponentManagement.Development.dll",
				"Egemin.Epia.Foundation.dll",
				"Egemin.Epia.Foundation.Globalization.Interfaces.dll",
				"Egemin.Epia.Foundation.License.Interfaces.dll",
				"Egemin.Epia.Foundation.Security.Interfaces.dll",
				"Egemin.Epia.Foundation.SqlRptServices.Discovery.Interfaces.dll",
				"Egemin.Epia.Foundation.VersionInformation.Interfaces.dll",
				"Egemin.Epia.Modules.RnD.dll",
				"Egemin.Epia.Modules.SqlRptServices.dll",
				"Egemin.Epia.Presentation.Configuration.Interfaces.dll",
				"Egemin.Epia.Presentation.Development.dll",
				"Egemin.Epia.Presentation.dll",
				"Infragistics.Practices.CompositeUI.WinForms.dll",
				"Infragistics2.Shared.v8.3.dll",
				"Infragistics2.Win.Misc.v8.3.dll",
				"Infragistics2.Win.UltraWinDock.v8.3.dll",
				"Infragistics2.Win.UltraWinEditors.v8.3.dll",
				"Infragistics2.Win.UltraWinExplorerBar.v8.3.dll",
				"Infragistics2.Win.UltraWinMaskedEdit.v8.3.dll",
				"Infragistics2.Win.UltraWinStatusBar.v8.3.dll",
				"Infragistics2.Win.UltraWinTabbedMdi.v8.3.dll",
				"Infragistics2.Win.UltraWinTabControl.v8.3.dll",
				"Infragistics2.Win.UltraWinToolbars.v8.3.dll",
				"Infragistics2.Win.UltraWinTree.v8.3.dll",
				"Infragistics2.Win.v8.3.dll",
				"Microsoft.Practices.CompositeUI.dll",
				"Microsoft.Practices.CompositeUI.WinForms.dll",
				"Microsoft.Practices.EnterpriseLibrary.Common.dll",
				"Microsoft.Practices.EnterpriseLibrary.Logging.dll",
				"Microsoft.Practices.EnterpriseLibrary.Validation.dll",
				"Microsoft.Practices.ObjectBuilder.dll",
				"Microsoft.Practices.ObjectBuilder2.dll",
				"Microsoft.Practices.Unity.dll",
				"Microsoft.ReportViewer.Common.dll",
				"Microsoft.ReportViewer.DataVisualization.dll",
				"Microsoft.ReportViewer.WinForms.dll",
				"Plossum.dll",
				"SCSFContrib.CompositeUI.WinForms.dll"
											   };
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
		static string sParentProgram = string.Empty;
		static string sTestType = "all";
		static string sTestDefinitionFile = string.Empty;
		static string[] mTestDefinitionTypes;
		static string sInfoFileKey = string.Empty;
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
		static string m_SystemDrive = @"C:\";
		static string UserPassword = "Egemin01";
		static string sTargetPlatform = string.Empty;
		static string sCurrentPlatform = string.Empty;
		static string sTestResultFolder = string.Empty;
        static string sInstallMsiDir = @"C:\LocalTest\";
        static string sNetworkMap = "LocalTest";
        static string sDemoCaseCount = "1";
		// Testcase not used =================================
		static string sBrand = "XXXXXX";
		static string sErrorMessage;
		static bool sEventEnd;
		static string sExcelVisible = string.Empty;
		static bool sAutoTest = true;
		public static string sLayoutName = string.Empty;
		static string sServerRunAs = "Service";
		static bool sDemo;
		static string sSendMail = "false";
		static string sTFSServer = "http://Team2010App.TeamSystems.Egemin.Be:8080";
		// LOG=================================================================
		public static string slogFilePath = @"C:\";
		static string sOutFilename = "OutFilename";
		static string sOutFilePath = string.Empty;
		static StreamWriter Writer;
		// Build param ========================================================
		static IBuildServer m_BuildSvc;
		static bool TFSConnected = true;
		// excel 	--------------------------------------------------------
		static Excel.Application xApp;
        static int sHeaderContentsLength = 10;
		// default layout --------------------------------------------------
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
		#endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
		static string sEpiaServerFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server";
		static string sEpiaShellFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell";

		private static bool sEpiaServerStartupOK = true;
        private static bool sMicrosoftVisualStudioShellDesignDllInstalled = false;
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
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
		static void Main(string[] args)
		{
			try  // Get test PC info======================================
			{
	
				TestTools.HelpUtilities.SavePCInfo("y");
				TestTools.HelpUtilities.GetPCInfo(out PCName, out OSName, out OSVersion, out UICulture, out TimeOnPC);
                Console.WriteLine("<PCName : " + PCName + ">, <OSName : " + OSName + ">, <OSVersion : " + OSVersion+">");
                Console.WriteLine("<TimeOnPC : " + TimeOnPC + ">, <UICulture : " + UICulture+">");
				string windir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
				m_SystemDrive = Path.GetPathRoot(windir);
			}
			catch (Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Check PC-info");
			}

			sOnlyUITest = false;
			sBrand = System.Configuration.ConfigurationManager.AppSettings.Get("Brand");
			sEpiaServerFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\"+sBrand+ "\\Epia Server";
			sEpiaShellFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\" + sBrand + "\\Epia Shell";

			string onlyUITest = System.Configuration.ConfigurationManager.AppSettings.Get("OnlyUITest");
            if (onlyUITest.ToLower().StartsWith("true"))
				sOnlyUITest = true;

			UserPassword = System.Configuration.ConfigurationManager.AppSettings.Get("CurrentUserPassword");

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
						if (args[10].ToLower().StartsWith("true"))
							sAutoTest = true;
						else
							sAutoTest = false;

						sTFSServer = args[11];
						sServerRunAs = args[12];
						sExcelVisible = args[13];
						if (args[14].ToLower().StartsWith("true"))
							sDemo = true;
						else
							sDemo = false;

						if (args[15].ToLower().StartsWith("true"))
							sSendMail = "true";
						else
							sSendMail = "false";

						sTestDefinitionFile = args[16];
						sInfoFileKey = args[17];
                        sNetworkMap = args[18];
                        sDemoCaseCount = args[19];

                        sDemoCaseCount  = System.Configuration.ConfigurationManager.AppSettings.Get("DemoCaseCount");

						sTestResultFolder = sBuildDropFolder + "\\TestResults";
						if (!System.IO.Directory.Exists(sTestResultFolder))
							System.IO.Directory.CreateDirectory(sTestResultFolder);

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
                        Epia3Common.WriteTestLogMsg(slogFilePath, "19) DemoCaseCount: " + sDemoCaseCount, sOnlyUITest);

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
                        Console.WriteLine("19) sDemoCaseCount: " + sDemoCaseCount);

						mTestDefinitionTypes = System.IO.File.ReadAllLines(sTestDefinitionFile);

						for (int i = 0; i < mTestDefinitionTypes.Length; i++)
						{
							Console.WriteLine(i + " testdefinition : " + mTestDefinitionTypes[i]);
						}
					}
				}
				catch (Exception ex)
				{
					System.Windows.Forms.MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Validate command-line params");
				}
			}
			else
			{   // if only UI test, the msi folder is from the config file
                string currentLocalTestedBuildDef = System.Configuration.ConfigurationManager.AppSettings.Get("CurrentLocalTestedBuildDef");
                sInstallMsiDir = OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand + @"\AutomaticTesting\Setup\Epia\" + currentLocalTestedBuildDef;
				if (!File.Exists(Path.Combine(sInstallMsiDir, "Epia.msi")))
					System.Windows.Forms.MessageBox.Show("Please select a folder where Epia.msi is included");
			}

			if (!sOnlyUITest)
			{
				if (sAutoTest)
				{
					try
					{
						Console.WriteLine(" Get TFS Server  slogFilePath: " + slogFilePath);

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
                                TestTools.MessageBoxEx.Show("Team Foundation services are not available from server\nWill try to reconnect the Server ...\nexception messge:"+ex.Message,
								kTime++ + " During E'pia UI Testing, please not touch the screen, time :" + DateTime.Now.ToLongTimeString(), (uint)Tfs.ReconnectDelay );
								System.Threading.Thread.Sleep( Tfs.ReconnectDelay );
								conn = false;
							}
							catch (Exception ex)
							{
								TestTools.MessageBoxEx.Show("TeamFoundation getService Exception:" + ex.Message + " ----- " + ex.StackTrace,
									 kTime++ + " This is automatic testing, please not touch the screen: exception time:" + DateTime.Now.ToLongTimeString(), (uint)Tfs.ReconnectDelay );
								System.Threading.Thread.Sleep( Tfs.ReconnectDelay );
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
					TFSConnected = false;
			}

			Console.WriteLine("Test started:");
			Epia3Common.WriteTestLogMsg(slogFilePath, "Test started: ", sOnlyUITest);
			sTestCaseName[0] = EPIA_SERVER_INSTALLATION_INTEGRITY;
			sTestCaseName[1] = EPIA_SHELL_INSTALLATION_INTEGRITY;
            sTestCaseName[2] = SHELL_CLOSEE_WITHIN_MIN_AFTER_SERVER_DOWN;
			sTestCaseName[3] = START_EPIA_SERVER_SHELL;
			sTestCaseName[4] = RESOURCEID_INTEGRITY_CHECK;
			sTestCaseName[5] = LAYOUT_FIND_LAYOUT_PANEL;
			sTestCaseName[6] = LAYOUT_INITIAL_X_POSITION;
			sTestCaseName[7] = LAYOUT_INITIAL_Y_POSITION;
			sTestCaseName[8] = LAYOUT_INITIAL_WIDTH;
			sTestCaseName[9] = LAYOUT_INITIAL_HEIGHT;
			sTestCaseName[10] = LAYOUT_ALLOW_RESIZE;
			sTestCaseName[11] = LAYOUT_FULL_SCREEN;
			sTestCaseName[12] = LAYOUT_MAXIMIZED;
			sTestCaseName[13] = LAYOUT_RIBBON_ON;
			//sTestCaseName[9] = LAYOUT_TITLE;
			sTestCaseName[14] = LAYOUT_CANCEL_BUTTON;
			sTestCaseName[15] = SECURITY_OPEN_ACCOUNT_DETAIL_SCREEN;
            sTestCaseName[16] = SECURITY_ADD_NEW_ROLE;
			sTestCaseName[17] = SECURITY_OPEN_ROLE_DETAIL_SCREEN;
			sTestCaseName[18] = SECURITY_EDIT_ROLE;
            sTestCaseName[19] = SECURITY_ADD_NEW_ACCOUNT;
            sTestCaseName[20] = SYSTEM_ADD_SERVICE;
            sTestCaseName[21] = SYSTEM_OPEN_SERVICE_DETAIL_SCREEN;
			sTestCaseName[22] = MULTI_LANGUAGE_CHECK;
			sTestCaseName[23] = SHELL_CONFIGURATION_SECURITY;
			sTestCaseName[24] = LOGON_CURRENT_USER;
            sTestCaseName[25] = ACCOUNT_INACTIVITY_TIMEOUT_LOGOUT;
            sTestCaseName[26] = ROLE_INACTIVITY_TIMEOUT_LOGOUT;
            sTestCaseName[27] = ACCOUNT_INACTIVITY_TIMEOUT_SHUTDOWN;
            sTestCaseName[28] = ROLE_INACTIVITY_TIMEOUT_SHUTDOWN;
			sTestCaseName[29] = SHELL_SHUTDOWN;
			sTestCaseName[30] = SHELL_LOGOFF;
			sTestCaseName[31] = EPIA4_CLOSEE;
			sTestCaseName[32] = EPIA4_CLEAN_UNINSTALL_CHECK;
			sTestCaseName[33] = EPIA4_CLEAN_SHELL_INSTALL_CHECK;
			sTestCaseName[34] = EPIA4_CLEAN_SERVER_INSTALL_CHECK;
			sTestCaseName[35] = UNINSTALL_EPIA_RESOURCE_EDITOR_IF_ALREADY_INSTALLED;
			sTestCaseName[36] = INSTALL_EPIA_RESOURCE_EDITOR;
			sTestCaseName[37] = LOAD_EPIA_RESOURCE_FILES;
			sTestCaseName[38] = FILTER_EPIA_RESOURCE_FILES;
			sTestCaseName[39] = EDIT_SAVE_RESOURCE_FILES;
			//=============================================//
			try
			{
				if (!sOnlyUITest)
				{
					ProcessUtilities.CloseProcess( "EXCEL" );
					ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SERVER );
					ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SHELL );
					Thread.Sleep(1000);
				}

                xApp = new Excel.Application();
                string[] sHeaderContents = { System.DateTime.Now.ToString("MMMM-dd") + "*" + "Epia" +  " UI Test Scenarios",
                                              "Test Machine:" + "*" + PCName,
                                               "Tester::" + "*" + System.Security.Principal.WindowsIdentity.GetCurrent().Name,
                                               "OSName:" + "*" + OSName,
                                               "OS Version:" + "*" + OSVersion,
                                               "UI Culture:" + "*" + UICulture,
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
				int aantal = 40;
                if (sDemo)
                {
                    aantal = Convert.ToInt16(sDemoCaseCount);
                    //aantal = 1;
                }

				if (sOnlyUITest)   // get test case from application config file, otherwise, test all
				{
					sTestType = System.Configuration.ConfigurationManager.AppSettings.Get("TestType");
					if (sTestType.ToLower().StartsWith("all"))
					{
						aantal = 40;
					}
					else
					{
						int thisTest = 0;
						if (sTestType.IndexOf("-") > 0)
						{
							Console.WriteLine("first num: " + (sTestType.Substring(0, sTestType.IndexOf("-"))));
							Console.WriteLine("second num: " + (sTestType.Substring(sTestType.IndexOf("-") + 1)));

							thisTest = Convert.ToInt16(sTestType.Substring(0, sTestType.IndexOf("-")));
							Console.WriteLine("thisTest: " + thisTest);
							Thread.Sleep(15000);
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
                        if (Counter < 20 && Counter > 4)
                            aeForm = EpiaUtilities.GetMainWindow("MainForm");
					}
				}

				while (Counter < aantal)
				{
					sResult = ConstCommon.TEST_UNDEFINED;
					switch (sTestCaseName[Counter])
					{
						case EPIA_SERVER_INSTALLATION_INTEGRITY:
                            if (DeployUtilities.getThisPCOS().StartsWith("Windows8.64") 
                                || DeployUtilities.getThisPCOS().StartsWith("WindowsServer2012.64"))
                            {
                                EpiaServerIntegrityCheckNet45(EPIA_SERVER_INSTALLATION_INTEGRITY, aeForm, out sResult);
                            }
                            else
							    EpiaServerIntegrityCheck(EPIA_SERVER_INSTALLATION_INTEGRITY, aeForm, out sResult);
							break;
						case EPIA_SHELL_INSTALLATION_INTEGRITY:
							EpiaShellIntegrityCheck(EPIA_SHELL_INSTALLATION_INTEGRITY, aeForm, out sResult);
							break;
						case START_EPIA_SERVER_SHELL:
							StartEpiaServerShell(START_EPIA_SERVER_SHELL, aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show("StartEpiaServerShell", "Click OK to continue test", MessageBoxButtons.OK);
							break;
						case RESOURCEID_INTEGRITY_CHECK:
							ResourceIdIntegrityCheck(RESOURCEID_INTEGRITY_CHECK, aeForm, out sResult);
							break;
						case LAYOUT_FIND_LAYOUT_PANEL:
							LayoutFindLayoutPanel(LAYOUT_FIND_LAYOUT_PANEL, aeForm, out sResult);
							break;
						case LAYOUT_INITIAL_X_POSITION:
							LayoutInitialXPosition(LAYOUT_INITIAL_X_POSITION, aeForm, out sResult);
							//System.Windows.Forms.MessageBox.Show("LayoutInitialXPosition", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case LAYOUT_INITIAL_Y_POSITION:
							LayoutInitialYPosition(LAYOUT_INITIAL_Y_POSITION, aeForm, out sResult);
							//System.Windows.Forms.MessageBox.Show("LayoutInitialYPosition", "Click OK to continue", MessageBoxButtons.OK);
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
						case SECURITY_OPEN_ACCOUNT_DETAIL_SCREEN:
                            EpiaOpenDetailScreen(SECURITY_OPEN_ACCOUNT_DETAIL_SCREEN, "Security", "Accounts", aeForm, out sResult);
							break;
						case SECURITY_ADD_NEW_ROLE:
                            SecurityAddNewRole(SECURITY_ADD_NEW_ROLE, aeForm, out sResult);
							break;
						case SECURITY_OPEN_ROLE_DETAIL_SCREEN:
							EpiaOpenDetailScreen(SECURITY_OPEN_ROLE_DETAIL_SCREEN, "Security", "Roles", aeForm, out sResult);
							break;
						case SECURITY_EDIT_ROLE:
							SecurityEditRole(SECURITY_EDIT_ROLE, aeForm, out sResult);
							//System.Windows.Forms.MessageBox.Show("SecurityEditRole", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case SECURITY_ADD_NEW_ACCOUNT:
                            SecurityAddNewGeneralAccount(SECURITY_ADD_NEW_ACCOUNT, aeForm, out sResult);
							//System.Windows.Forms.MessageBox.Show("SecurityAddNewGeneralAccount", "Click OK to continue", MessageBoxButtons.OK);
							break;
                        case SYSTEM_ADD_SERVICE:
                            SystemAddService(SYSTEM_ADD_SERVICE, aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show("SystemAddService", "Click OK to continue", MessageBoxButtons.OK);
							break;
                        case SYSTEM_OPEN_SERVICE_DETAIL_SCREEN:
                            EpiaOpenDetailScreen(SECURITY_OPEN_ROLE_DETAIL_SCREEN, "System", "Services", aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show("EpiaOpenDetailScreen", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case MULTI_LANGUAGE_CHECK:
							MultiLanguageCheck(MULTI_LANGUAGE_CHECK, aeForm, out sResult);
							//System.Windows.Forms.MessageBox.Show(" MULTI_LANGUAGE_CHECK", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case SHELL_CONFIGURATION_SECURITY:
							ShellConfigSecurity(SHELL_CONFIGURATION_SECURITY, aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show(" SHELL_CONFIGURATION_SECURITY", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case LOGON_CURRENT_USER:
							LogonCurrentUser(LOGON_CURRENT_USER, aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show(" LogonCurrentUser", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case ACCOUNT_INACTIVITY_TIMEOUT_LOGOUT:
                            if (DeployUtilities.getThisPCOS().StartsWith("WindowsXP.32"))
                            {
                                InactivityTimeoutLogoutXP(ACCOUNT_INACTIVITY_TIMEOUT_LOGOUT, aeForm, out sResult);
                            }
                            else
                                InactivityTimeoutLogout(ACCOUNT_INACTIVITY_TIMEOUT_LOGOUT, aeForm, out sResult);

							System.Windows.Forms.MessageBox.Show(" ACCOUNT InactivityTimeoutLogout", "Click OK to continue", MessageBoxButtons.OK);
							break;
                        case ROLE_INACTIVITY_TIMEOUT_LOGOUT:
                            if (DeployUtilities.getThisPCOS().StartsWith("WindowsXP.32"))
                            {
                                InactivityTimeoutLogoutXP(ROLE_INACTIVITY_TIMEOUT_LOGOUT, aeForm, out sResult);
                            }
                            else
                                InactivityTimeoutLogout(ROLE_INACTIVITY_TIMEOUT_LOGOUT, aeForm, out sResult);

							System.Windows.Forms.MessageBox.Show(" ROLE InactivityTimeoutLogout", "Click OK to continue", MessageBoxButtons.OK);
							break;
                        case ACCOUNT_INACTIVITY_TIMEOUT_SHUTDOWN:
                            if (DeployUtilities.getThisPCOS().StartsWith("WindowsXP.32"))
                            {
                                InactivityTimeoutShutdownXP(ACCOUNT_INACTIVITY_TIMEOUT_SHUTDOWN, aeForm, out sResult);
                            }
                            else
                                InactivityTimeoutShutdown(ACCOUNT_INACTIVITY_TIMEOUT_SHUTDOWN, aeForm, out sResult);

							System.Windows.Forms.MessageBox.Show(" ACCOUNT InactivityTimeoutShutdown", "Click OK to continue", MessageBoxButtons.OK);
							break;
                        case ROLE_INACTIVITY_TIMEOUT_SHUTDOWN:
                            if (DeployUtilities.getThisPCOS().StartsWith("WindowsXP.32"))
                            {
                                InactivityTimeoutShutdownXP(ROLE_INACTIVITY_TIMEOUT_SHUTDOWN, aeForm, out sResult);
                            }
                            else
                                InactivityTimeoutShutdown(ROLE_INACTIVITY_TIMEOUT_SHUTDOWN, aeForm, out sResult);

							System.Windows.Forms.MessageBox.Show(" ROLE InactivityTimeoutShutdown", "Click OK to continue", MessageBoxButtons.OK);
							break;
						
						
						case LOGON_EPIA_ADMINISTRATOR:
							LogonEpiaAdministrator(LOGON_EPIA_ADMINISTRATOR, aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show(" LOGON_EPIA_ADMINISTRATOR", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case SHELL_SHUTDOWN:
							ShellShutdown(SHELL_SHUTDOWN, aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show(" SHELL_SHUTDOWN", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case SHELL_LOGOFF:
							ShellLogoff(SHELL_LOGOFF, aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show(" SHELL_LOGOFF", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case EPIA4_CLOSEE:
							Epia4Close(EPIA4_CLOSEE, aeForm, out sResult);
							break;
                        case SHELL_CLOSEE_WITHIN_MIN_AFTER_SERVER_DOWN:
                            ShellCloseWithinOneMinuteAfterServerDown(SHELL_CLOSEE_WITHIN_MIN_AFTER_SERVER_DOWN, aeForm, out sResult);
							//System.Windows.Forms.MessageBox.Show(" SHELL_CLOSEE_WITHIN_MIN_AFTER_SERVER_DOWN", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case EPIA4_CLEAN_UNINSTALL_CHECK:
                            Epia4CleanUninstallCheck(EPIA4_CLEAN_UNINSTALL_CHECK, aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show(" Epia4CleanUninstallCheck", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case EPIA4_CLEAN_SHELL_INSTALL_CHECK:
							Epia4CleanShellInstallCheck(EPIA4_CLEAN_SHELL_INSTALL_CHECK, aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show(" EPIA4_CLEAN_SHELL_INSTALL_CHECK", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case EPIA4_CLEAN_SERVER_INSTALL_CHECK:
							Epia4CleanServerInstallCheck(EPIA4_CLEAN_SERVER_INSTALL_CHECK, aeForm, out sResult);
							System.Windows.Forms.MessageBox.Show(" EPIA4_CLEAN_SERVER_INSTALL_CHECK", "Click OK to continue", MessageBoxButtons.OK);
							break;
						case UNINSTALL_EPIA_RESOURCE_EDITOR_IF_ALREADY_INSTALLED:
							UninstallEpiaResourceFileEditorIfAlreadyInstalled(UNINSTALL_EPIA_RESOURCE_EDITOR_IF_ALREADY_INSTALLED, aeForm, out sResult);
							break;
						case INSTALL_EPIA_RESOURCE_EDITOR:
							InstallEpiaResourceFileEditor(INSTALL_EPIA_RESOURCE_EDITOR, aeForm, out sResult);
							break;
						case LOAD_EPIA_RESOURCE_FILES:
                            if (DeployUtilities.getThisPCOS().StartsWith("WindowsXP.32"))
                            {
                                LoadEpiaResourceFilesXP(LOAD_EPIA_RESOURCE_FILES, aeForm, out sResult);
                            }
                            else
							    LoadEpiaResourceFiles(LOAD_EPIA_RESOURCE_FILES, aeForm, out sResult);
							break;
						case FILTER_EPIA_RESOURCE_FILES:
							FilterResourceFiles(FILTER_EPIA_RESOURCE_FILES, aeForm, out sResult);
							break;
						case EDIT_SAVE_RESOURCE_FILES:
							EditAndSaveResourceFiles(EDIT_SAVE_RESOURCE_FILES, aeForm, out sResult);
							break;
					}

                    FileManipulation.WriteExcelTestCaseResult(xApp, sResult, sHeaderContentsLength, Counter, sTestCaseName[Counter], sErrorMessage);
                 
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

                FileManipulation.WriteExcelFoot(xApp, sHeaderContentsLength, Counter, sTotalCounter, sTotalPassed, sTotalFailed);
                Epia3Common.WriteTestLogMsg(slogFilePath, "sTotalCounter: " + sTotalCounter, sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "sTotalPassed: " + sTotalPassed, sOnlyUITest);
                Epia3Common.WriteTestLogMsg(slogFilePath, "sTotalFailed: " + sTotalFailed, sOnlyUITest);
               
				#region // save Excel to Local machine and remote machine
				string sXLSPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), sOutFilename + ".xls");
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
				else
				{
					string sXLSPath2 = System.IO.Path.Combine(sOutFilePath, sOutFilename + ".xls");
					Console.WriteLine("Save2 : " + sXLSPath2);
					Epia3Common.WriteTestLogMsg(slogFilePath, "sXLSPath2 =: " + sXLSPath2, sOnlyUITest);
                    if (FileManipulation.SaveExcel(xApp, sXLSPath2, ref sErrorMessage) == false)
                    {
                        string sTXTPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), sOutFilename + ".txt");
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
					TFSConnected = TfsUtilities.CheckTFSConnection( ref msgX );
					while (TFSConnected == false)
					{
						TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
								"update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay );
						System.Threading.Thread.Sleep( Tfs.ReconnectDelay );
						Console.WriteLine(" Reconnect TFS Server:::::: ");
						TFSConnected = TfsUtilities.CheckTFSConnection( ref msgX );
					}

					if (TFSConnected)
					{
						// added check sTestResultFolder exist; some time during testing this build can be completely deleted by WVB
						if (Directory.Exists(sTestResultFolder))
						{
							#region  // update testinfo file first and then update build quality
							if (sAutoTest)
							{
                                Epia3Common.WriteTestLogMsg(slogFilePath, "TfsUtilities.GetBuildUriFromBuildNumber: " + sTotalFailed, sOnlyUITest);
                                string prjName = TfsUtilities.GetProjectName(TFSQATestTools.TestApp.EPIA4);
                                Epia3Common.WriteTestLogMsg(slogFilePath, "prjName: " + prjName, sOnlyUITest);
                                Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, TfsUtilities.GetProjectName(TFSQATestTools.TestApp.EPIA4), sBuildNr);
								string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;
                                Epia3Common.WriteTestLogMsg(slogFilePath, "m_BuildSvc.GetMinimalBuildDetails(uri).Quality: " + quality, sOnlyUITest);
								if (sTotalFailed == 0)
								{
                                    if (TestListUtilities.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Passed", "Tests OK", sInfoFileKey) == false)
                                    {
                                        // build is deleted by Wim, exit this app 
                                        System.Environment.Exit(0);
                                    }
                                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Passed:" + TFSQATestTools.TestApp.EPIA4, sOnlyUITest);

									Console.WriteLine(" Update build quality:  quality: " + quality);
									if (quality.Equals("GUI Tests Failed"))
									{
										Console.WriteLine("Quality is GUI Tests Failed : No Update Needed ");
										Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
									}
                                    else if (TestListUtilities.IsAllTestDefinitionsTested(mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage) == false)
                                    //else if ( TestListUtilities.IsAllTestDefinitionsTested( mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage ) == false )
									{
										Console.WriteLine("NOT All Test definitions tested " + sErrorMessage);
									}
									else
									{
										if ( TestListUtilities.IsAllTestStatusPassed( mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage ) == true )
										{
											// update quality to GUI Tests Passed
                                            TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(TFSQATestTools.TestApp.EPIA4),
											//"GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
                                            "GUI Tests Passed", m_BuildSvc, sDemo ? "false" : "false");   // only for demo

											Console.WriteLine("update quality to true -----  ");
											Thread.Sleep(1000);
										}
									}
								}
								else
								{
									TestListUtilities.UpdateStatusInTestInfoFile( sTestResultFolder, "GUI Tests Failed", "--->" + sOutFilename + ".log", sInfoFileKey );
                                    Epia3Common.WriteTestLogMsg(slogFilePath, " sInfoFileKey:" + sInfoFileKey + " GUI Tests Failed:" + TFSQATestTools.TestApp.EPIA4, sOnlyUITest);

									Console.WriteLine(" Update build quality:  quality: " + quality);
									if (quality.Equals("GUI Tests Failed"))
									{
										Console.WriteLine("Quality is GUI Tests Failed : No Update Needed ");
										Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
									}
									else
									{
										// update quality to GUI Tests Passed
                                        TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(TFSQATestTools.TestApp.EPIA4),
										//"GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
                                        "GUI Tests Failed", m_BuildSvc, sDemo ? "false" : "false");   // only for demo

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
				
				// close CommandHost
				Thread.Sleep(10000);
				ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SHELL );
				Thread.Sleep(10000);
				ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SERVER );
				Thread.Sleep(5000);
				ProcessUtilities.CloseProcess( "EXCEL" );
				TestTools.ProcessUtilities.CloseProcess( "cmd" );
				Console.WriteLine("\nEnd test run\n");
				//Console.ReadLine();
			}
			catch (Exception ex)
			{
				Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace , sOnlyUITest);
				Thread.Sleep(2000);
				if (sAutoTest)
				{
                    try
                    {                   
					    #region // test exception : update infofile and build quality
				        TestListUtilities.UpdateStatusInTestInfoFile( sTestResultFolder, "GUI Tests Exception", " -->" + sOutFilename + ".log", sInfoFileKey );
                        Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Exception -->" + sOutFilename + ".log:" + TFSQATestTools.TestApp.EPIA4, sOnlyUITest);
				        string msgX = "epia exception build quality test status to Failed if needed";
                        Epia3Common.WriteTestLogMsg(slogFilePath, "=====> :" + msgX, sOnlyUITest);
				        TFSConnected = TfsUtilities.CheckTFSConnection( ref msgX );
				        while (TFSConnected == false)
				        {
					        TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
							        "update build quality Deployment Failed2", (uint)Tfs.ReconnectDelay );
					        System.Threading.Thread.Sleep( Tfs.ReconnectDelay );
					        TFSConnected = TfsUtilities.CheckTFSConnection( ref msgX );
				        }

				        if (TFSConnected)
				        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, "2TfsUtilities.GetBuildUriFromBuildNumber: " + sTotalFailed, sOnlyUITest);
                            string prjName = TfsUtilities.GetProjectName(TFSQATestTools.TestApp.EPIA4);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "2prjName: " + prjName, sOnlyUITest);

                            Uri uri = TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc, TfsUtilities.GetProjectName(TFSQATestTools.TestApp.EPIA4), sBuildNr);
					        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

					        if (quality.Equals("GUI Tests Failed"))
					        {
						        Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
					        }
					        else
					        {
                                TfsUtilities.UpdateBuildQualityStatus(logger, uri, TfsUtilities.GetProjectName(TFSQATestTools.TestApp.EPIA4),
							    //    "GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
                                 "GUI Tests Failed", m_BuildSvc, sDemo ? "false" : "false");   // only for demo
					        }
				        }
					    #endregion
                    }
                    catch (Exception exc)
                    {
                        Epia3Common.WriteTestLogMsg(slogFilePath, "END Exception :" + exc.Message+" --- "+ exc.StackTrace, sOnlyUITest);
                        //throw;
                    }
				}
            }
            finally
            {
                ProcessUtilities.CloseProcess("DW20");
				ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SHELL );
				ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SERVER );
				ProcessUtilities.CloseProcess( "cmd" );
				ProcessUtilities.CloseProcess( "EXCEL" );
                System.Environment.Exit(1);
			}
		}
		#endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region TestCase Name
		private const string EPIA_SERVER_INSTALLATION_INTEGRITY = "EpiaServerInstallerIntegrityCheck";
		private const string EPIA_SHELL_INSTALLATION_INTEGRITY = "EpiaShellInstallerIntegrityCheck";
		private const string START_EPIA_SERVER_SHELL = "StartEpiaServerShell";
		private const string RESOURCEID_INTEGRITY_CHECK = "ResourceIDIntegrityCheck";
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
		private const string SECURITY_OPEN_ACCOUNT_DETAIL_SCREEN = "OpenAccountDetailScreen";
		private const string SECURITY_ADD_NEW_ROLE = "AddNewRole";
		private const string SECURITY_OPEN_ROLE_DETAIL_SCREEN = "OpenRoleDetailScreen";
		private const string SECURITY_EDIT_ROLE = "EditRole";
        private const string SECURITY_ADD_NEW_ACCOUNT = "AddNewAccount";
        private const string SYSTEM_ADD_SERVICE = "AddService";
        private const string SYSTEM_OPEN_SERVICE_DETAIL_SCREEN = "OpenServiceDetailScreen";
        private const string ROLE_INACTIVITY_TIMEOUT = "RoleInactivityTimeout";
        private const string ACCOUNT_INACTIVITY_TIMEOUT_LOGOUT = "AccountInactivityTimeoutLogout";
        private const string ACCOUNT_INACTIVITY_TIMEOUT_SHUTDOWN = "AccountInactivityTimeoutShutdown";
        private const string ROLE_INACTIVITY_TIMEOUT_LOGOUT = "RoleInactivityTimeoutLogout";
        private const string ROLE_INACTIVITY_TIMEOUT_SHUTDOWN = "RoleInactivityTimeoutShutdown";
		private const string MULTI_LANGUAGE_CHECK = "MultiLanguageCheck";
		private const string SHELL_CONFIGURATION_SECURITY = "ShellConfigSecurity";
		private const string LOGON_CURRENT_USER = "LogonCurrentUser";
		private const string LOGON_EPIA_ADMINISTRATOR = "LogonEpiaAdmin";
		private const string SHELL_SHUTDOWN = "ShellShutdown";
		private const string SHELL_LOGOFF = "ShellLogOff";
		private const string EPIA4_CLOSEE = "Epia4Close";
        private const string SHELL_CLOSEE_WITHIN_MIN_AFTER_SERVER_DOWN = "ShellCloseWithinMinuteAfterServerDown";
		private const string EPIA4_CLEAN_UNINSTALL_CHECK = "Epia4CleanUnstallCheck";
		private const string EPIA4_CLEAN_SHELL_INSTALL_CHECK = "Epia4CleanShellInstallCheck";
		private const string EPIA4_CLEAN_SERVER_INSTALL_CHECK = "Epia4CleanServerInstallCheck";
		private const string UNINSTALL_EPIA_RESOURCE_EDITOR_IF_ALREADY_INSTALLED = "UninstallEpiaResourceFileEditorIfAlreadyInstalled";
		private const string INSTALL_EPIA_RESOURCE_EDITOR = "installEpiaResourceFileEditor";
		private const string LOAD_EPIA_RESOURCE_FILES = "LoadEpiaResourceFiles";
		private const string FILTER_EPIA_RESOURCE_FILES = "FilterEpiaResourceFiles";
		private const string EDIT_SAVE_RESOURCE_FILES = "EditAndSaveResourceFiles";

		//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		private const string LAYOUT_NAVIGATOR_OFF = "LayoutNavigatorOff";
		#endregion TestCase Name
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Test Cases -------------------------------------------------------------------------------------------
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region EpiaServerIntegrityCheck
		public static void EpiaServerIntegrityCheck(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			try
			{
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 5, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }
                
				//find the dll in the Server folder
				//DirectoryInfo DirInfo = new DirectoryInfo(sEpiaServerFolder);
				//FileInfo[] installedDlls = DirInfo.GetFiles("*.dll");

				//StreamWriter writeInfo = File.CreateText(System.IO.Path.Combine(Directory.GetCurrentDirectory(), "EpiaServerFolderNET4Files.txt"));
				//writeInfo.WriteLine(installedDllsName[i]);
				//writeInfo.Close();

				string[] installedDllsName = EpiaUtilities.GetFiles(sEpiaServerFolder, "*.exe;*.config;*.dll;*.pdb");
				//string[] installedDllsName = new string[installedDlls.Length];
				string dlls = string.Empty;
				for (int i = 0; i < installedDllsName.Length; i++)
				{
					//installedDllsName[i] = installedDlls[i].Name;
					Console.WriteLine("installedDllsName[i] : " + installedDllsName[i]);
					dlls = dlls + System.Environment.NewLine + installedDllsName[i];
				}

				// already converted to DOTNET 4 build already 
				/*string[] EpiaServerDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaServerFolderFiles.txt"));
				if (sInstallMsiDir.IndexOf("Net4") > 0
					|| sInstallMsiDir.IndexOf("Dev01") > 0
					|| sInstallMsiDir.IndexOf("Dev02") > 0
					 || sInstallMsiDir.IndexOf("Dev03.") > 0
					 || sInstallMsiDir.IndexOf("Dev05") > 0
					   || sInstallMsiDir.IndexOf("Dev08.") > 0
					|| sInstallMsiDir.IndexOf("Main") > 0
					 || sInstallMsiDir.IndexOf("Production") > 0
					)*/
				string[] EpiaServerDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaServerFolderNET4Files.txt"));

                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2013, 2, 26), ref sErrorMessage) == true)
                    EpiaServerDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaServerFolderNET4Files.txt"));
                else
                    EpiaServerDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaServerFolderNET4Files20130226.txt"));

				// already implemented Service function for all
				if (sInstallMsiDir.IndexOf("Dev01") > 0
					|| sInstallMsiDir.IndexOf("Dev02") > 0
				     || sInstallMsiDir.IndexOf("Dev03") > 0
					|| sInstallMsiDir.IndexOf("Dev05") > 0
					|| sInstallMsiDir.IndexOf("Dev08") > 0
				//	 || sInstallMsiDir.IndexOf("Net4") > 0
				   || sInstallMsiDir.IndexOf("Production") > 0
				   || sInstallMsiDir.IndexOf("Main") > 0
					)
				{

					List<string> NewList = new List<string>();
					for (int i = 0; i < EpiaServerDlls.Length; i++)
					{
						NewList.Add(EpiaServerDlls[i]);
					}
                    NewList.Add("Microsoft.ReportViewer.WinForms.dll");
					EpiaServerDlls = NewList.ToArray();

					dlls = string.Empty;
					for (int i = 0; i < EpiaServerDlls.Length; i++)
					{
						dlls = dlls + System.Environment.NewLine + EpiaServerDlls[i];
					}
					//System.Windows.Forms.MessageBox.Show(dlls, "installedServerDllCnt: " + EpiaServerDlls.Length);
				}
                

				// before compare, For release version should exclude .pdb files
				if (sInstallMsiDir.IndexOf("Release") > 0
					&& sInstallMsiDir.IndexOf("Debug") < 0
					&& sInstallMsiDir.IndexOf("Protect") < 0
					)
				{
					List<string> NewList = new List<string>();
					for (int i = 0; i < EpiaServerDlls.Length; i++)
					{
						if (EpiaServerDlls[i].IndexOf(".pdb") < 0)
							NewList.Add(EpiaServerDlls[i]);
					}
					EpiaServerDlls = NewList.ToArray();

					dlls = string.Empty;
					for (int i = 0; i < EpiaServerDlls.Length; i++)
					{
						dlls = dlls + System.Environment.NewLine + EpiaServerDlls[i];
					}
					//System.Windows.Forms.MessageBox.Show(dlls, "EpiaServerDlls.length: " + EpiaShellDlls.Length );
				}

				// compare dlls in Server folder
				int installedServerDllCnt = installedDllsName.Length;
				int standardServerDllCnt = EpiaServerDlls.Length;

				if (sOnlyUITest)
					System.Windows.Forms.MessageBox.Show("installedServerDllCnt: " + installedServerDllCnt + Environment.NewLine + dlls,
					"standardServerDllCnt: " + standardServerDllCnt);


				if (installedServerDllCnt < standardServerDllCnt)  // find missing dll
				{
					EpiaUtilities.CompareFileLists(installedDllsName, EpiaServerDlls, ref sErrorMessage);
					sErrorMessage = "Too lees files installed : missing dll --> " + sErrorMessage;
					TestCheck = ConstCommon.TEST_FAIL;
					Console.WriteLine(sErrorMessage);
				}
				else if (installedServerDllCnt > standardServerDllCnt)  // find unnecessary dll
				{
					EpiaUtilities.CompareFileLists(EpiaServerDlls, installedDllsName, ref sErrorMessage);
					sErrorMessage = "Too much files installed : extra installed dll --> " + sErrorMessage;
					TestCheck = ConstCommon.TEST_FAIL;
					Console.WriteLine(sErrorMessage);
				}
				else
				{
					sErrorMessage = string.Empty;
					EpiaUtilities.CompareFileLists(EpiaServerDlls, installedDllsName, ref sErrorMessage);
					if (sErrorMessage.Length > 2)
					{
						TestCheck = ConstCommon.TEST_FAIL;
						sErrorMessage = "Different files are installed : differ dll --> " + sErrorMessage;
						Console.WriteLine(sErrorMessage);
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
				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
                result = ConstCommon.TEST_EXCEPTION;
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
        #region EpiaServerIntegrityCheckNet45
        public static void EpiaServerIntegrityCheckNet45(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 5, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                //find the dll in the Server folder
                //DirectoryInfo DirInfo = new DirectoryInfo(sEpiaServerFolder);
                //FileInfo[] installedDlls = DirInfo.GetFiles("*.dll");

                //StreamWriter writeInfo = File.CreateText(System.IO.Path.Combine(Directory.GetCurrentDirectory(), "EpiaServerFolderNET4Files.txt"));
                //writeInfo.WriteLine(installedDllsName[i]);
                //writeInfo.Close();

                string[] installedDllsName = EpiaUtilities.GetFiles(sEpiaServerFolder, "*.exe;*.config;*.dll;*.pdb");
                //string[] installedDllsName = new string[installedDlls.Length];
                string dlls = string.Empty;
                for (int i = 0; i < installedDllsName.Length; i++)
                {
                    //installedDllsName[i] = installedDlls[i].Name;
                    Console.WriteLine("installedDllsName[i] : " + installedDllsName[i]);
                    dlls = dlls + System.Environment.NewLine + installedDllsName[i];
                }

                // already converted to DOTNET 4 build already 
                /*string[] EpiaServerDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaServerFolderFiles.txt"));
                if (sInstallMsiDir.IndexOf("Net4") > 0
                    || sInstallMsiDir.IndexOf("Dev01") > 0
                    || sInstallMsiDir.IndexOf("Dev02") > 0
                     || sInstallMsiDir.IndexOf("Dev03.") > 0
                     || sInstallMsiDir.IndexOf("Dev05") > 0
                       || sInstallMsiDir.IndexOf("Dev08.") > 0
                    || sInstallMsiDir.IndexOf("Main") > 0
                     || sInstallMsiDir.IndexOf("Production") > 0
                    )*/
                string[] EpiaServerDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaServerFolderNET4Files.txt"));

                // already implemented Service function for all
                //if (sInstallMsiDir.IndexOf("Dev03") > 0
                    //	|| sInstallMsiDir.IndexOf("Dev01") > 0
                    //  || sInstallMsiDir.IndexOf("Dev03.") > 0
                    //	|| sInstallMsiDir.IndexOf("Dev05") > 0
                    //	|| sInstallMsiDir.IndexOf("Dev08") > 0
                    //	 || sInstallMsiDir.IndexOf("Net4") > 0
                    //   || sInstallMsiDir.IndexOf("Production") > 0
                    //   || sInstallMsiDir.IndexOf("Main") > 0
               //     )
               // {

                    List<string> NewList = new List<string>();
                    for (int i = 0; i < EpiaServerDlls.Length; i++)
                    {
                        NewList.Add(EpiaServerDlls[i]);
                    }
                    //NewList.Add("Microsoft.ReportViewer.WinForms.dll");
                    NewList.Add("Microsoft.Practices.Unity.dll");
                    NewList.Add("Microsoft.VisualStudio.ComponentModelHost.dll");
                    NewList.Add("Microsoft.VisualStudio.GraphModel.dll");

                    NewList.Add("Microsoft.VisualStudio.Shell.11.0.dll");
                    NewList.Add("Microsoft.VisualStudio.Shell.Immutable.10.0.dll");
                    NewList.Add("Microsoft.VisualStudio.Shell.Immutable.11.0.dll");
                    NewList.Add("System.ComponentModel.Composition.dll");
                    NewList.Add("System.Web.ApplicationServices.dll");
                    NewList.Add("System.Xaml.dll");
                    
                    EpiaServerDlls = NewList.ToArray();

                    dlls = string.Empty;
                    for (int i = 0; i < EpiaServerDlls.Length; i++)
                    {
                        dlls = dlls + System.Environment.NewLine + EpiaServerDlls[i];
                    }
                    //System.Windows.Forms.MessageBox.Show(dlls, "installedServerDllCnt: " + EpiaServerDlls.Length);
             //   }


                // before compare, For release version should exclude .pdb files
                //if (sInstallMsiDir.IndexOf("Release") > 0
               //     && sInstallMsiDir.IndexOf("Debug") < 0
               //     && sInstallMsiDir.IndexOf("Protect") < 0
               //     )
               // {
                    List<string> NewListX = new List<string>();
                    for (int i = 0; i < EpiaServerDlls.Length; i++)
                    {
                        if (EpiaServerDlls[i].IndexOf(".pdb") < 0)
                            NewListX.Add(EpiaServerDlls[i]);
                    }
                    EpiaServerDlls = NewListX.ToArray();

                    dlls = string.Empty;
                    for (int i = 0; i < EpiaServerDlls.Length; i++)
                    {
                        dlls = dlls + System.Environment.NewLine + EpiaServerDlls[i];
                    }
                    //System.Windows.Forms.MessageBox.Show(dlls, "EpiaServerDlls.length: " + EpiaShellDlls.Length );
               // }

                // compare dlls in Server folder
                int installedServerDllCnt = installedDllsName.Length;
                int standardServerDllCnt = EpiaServerDlls.Length;

                if (sOnlyUITest)
                    System.Windows.Forms.MessageBox.Show("installedServerDllCnt: " + installedServerDllCnt + Environment.NewLine + dlls,
                    "standardServerDllCnt: " + standardServerDllCnt);


                if (installedServerDllCnt < standardServerDllCnt)  // find missing dll
                {
                    EpiaUtilities.CompareFileLists(installedDllsName, EpiaServerDlls, ref sErrorMessage);
                    sErrorMessage = "Too lees files installed : missing dll --> " + sErrorMessage;
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                }
                else if (installedServerDllCnt > standardServerDllCnt)  // find unnecessary dll
                {
                    EpiaUtilities.CompareFileLists(EpiaServerDlls, installedDllsName, ref sErrorMessage);
                    sErrorMessage = "Too much files installed : extra installed dll --> " + sErrorMessage;
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                }
                else
                {
                    sErrorMessage = string.Empty;
                    EpiaUtilities.CompareFileLists(EpiaServerDlls, installedDllsName, ref sErrorMessage);
                    if (sErrorMessage.Length > 2)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "Different files are installed : differ dll --> " + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
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
                Thread.Sleep(2000);
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
		#region EpiaShellIntegrityCheck
		public static void EpiaShellIntegrityCheck(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			try
			{
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 5, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

				//find the dll in the Server folder
				//DirectoryInfo DirInfo = new DirectoryInfo(sEpiaShellFolder);
				//FileInfo[] installedDlls = DirInfo.GetFiles("*.dll");
				//string[] installedDllsName = new string[installedDlls.Length];
				string[] installedDllsName = EpiaUtilities.GetFiles(sEpiaShellFolder, "*.exe;*.config;*.dll;*.pdb");
				string dlls = string.Empty;

				//StreamWriter writeInfo = File.CreateText(System.IO.Path.Combine(Directory.GetCurrentDirectory(), "EpiaShellFolderNET4Files.txt"));
				//writeInfo.WriteLine(installedDllsName[i]);
				// writeInfo.Close();
				for (int i = 0; i < installedDllsName.Length; i++)
				{
					//installedDllsName[i] = installedDllsName[i].Name;
					Console.WriteLine("installedDllsName[i] : " + installedDllsName[i]);
					dlls = dlls + System.Environment.NewLine + installedDllsName[i];

				}

				// already converted to DOTNET 4
				/*string[] EpiaShellDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaShellFolderFiles.txt"));
				if ( sInstallMsiDir.IndexOf("Dev01") > 0
					  || sInstallMsiDir.IndexOf("Dev02") > 0
					 || sInstallMsiDir.IndexOf("Dev03.") > 0
					 || sInstallMsiDir.IndexOf("Dev03-Net4") > 0
					 || sInstallMsiDir.IndexOf("Dev05") > 0
					 || sInstallMsiDir.IndexOf("Dev08") > 0
					  || sInstallMsiDir.IndexOf("Main") > 0
					  || sInstallMsiDir.IndexOf("Production") > 0
					)*/
				//{
				string[] EpiaShellDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaShellFolderNET4Files.txt"));
				//}


                // already implemented DocumentFormat.OpenXml
                if (sInstallMsiDir.IndexOf("Dev02") > 0 && sInstallMsiDir.IndexOf("Dev02") > 0
			//		|| sInstallMsiDir.IndexOf("Dev02") > 0
			//		 || sInstallMsiDir.IndexOf("Dev03.") > 0
			//		  || sInstallMsiDir.IndexOf("Dev03-Net4") > 0
				   || sInstallMsiDir.IndexOf("Dev05") > 0
			//	   || sInstallMsiDir.IndexOf("Dev08") > 0
				   || sInstallMsiDir.IndexOf("Production") > 0
				   || sInstallMsiDir.IndexOf("Main") > 0
				   )
				{
					List<string> NewList = new List<string>();
					for (int i = 0; i < EpiaShellDlls.Length; i++)
					{
						NewList.Add(EpiaShellDlls[i]);
					}

                    NewList.Add("DocumentFormat.OpenXml.dll");
                    NewList.Add("System.Reactive.Core.dll");
                    NewList.Add("System.Reactive.Interfaces.dll");
                    NewList.Add("System.Reactive.Linq.dll");
                    NewList.Add("System.Reactive.Windows.Threading.dll");
					EpiaShellDlls = NewList.ToArray();

					dlls = string.Empty;
					for (int i = 0; i < EpiaShellDlls.Length; i++)
					{
						dlls = dlls + System.Environment.NewLine + EpiaShellDlls[i];
					}
					//System.Windows.Forms.MessageBox.Show(dlls, "EpiaShellDlls.length: " + EpiaShellDlls.Length );
				}

                if (sInstallMsiDir.IndexOf("Dev03") > 0 && sInstallMsiDir.IndexOf("Dev03") > 0
                    //		|| sInstallMsiDir.IndexOf("Dev02") > 0
                    //		 || sInstallMsiDir.IndexOf("Dev03.") > 0
                    //		  || sInstallMsiDir.IndexOf("Dev03-Net4") > 0
                    	   || sInstallMsiDir.IndexOf("Dev05") > 0
                    //	   || sInstallMsiDir.IndexOf("Dev08") > 0
                    	   || sInstallMsiDir.IndexOf("Production") > 0
                    	   || sInstallMsiDir.IndexOf("Main") > 0
                   )
                {
                    List<string> NewList = new List<string>();
                    for (int i = 0; i < EpiaShellDlls.Length; i++)
                    {
                        NewList.Add(EpiaShellDlls[i]);
                    }

                    NewList.Add("InfragisticsWPF4.Controls.Editors.XamCalendar.v11.2.dll");
                    //NewList.Add("InfragisticsWPF4.Controls.Editors.XamCalendar.v11.2.dll");
                    NewList.Add("InfragisticsWPF4.Controls.Editors.XamDateTimeInput.v11.2.dll");
                    NewList.Add("InfragisticsWPF4.Controls.Editors.XamMaskedInput.v11.2.dll");
                    NewList.Add("InfragisticsWPF4.v11.2.dll");
                    EpiaShellDlls = NewList.ToArray();

                    dlls = string.Empty;
                    for (int i = 0; i < EpiaShellDlls.Length; i++)
                    {
                        dlls = dlls + System.Environment.NewLine + EpiaShellDlls[i];
                    }
                    //System.Windows.Forms.MessageBox.Show(dlls, "EpiaShellDlls.length: " + EpiaShellDlls.Length );
                }

				// obsolete Infragistics4.Win.UltraWinExplorerBar.v11.2.dll
				if (sInstallMsiDir.IndexOf("Dev08") > 0
					|| sInstallMsiDir.IndexOf("Dev01") > 0
					  || sInstallMsiDir.IndexOf("Dev02") > 0
					|| sInstallMsiDir.IndexOf("Dev03") > 0
                    || sInstallMsiDir.IndexOf("Dev05") > 0
					 || sInstallMsiDir.IndexOf("Main") > 0
                     || sInstallMsiDir.IndexOf("Production") > 0
				   )
				{
					List<string> NewList = new List<string>();
					for (int i = 0; i < EpiaShellDlls.Length; i++)
					{
						if (EpiaShellDlls[i].IndexOf("Win.UltraWinExplorerBar") < 0)
							NewList.Add(EpiaShellDlls[i]);
					}
					EpiaShellDlls = NewList.ToArray();

                   
				}
                 
				// before compare, For release version should exclude .pdb files
				if ( sInstallMsiDir.IndexOf("Release") > 0
						&& sInstallMsiDir.IndexOf("Debug") < 0
						&& sInstallMsiDir.IndexOf("Protect") < 0
					)
				{
					List<string> NewList = new List<string>();
					for (int i = 0; i < EpiaShellDlls.Length; i++)
					{
						if (EpiaShellDlls[i].IndexOf(".pdb") < 0)
							NewList.Add(EpiaShellDlls[i]);
					}
					EpiaShellDlls = NewList.ToArray();

					dlls = string.Empty;
					for (int i = 0; i < EpiaShellDlls.Length; i++)
					{
						dlls = dlls + System.Environment.NewLine + EpiaShellDlls[i];
					}

                   
					//System.Windows.Forms.MessageBox.Show(dlls, "EpiaShellDlls.length: " + EpiaShellDlls.Length );
				}
                // -------------------------------
				// compare dlls in Shell folder
                //-------------------------------------
                string StandardShellDlls = string.Empty;
                for (int i = 0; i < EpiaShellDlls.Length; i++)
                {
                    StandardShellDlls = StandardShellDlls + "\t<"+i+">" + EpiaShellDlls[i];
                }
                if (System.Configuration.ConfigurationManager.AppSettings.Get("MsgDebug").ToLower().Equals("true"))
                    System.Windows.Forms.MessageBox.Show(StandardShellDlls, " EpiaShellDlls.length: " + EpiaShellDlls.Length);

                string InstalledDlls = string.Empty;
				for (int i = 0; i < installedDllsName.Length; i++)
				{
                    InstalledDlls = InstalledDlls + "\t<" + i + ">" + installedDllsName[i];
				}
                if ( System.Configuration.ConfigurationManager.AppSettings.Get("MsgDebug").ToLower().Equals("true"))
                    System.Windows.Forms.MessageBox.Show(InstalledDlls, " installedDllsName.length: " + installedDllsName.Length);

				int installedShellDllCnt = installedDllsName.Length;
				int standardShellDllCnt = EpiaShellDlls.Length;
				if (installedShellDllCnt < standardShellDllCnt)  // find missing dll
				{
					EpiaUtilities.CompareFileLists(installedDllsName, EpiaShellDlls, ref sErrorMessage);
					sErrorMessage = "Too lees files installed : missing dll --> " + sErrorMessage;
					TestCheck = ConstCommon.TEST_FAIL;
					Console.WriteLine(sErrorMessage);
				}
				else if (installedShellDllCnt > standardShellDllCnt)  // find unnecessary dll
				{
                    EpiaUtilities.CompareFileLists(EpiaShellDlls, installedDllsName, ref sErrorMessage);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "installedShellDllCnt > standardShellDllCnt:" + sErrorMessage, sOnlyUITest);
                    //Too much files installed : extra installed dll --> ; Microsoft.VisualStudio.Shell.Design.dll
                    if (sErrorMessage.IndexOf("Microsoft.VisualStudio.Shell.Design.dll") >= 0 
                     || sErrorMessage.IndexOf("stdole.dll") >= 0)
                    {
                        // THIS CAUSED BY BUILD AGENT, NOT CONSIDERED AS ERROR -->  see email Wim  18-10-2012
                        // Afhankleijk van de build agent wordt deze file al of niet automatisch toegevoegd aan de msi-file.
                        // Gelieve deze file in je testen te negeren.
                        //if (sInstallMsiDir.IndexOf("CI") < 0)
                        //{
                            sMicrosoftVisualStudioShellDesignDllInstalled = true;
                        //}
                    }
                    else
                    {
				        sErrorMessage = "Too much files installed : extra installed dll --> " + sErrorMessage;
				        TestCheck = ConstCommon.TEST_FAIL;
				        Console.WriteLine(sErrorMessage);
                    }
				}
				else
				{
					sErrorMessage = string.Empty;
					EpiaUtilities.CompareFileLists(EpiaShellDlls, installedDllsName, ref sErrorMessage);
					if (sErrorMessage.Length > 2)
					{
						TestCheck = ConstCommon.TEST_FAIL;
						sErrorMessage = "Different files are installed : differ dll --> " + sErrorMessage;
						Console.WriteLine(sErrorMessage);
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

				Thread.Sleep(10000);

			}
			catch (Exception ex)
			{
                result = ConstCommon.TEST_EXCEPTION;
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
		#region StartEpiaServerShell
		public static void StartEpiaServerShell(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			try
			{
				//========================   SERVER =================================================
                if (ProjServerOrShellStartup.ServerStartup(sBrand, "Epia Server", sServerRunAs, ref sErrorMessage, slogFilePath, sOnlyUITest) == false)
                {
                    sEpiaServerStartupOK = false;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

				//========================   SHELL =================================================
                AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("EPIA SERVER Service Started : ");
					// Add Open window Event Handler
					Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					  AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);

					sEventEnd = false;
					#region  Shell
                    sErrorMessage = string.Empty;
					ProcessUtilities.StartProcessNoWait(OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand+ @"\Epia Shell",
						ConstCommon.EGEMIN_EPIA_SHELL_EXE, string.Empty);

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
						   AutomationElement.RootElement, UIAShellEventHandler);

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
                    aeForm = ProjBasicUI.GetMainWindowWithinTime("MainForm", 300);
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
					sEpiaServerStartupOK = false;
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
                result = ConstCommon.TEST_EXCEPTION;
				sEpiaServerStartupOK = false;
				sErrorMessage = ex.Message + "  ---  " + ex.StackTrace;
				Console.WriteLine(testname + " === " + sErrorMessage);
				Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);

				//System.Windows.Forms.MessageBox.Show("where ");
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
		#region ResourceIdIntegrityCheck
		public static void ResourceIdIntegrityCheck(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}

			try
			{
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 5, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

				string shellServiceLogFile = Path.Combine(OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand + @"\Epia Server\Log", "ShellServices.log");
				string shellServiceDestLogFile = Path.Combine(OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand + @"\Epia Server\Log", "ShellServicesDest.log");

				Console.WriteLine("shellServiceLogFile:  " + shellServiceLogFile);
				if (File.Exists(shellServiceLogFile))
				{
					System.IO.File.Copy(shellServiceLogFile, shellServiceDestLogFile, true);
					string[] loglines = System.IO.File.ReadAllLines(shellServiceDestLogFile);

					for (int i = 0; i < loglines.Length; i++)
					{
						//Console.WriteLine("loglines[i]:  " + loglines[i]);
						if (loglines[i].IndexOf("Warning") > 0 && loglines[i].IndexOf("ResourceId") > 0)
						{
							sErrorMessage = "-------- resource file warning log  --> " + loglines[i];
							TestCheck = ConstCommon.TEST_FAIL;
							Console.WriteLine(sErrorMessage);
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
                result = ConstCommon.TEST_EXCEPTION;
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
		#region LayoutFindLayoutPanel
		public static void LayoutFindLayoutPanel(string testname, AutomationElement root, out int result)
		{
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

			AutomationEventHandler UIAFindLayoutPanelEventHandler = new AutomationEventHandler(OnFindLayoutPanelUIAEvent);
			try
			{
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 9, 10), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                if (sEpiaServerStartupOK == false)
                {
                    sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                    return;
                }

				if (sOnlyUITest)
					root = EpiaUtilities.GetMainWindow("MainForm");

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
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);
				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					   AutomationElement.RootElement,
					  UIAFindLayoutPanelEventHandler);

				if (mTime.TotalSeconds > 600)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "After 10 min, Test is still running";
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					return;
				}

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
                result = ConstCommon.TEST_EXCEPTION;
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

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}

			try
			{
				// Add Open MyLayoutScreen window Event Handler
				Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					AutomationElement.RootElement, TreeScope.Descendants, UIALayoutXPosEventHandler);

				int k = 0;
				while (k < 5)
				{
					#region
                    root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
					AutomationElement aeYourLayouts = AUICommon.FindTreeViewNodeLevel1(testname, root, ConstCommon.MY_PLACE,
						ConstCommon.MY_LAYOUT, ref sErrorMessage);
					if (aeYourLayouts == null)
					{
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					sEventEnd = false;
					Point Pnt = AUIUtilities.GetElementCenterPoint(aeYourLayouts);
					Input.MoveToAndClick(Pnt);
					Console.WriteLine("click my layout :");
					Thread.Sleep(5000);
					while (root.Current.IsEnabled)
					{
						Console.WriteLine("re click my layout :");
						Input.MoveToAndClick(Pnt);
						Thread.Sleep(5000);
					}

					DateTime mStartTime = DateTime.Now;
					TimeSpan mTime = DateTime.Now - mStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalSeconds);
					while (sEventEnd == false && mTime.TotalSeconds <= 600)
					{
						Thread.Sleep(2000);
						mTime = DateTime.Now - mStartTime;
                        Console.WriteLine("wait time is :" + mTime.TotalSeconds);
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
						double leftValue = root.Current.BoundingRectangle.Left;
						Console.WriteLine("Current UI left value " + leftValue);

						if (Math.Abs(leftValue - 200) < 3)
						{
							k = 10;
							Console.WriteLine("left value near 200");
							Console.WriteLine("\nTest scenario Check Initial X position: Pass");
							result = ConstCommon.TEST_PASS;
							Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
						}
						else
						{
							k++;
							Console.WriteLine("Failed Xpos, should near to 200 , left value =" + leftValue);
							Console.WriteLine("\nTest scenario Check Initial X position: *FAIL*");
							result = ConstCommon.TEST_FAIL;
							Epia3Common.WriteTestLogFail(slogFilePath, testname + " k " + ": Failed Xpos, should near 200 , left value =" + leftValue, sOnlyUITest);
						}
					}
					#endregion
				}
			}
			catch (Exception ex)
			{
                result = ConstCommon.TEST_EXCEPTION;
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

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}
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
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					  AutomationElement.RootElement,
					 UIALayoutYPosEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
                result = ConstCommon.TEST_EXCEPTION;
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

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}
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
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					   AutomationElement.RootElement,
					  UIALayoutWidthEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
                result = ConstCommon.TEST_EXCEPTION;
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

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}
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
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					   AutomationElement.RootElement,
					  UIALayoutHeightEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
                result = ConstCommon.TEST_EXCEPTION;
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

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}
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
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					   AutomationElement.RootElement,
					  UIALayoutTitleEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
                result = ConstCommon.TEST_EXCEPTION;
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

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}
			try
			{
				if (sOnlyUITest)
					root = EpiaUtilities.GetMainWindow("MainForm");

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
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}
				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

				if (TestCheck == ConstCommon.TEST_FAIL)
				{
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					result = ConstCommon.TEST_FAIL;
				}

				if (TestCheck == ConstCommon.TEST_PASS)
				{
					double Width = 400;
					double Height = 700;
					TransformPattern tranform =
					root.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
					if (tranform != null)
					{
						tranform.Move(10, 10);
						Thread.Sleep(3000);
						tranform.Resize(Width, Height);
					}

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
                result = ConstCommon.TEST_EXCEPTION;
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

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}
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
				Thread.Sleep(5000);
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					  AutomationElement.RootElement,
					 UIALayoutFullScreenEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
                result = ConstCommon.TEST_EXCEPTION;
				sErrorMessage = ex.Message;
				Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
				Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
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
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region LayoutMaximized
		public static void LayoutMaximized(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			AutomationEventHandler UIALayoutMaximizedScreenEventHandler = new AutomationEventHandler(OnLayoutMaximizedScreenUIAEvent);

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}
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
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					   AutomationElement.RootElement,
					  UIALayoutMaximizedScreenEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
                result = ConstCommon.TEST_EXCEPTION;
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

            sErrorMessage = "test case removed see changeset 24065 workitem:4115 ";
            return;
            /*
			AutomationEventHandler UIALayoutRibbonOnEventHandler = new AutomationEventHandler(OnLayoutRibbonOnUIAEvent);

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}
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
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					  AutomationElement.RootElement,
					 UIALayoutRibbonOnEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
			}*/
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region LayoutCancelButton
		public static void LayoutCancelButton(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}
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
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					   AutomationElement.RootElement,
					  UIALayoutCancelButtonEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
						sErrorMessage = string.Empty;
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
		#region SecurityAddNewRole
		public static void SecurityAddNewRole(string testname, AutomationElement rootXXX, out int result)
		{
			Console.WriteLine("=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 4, 4), ref sErrorMessage) == true)
            {
                result = ConstCommon.TEST_UNDEFINED;
                return;
            }

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}

			AutomationElement aeWindow = null;
			AutomationElement aeSecurityRoles = null;
			try
			{
				aeWindow = EpiaUtilities.GetMainWindow("MainForm");
				if (aeWindow != null)
				{
					EpiaUtilities.ClearDisplayedScreens(aeWindow);
					aeSecurityRoles = AUICommon.FindTreeViewNodeLevel1(testname, aeWindow, "Security", "Roles", ref sErrorMessage);
					if (aeSecurityRoles == null)
					{
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point Pnt = AUIUtilities.GetElementCenterPoint(aeSecurityRoles);
						Input.MoveToAndClick(Pnt);
					}
				}
				else
				{
					sErrorMessage = "aeWindow not found";
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					TestCheck = ConstCommon.TEST_FAIL;
				}

                int k = 0;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
                    Thread.Sleep(5000);
					AutomationElement aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow("Roles", ref sErrorMessage);
					while (aeSelectedWindow == null && k < 5)
					{
                        Console.WriteLine("wait until selected Roles window open :" + k++);
                        aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow("Roles", ref sErrorMessage);
						Thread.Sleep(5000);
					}

					if (aeSelectedWindow == null)
					{
						Console.WriteLine("Window not opened :" + sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						// find Add...button
						System.Windows.Automation.Condition cButtonAdd = new AndCondition(
							new PropertyCondition(AutomationElement.NameProperty, "m_TxtFilter"),
							new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
						);

						AutomationElement aeButtonAdd = aeSelectedWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cButtonAdd);
						if (aeButtonAdd == null)
						{
							Console.WriteLine("aeButtonAdd not find :" + aeSelectedWindow.Current.Name);
							Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
                            double x = aeButtonAdd.Current.BoundingRectangle.Right + 70.0;
                            double y = (aeButtonAdd.Current.BoundingRectangle.Bottom + aeButtonAdd.Current.BoundingRectangle.Top) / 2.0;
                            Point pt = new Point(x, y);
                            for (int irole = 1; irole < 3; irole++)
                            {
                                Input.MoveTo(pt);
                                Input.ClickAtPoint(pt);
                                Thread.Sleep(3000);

                                if (irole == 1 && TestCheck == ConstCommon.TEST_PASS )
                                {
                                    if (EpiaUtilities.AddNewRole(slogFilePath, "RoleLogoutInOneMinute", "DescriptionA", "exitModeLogoutRadioButton", sOnlyUITest, ref sErrorMessage) == false)
                                    {
                                        Console.WriteLine(sErrorMessage);
                                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                        break;
                                    }
                                }

                                if (irole == 2 && TestCheck == ConstCommon.TEST_PASS)
                                {
                                    if (EpiaUtilities.AddNewRole(slogFilePath, "RoleShutdownInOneMinute", "DescriptionB", "exitModeShutdownRadioButton", sOnlyUITest, ref sErrorMessage) == false)
                                    {
                                        Console.WriteLine(sErrorMessage);
                                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                        break;
                                    }
                                }
                            }
						}
					}
				}

				// validate result
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					#region // Find GridView
					aeWindow = EpiaUtilities.GetMainWindow("MainForm");
					AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeWindow);
					if (aeGrid == null)
					{
						sErrorMessage = aeWindow.Current.Name + " GridData not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Console.WriteLine(aeWindow.Current.Name + " GridData found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						Thread.Sleep(3000);


                        string roleName = "RoleLogoutInOneMinute";
						// Construct the Grid Cell Element Name
                        for (int irolename = 0; irolename < 2; irolename++)
                        {
                            if (irolename == 0)
                                roleName = "RoleLogoutInOneMinute";
                            else if (irolename == 1)
                                roleName = "RoleShutdownInOneMinute";

                            string cellname = "Name" + " Row " + irolename;
                            // Get the Element with the Row Col Coordinates
                            AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                            if (aeCell == null)
                            {
                                sErrorMessage = "Find aeCell failed:" + cellname;
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                                break;
                            }
                            else
                            {
                                string cellValue = string.Empty;
                                try
                                {
                                    ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                    cellValue = vp.Current.Value;
                                    Console.WriteLine("Get element.Current Value:" + cellValue);
                                }
                                catch (System.NullReferenceException)
                                {
                                    cellValue = string.Empty;
                                }

                                if (cellValue == null || cellValue == string.Empty)
                                {
                                    sErrorMessage = "aeCell Value not found:" + cellname;
                                    Console.WriteLine(sErrorMessage);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    break;
                                }
                                else if (!cellValue.Equals(roleName))
                                {
                                    sErrorMessage = "aeCell Value not equal " +roleName+" , but :" + cellValue;
                                    Console.WriteLine(sErrorMessage);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    break;
                                }
                            }
                        }
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
                result = ConstCommon.TEST_EXCEPTION;
				sErrorMessage = ex.Message;
				Console.WriteLine(testname + " === " + sErrorMessage);
				Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
			}
			finally
			{
				Thread.Sleep(3000);
				//Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
				//	   AutomationElement.RootElement,
				//	  UIALayoutCancelButtonEventHandler);
			}
		}
		#endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region SystemAddService
        public static void SystemAddService(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (sEpiaServerStartupOK == false)
            {
                sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                return;
            }

            AutomationElement aeWindow = null;
            AutomationElement aeSystemService = null;
            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 10, 14), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                if (aeWindow != null)
                {
                    EpiaUtilities.ClearDisplayedScreens(aeWindow);
                    aeSystemService = AUICommon.FindTreeViewNodeLevel1(testname, aeWindow, "System", "Services", ref sErrorMessage);
                    if (aeSystemService == null)
                    {
                        sErrorMessage = "Services not found " + " === " + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Point Pnt = AUIUtilities.GetElementCenterPoint(aeSystemService);
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

                int k = 0;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    AutomationElement aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow("Services", ref sErrorMessage);
                    while (aeSelectedWindow == null && k < 5)
                    {
                        Console.WriteLine("wait until selected window open :" + k++);
                        aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow("Services", ref sErrorMessage);
                        Thread.Sleep(5000);
                    }

                    if (aeSelectedWindow == null)
                    {
                        Console.WriteLine("Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // find Add...button
                        //System.Windows.Automation.Condition cButtonAdd = new AndCondition(
                        //    new PropertyCondition(AutomationElement.NameProperty, "m_TxtFilter"),
                        //    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                        //);
                        //AutomationElement aeButtonAdd = aeSelectedWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cButtonAdd);
                        AutomationElement aeButtonAdd = AUIUtilities.FindElementByType(ControlType.Edit, aeSelectedWindow);
                        if (aeButtonAdd == null)
                        {
                            Console.WriteLine("aeButtonAdd not find :" + aeSelectedWindow.Current.Name);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            double x = aeButtonAdd.Current.BoundingRectangle.Right + 70.0;
                            double y = (aeButtonAdd.Current.BoundingRectangle.Bottom + aeButtonAdd.Current.BoundingRectangle.Top) / 2.0;
                            Point pt = new Point(x, y);
                            for (int iservice = 1; iservice < 2; iservice++)
                            {
                                Input.MoveTo(pt);
                                Input.ClickAtPoint(pt);
                                Thread.Sleep(3000);
                                if (iservice == 1 && TestCheck == ConstCommon.TEST_PASS)
                                {
                                    if (EpiaUtilities.AddService(slogFilePath, "Egemin Epia Server", PCName, sOnlyUITest, ref sErrorMessage) == false)
                                    {
                                        Console.WriteLine(sErrorMessage);
                                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                // validate result
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // Find GridView
                    aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                    AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeWindow);
                    if (aeGrid == null)
                    {
                        sErrorMessage = aeWindow.Current.Name + " GridData not found";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine(aeWindow.Current.Name + " GridData found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(3000);


                        string ServiceName = "Egemin Epia Server";
                        // Construct the Grid Cell Element Name
                        for (int iservicename = 0; iservicename < 1; iservicename++)
                        {
                            if (iservicename == 0)
                                ServiceName = "Egemin Epia Server";

                            string cellname = "Service" + " Row " + iservicename;
                            // Get the Element with the Row Col Coordinates
                            AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                            if (aeCell == null)
                            {
                                sErrorMessage = "Find aeCell failed:" + cellname;
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                                break;
                            }
                            else
                            {
                                string cellValue = string.Empty;
                                try
                                {
                                    ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                    cellValue = vp.Current.Value;
                                    Console.WriteLine("Get element.Current Value:" + cellValue);
                                }
                                catch (System.NullReferenceException)
                                {
                                    cellValue = string.Empty;
                                }

                                if (cellValue == null || cellValue == string.Empty)
                                {
                                    sErrorMessage = "aeCell Value not found:" + cellname;
                                    Console.WriteLine(sErrorMessage);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    break;
                                }
                                else if (!cellValue.Equals(ServiceName))
                                {
                                    sErrorMessage = "aeCell Value not equal " + ServiceName + " , but :" + cellValue;
                                    Console.WriteLine(sErrorMessage);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    break;
                                }
                            }
                        }
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
                result = ConstCommon.TEST_EXCEPTION;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                //Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                //	   AutomationElement.RootElement,
                //	  UIALayoutCancelButtonEventHandler);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region EpiaOpenDetailScreen
        public static void EpiaOpenDetailScreen(string testname, string category, string treeNode,  AutomationElement root, out int result)
		{
			Console.WriteLine("=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 4, 4), ref sErrorMessage) == true)
            {
                result = ConstCommon.TEST_UNDEFINED;
                return;
            }

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}

            string columnName = string.Empty;  // Name etc,
            string detailsWindowName = string.Empty;
			AutomationElement aeWindow = null;
			try
			{   // 
                if (treeNode.Equals("Roles"))
                {
                    columnName = "Name";
                }
                else if (treeNode.Equals("Accounts"))
                {
                    columnName = "Account name";
                }
                else if (treeNode.Equals("Services"))
                {
                    if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 10, 14), ref sErrorMessage) == true)
                    {
                        result = ConstCommon.TEST_UNDEFINED;
                        return;
                    }

                    columnName = "Service";
                }
               
				aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
				if (aeWindow != null)
				{
                    aeWindow.SetFocus();
					EpiaUtilities.ClearDisplayedScreens(aeWindow);
                    AutomationElement aeTreeViewNode = AUICommon.FindTreeViewNodeLevel1(testname, aeWindow, category, treeNode, ref sErrorMessage);
                    if (aeTreeViewNode == null)
					{
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
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
                    aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow(treeNode, ref sErrorMessage);
					if (aeSelectedWindow == null)
					{
						Console.WriteLine("Selected Window not opened :" + sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
                        if (EpiaUtilities.WindowMenuAction(aeSelectedWindow, columnName, 0, "Details...", ref sErrorMessage))
							TestCheck = ConstCommon.TEST_PASS;
						else
						{
							Console.WriteLine(sErrorMessage);
							Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
							TestCheck = ConstCommon.TEST_FAIL;
						}
					}
				}

                // Get cell name
                string cellValue = string.Empty;
                if (TestCheck == ConstCommon.TEST_PASS)
				{
                    AutomationElement aeCell = EpiaUtilities.GetCellElementFromOverviewWindow(aeSelectedWindow, columnName, 0, ref sErrorMessage);
					if (aeCell == null)
					{
						Console.WriteLine("aeCell not find :" + aeSelectedWindow.Current.Name + "---" + sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						try
						{
							ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
							cellValue = vp.Current.Value;
							Console.WriteLine("Get element.Current Value:" + cellValue);
						}
						catch (System.NullReferenceException)
						{
							cellValue = string.Empty;
						}

						if (cellValue == null || cellValue == string.Empty)
						{
							sErrorMessage = "aeCell Value not found:";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
					}
                }

                // check is details window is opend
                if (TestCheck == ConstCommon.TEST_PASS)
				{
                    if (treeNode.Equals("Roles"))
                    {
                        detailsWindowName = "Role '" + cellValue + "'";
                    }
                    else if (treeNode.Equals("Accounts"))
                    {
                        detailsWindowName = "Account '" + cellValue + "'";
                    }
                    else if (treeNode.Equals("Services"))
                    {
                        detailsWindowName = "Service detail";
                    }
                    
                    Console.WriteLine("detail window name  :" + detailsWindowName);
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
                }

				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeWindow = EpiaUtilities.GetMainWindow("MainForm");
					Console.WriteLine("aeWindow.Current.IsEnabled:   " + aeWindow.Current.IsEnabled);
					if (!aeWindow.Current.IsEnabled)
					{
						TestCheck = ConstCommon.TEST_FAIL;
						sErrorMessage = " aeWindow mainwindow is not enable --> should error screen displayed";
						Console.WriteLine(sErrorMessage);
						EpiaUtilities.TryToGetErrorMessageAndCloseErrorScreen(ref sErrorMessage);
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
                result = ConstCommon.TEST_EXCEPTION;
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
		#region SecurityEditRole
		public static void SecurityEditRole(string testname, AutomationElement rootXXX, out int result)
		{
			Console.WriteLine("=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 4, 4), ref sErrorMessage) == true)
            {
                result = ConstCommon.TEST_UNDEFINED;
                return;
            }


            if (sEpiaServerStartupOK == false)
            {
                sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                return;
            }

			AutomationElement aeWindow = null;
			AutomationElement aeSecurityRoles = null;
			try
			{
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 4, 17), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

				aeWindow = EpiaUtilities.GetMainWindow("MainForm");
				if (aeWindow != null)
				{
					EpiaUtilities.ClearDisplayedScreens(aeWindow);
					aeSecurityRoles = AUICommon.FindTreeViewNodeLevel1(testname, aeWindow, "Security", "Roles", ref sErrorMessage);
					if (aeSecurityRoles == null)
					{
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point Pnt = AUIUtilities.GetElementCenterPoint(aeSecurityRoles);
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

				int k = 0;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					AutomationElement aeSelectedWindow = EpiaUtilities.GetSelectedOverviewWindow(string.Empty, ref sErrorMessage);
					while (aeSelectedWindow == null && k < 5)
					{
						Console.WriteLine("wait until selected window open :" + k++);
						aeSelectedWindow = EpiaUtilities.GetSelectedOverviewWindow(string.Empty, ref sErrorMessage);
						Thread.Sleep(5000);
					}

					if (aeSelectedWindow == null)
					{
						Console.WriteLine("Window not opened :" + sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						// Select Rola A
						AutomationElement aeCell = EpiaUtilities.GetCellElementFromOverviewWindow(aeSelectedWindow, "Name", 0, ref sErrorMessage);
						if (aeCell == null)
						{
							Console.WriteLine("aeCell not find :" + aeSelectedWindow.Current.Name + "---" + sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeCell));
							Thread.Sleep(3000);
						}

						// find Edit...button
						System.Windows.Automation.Condition cButtonAdd = new AndCondition(
							new PropertyCondition(AutomationElement.NameProperty, "m_TxtFilter"),
							new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
						);

						if (TestCheck == ConstCommon.TEST_PASS)
						{
							AutomationElement aeButtonAdd = aeSelectedWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cButtonAdd);
							if (aeButtonAdd == null)
							{
								Console.WriteLine("aeButtonAdd not find :" + aeSelectedWindow.Current.Name);
								Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
								TestCheck = ConstCommon.TEST_FAIL;
							}
							else
							{
								Thread.Sleep(2000);
								double x = aeButtonAdd.Current.BoundingRectangle.Right + 130.0;
								double y = (aeButtonAdd.Current.BoundingRectangle.Bottom + aeButtonAdd.Current.BoundingRectangle.Top) / 2.0;
								Point pt = new Point(x, y);
								Input.MoveTo(pt);
								Input.ClickAtPoint(pt);
								Thread.Sleep(3000);
							}
						}
					}
				}

				AutomationElement aeRoleAddEditDialog = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					string RoleAddEditDialogId = "Dialog - Egemin.Epia.Modules.RnD.Screens.RoleAddEditDialog";
					aeWindow = EpiaUtilities.GetMainWindow("MainForm");
					aeRoleAddEditDialog = AUIUtilities.FindElementByID(RoleAddEditDialogId, aeWindow);
					if (aeRoleAddEditDialog == null)
					{
						Console.WriteLine("aeRoleAddEditDialog not opened :");
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						//ControlType:	"ControlType.Edit"
						//AutomationId:	"nameTextBox"
						//Name:	"Name:
						string origRole = string.Empty;
						if (TestCheck == ConstCommon.TEST_PASS)
						{
							// ControlType:	"ControlType.Edit"
							//AutomationId:	"descriptionTextBox"
							//LocalizedControlType:	"edit"
							//Name:	"Description:"

							if (AUIUtilities.FindTextBoxAndChangeValue("descriptionTextBox", aeRoleAddEditDialog, out origRole, "DescriptionB", ref sErrorMessage))
								Thread.Sleep(3000);
							else
							{
								sErrorMessage = "FindTextBoxAndChangeValue failed:" + "descriptionTextBox";
								Console.WriteLine(sErrorMessage);
								Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
								TestCheck = ConstCommon.TEST_FAIL;
							}
						}

						if (TestCheck == ConstCommon.TEST_PASS)
						{
							string inActivitySettingId = "exitConfigurationDefinedCheckbox";
							bool check = AUIUtilities.FindElementAndToggle(inActivitySettingId, aeRoleAddEditDialog, ToggleState.On);
							if (check)
								Thread.Sleep(3000);
							else
							{
								Console.WriteLine("FindElementAndToggle failed:" + inActivitySettingId);
								Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + inActivitySettingId, sOnlyUITest);
								TestCheck = ConstCommon.TEST_FAIL;
							}

							/*
							AutomationElement aeInactivityCheckBox = AUIUtilities.FindElementByID(inActivitySettingId, aeRoleAddEditDialog);
							if (aeInactivityCheckBox == null)
							 {
								 Console.WriteLine("aeInactivityCheckBox not found :");
								 Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
								 TestCheck = ConstCommon.TEST_FAIL;
							 }
							 else
							 {
							 }*/
						}


					}
					//ControlType:	"ControlType.ComboBox"
					//AutomationId:	"m_DefaultLanguageIdComboBox"
					//LocalizedControlType:	"combo box"
					//Name:	"Default language:"
				}


				if (TestCheck == ConstCommon.TEST_PASS)
				{
					string BtnSaveId = "m_btnSave";
					AutomationElement aeSave = AUIUtilities.FindElementByID(BtnSaveId, aeRoleAddEditDialog);
					if (aeSave == null)
					{
						sErrorMessage = "failed to find aeSave of aeRoleAddEditDialog";
						Console.WriteLine(testname + " failed to find aeSave at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Console.WriteLine(testname + " aeSave found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						InvokePattern ipc =
							aeSave.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
						ipc.Invoke();
					}
					Thread.Sleep(5000);
				}
				else
				{
					string BtnnCancelId = "m_btnCancel";
					AutomationElement aeCancel = AUIUtilities.FindElementByID(BtnnCancelId, aeRoleAddEditDialog);
					if (aeCancel == null)
					{
						sErrorMessage = "failed to find aeCancel of aeRoleAddEditDialog";
						Console.WriteLine(testname + " failed to find aeCancel at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						TestCheck = ConstCommon.TEST_FAIL;
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

				// check error screen displayed
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine(testname + "Edit screen Saveed  at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					aeWindow = EpiaUtilities.GetMainWindow("MainForm");
					Console.WriteLine("aeWindow.Current.IsEnabled:   " + aeWindow.Current.IsEnabled);
					if (aeWindow.Current.IsEnabled)
					{
						Console.WriteLine("aeWindow.Current.IsEnabled:   " + aeWindow.Current.IsEnabled);
						Thread.Sleep(10000);
					}
					else
					{
						TestCheck = ConstCommon.TEST_FAIL;
						sErrorMessage = " aeWindow mainwindow is not enable --> should error screen displayed";
						Console.WriteLine(sErrorMessage);
						EpiaUtilities.TryToGetErrorMessageAndCloseErrorScreen(ref sErrorMessage);
					}
				}

				// validate result
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					#region // Find GridView
					aeWindow = EpiaUtilities.GetMainWindow("MainForm");
					AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeWindow);
					if (aeGrid == null)
					{
						sErrorMessage = aeWindow.Current.Name + " GridData not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Console.WriteLine(aeWindow.Current.Name + " GridData found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						Thread.Sleep(3000);

						// Construct the Grid Cell Element Name
						string cellname = "Description" + " Row " + 0;
						// Get the Element with the Row Col Coordinates
						AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

						if (aeCell == null)
						{
							sErrorMessage = "Find aeCell failed:" + cellname;
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							string cellValue = string.Empty;
							try
							{
								ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
								cellValue = vp.Current.Value;
								Console.WriteLine("Get element.Current Value:" + cellValue);
							}
							catch (System.NullReferenceException)
							{
								cellValue = string.Empty;
							}

							if (cellValue == null || cellValue == string.Empty)
							{
								sErrorMessage = "aeCell Value not found:" + cellname;
								Console.WriteLine(sErrorMessage);
								TestCheck = ConstCommon.TEST_FAIL;
							}
							else if (!cellValue.Equals("DescriptionB"))
							{
								sErrorMessage = "aeCell Value not equal DescriptionB   , but :" + cellValue;
								Console.WriteLine(sErrorMessage);
								TestCheck = ConstCommon.TEST_FAIL;
							}
						}
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
                result = ConstCommon.TEST_EXCEPTION;
				sErrorMessage = ex.Message;
				Console.WriteLine(testname + " === " + sErrorMessage);
				Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
			}
			finally
			{
				Thread.Sleep(3000);
				//Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
				//	   AutomationElement.RootElement,
				//	  UIALayoutCancelButtonEventHandler);
			}
		}
		#endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region SecurityAddNewGeneralAccount
        public static void SecurityAddNewGeneralAccount(string testname, AutomationElement rootXXX, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;


            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 4, 4), ref sErrorMessage) == true)
            {
                result = ConstCommon.TEST_UNDEFINED;
                return;
            }

            if (sEpiaServerStartupOK == false)
            {
                sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                return;
            }

            AutomationElement aeWindow = null;
            AutomationElement aeSecurityRoles = null;
            try
            {
                aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                if (aeWindow != null)
                {
                    EpiaUtilities.ClearDisplayedScreens(aeWindow);
                    aeSecurityRoles = AUICommon.FindTreeViewNodeLevel1(testname, aeWindow, "Security", "Accounts", ref sErrorMessage);
                    if (aeSecurityRoles == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Point Pnt = AUIUtilities.GetElementCenterPoint(aeSecurityRoles);
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

                int k = 0;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    AutomationElement aeSelectedWindow = EpiaUtilities.GetSelectedOverviewWindow(string.Empty, ref sErrorMessage);
                    while (aeSelectedWindow == null && k < 5)
                    {
                        Console.WriteLine("wait until selected window open :" + k++);
                        aeSelectedWindow = EpiaUtilities.GetSelectedOverviewWindow(string.Empty, ref sErrorMessage);
                        Thread.Sleep(5000);
                    }

                    if (aeSelectedWindow == null)
                    {
                        Console.WriteLine("Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // find Add...button
                        System.Windows.Automation.Condition cButtonAdd = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "m_TxtFilter"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                        );

                        AutomationElement aeButtonAdd = aeSelectedWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cButtonAdd);
                        if (aeButtonAdd == null)
                        {
                            Console.WriteLine("aeButtonAdd not find :" + aeSelectedWindow.Current.Name);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            double x = aeButtonAdd.Current.BoundingRectangle.Right + 100.0;
                            double y = (aeButtonAdd.Current.BoundingRectangle.Bottom + aeButtonAdd.Current.BoundingRectangle.Top) / 2.0;
                            Point pt = new Point(x, y);
                            for (int irole = 1; irole < 5; irole++)
                            {
                                Input.MoveTo(pt);
                                Input.ClickAtPoint(pt);
                                Thread.Sleep(3000);
                                
                                if (irole == 1 && TestCheck == ConstCommon.TEST_PASS)
                                {
                                    if (EpiaUtilities.AddNewAccount(slogFilePath, "AccountLogoutInOneMinute", UserPassword, "DescriptionA", true, "exitModeLogoutRadioButton", 1,  string.Empty, 
                                        sOnlyUITest, ref sErrorMessage) == false)
                                    {
                                        Console.WriteLine(sErrorMessage);
                                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                        break;
                                    }
                                }

                                if (irole == 2 && TestCheck == ConstCommon.TEST_PASS)
                                {
                                    if (EpiaUtilities.AddNewAccount(slogFilePath, "AccountShutdownInOneMinute", UserPassword, "DescriptionB", true, "exitModeShutdownRadioButton", 1, string.Empty, 
                                        sOnlyUITest, ref sErrorMessage) == false)
                                    {
                                        Console.WriteLine(sErrorMessage);
                                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                        break;
                                    }
                                }

                                if (irole == 3 && TestCheck == ConstCommon.TEST_PASS)
                                {
                                    if (EpiaUtilities.AddNewAccount(slogFilePath, "RoleLogoutInOneMinute", UserPassword, "DescriptionC", false, null, 0, "RoleLogoutInOneMinute", sOnlyUITest, ref sErrorMessage) == false)
                                    {
                                        Console.WriteLine(sErrorMessage);
                                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                        break;
                                    }
                                }

                                if (irole == 4 && TestCheck == ConstCommon.TEST_PASS)
                                {
                                    if (EpiaUtilities.AddNewAccount(slogFilePath, "RoleShutdownInOneMinute", UserPassword, "DescriptionC", false, null, 0, "RoleShutdownInOneMinute", sOnlyUITest, ref sErrorMessage) == false)
                                    {
                                        Console.WriteLine(sErrorMessage);
                                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                // validate result
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // Find GridView
                    aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                    AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeWindow);
                    if (aeGrid == null)
                    {
                        sErrorMessage = aeWindow.Current.Name + " GridData not found";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine(aeWindow.Current.Name + " GridData found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(3000);


                        string[] accountName = {"", "AccountLogoutInOneMinute", "AccountShutdownInOneMinute", "RoleLogoutInOneMinute", "RoleShutdownInOneMinute" }; 
                        // Construct the Grid Cell Element Name
                        for (int iname = 1; iname < 5; iname++)
                        {
                            string cellname = "Account name" + " Row " + iname;
                            // Get the Element with the Row Col Coordinates
                            AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                            if (aeCell == null)
                            {
                                sErrorMessage = "Find aeCell failed:" + cellname;
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                                break;
                            }
                            else
                            {
                                string cellValue = string.Empty;
                                try
                                {
                                    ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                    cellValue = vp.Current.Value;
                                    Console.WriteLine("Get element.Current Value:" + cellValue);
                                }
                                catch (System.NullReferenceException)
                                {
                                    cellValue = string.Empty;
                                }

                                if (cellValue == null || cellValue == string.Empty)
                                {
                                    sErrorMessage = "aeCell Value not found:" + cellname;
                                    Console.WriteLine(sErrorMessage);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    break;
                                }
                                else if (!cellValue.Equals(accountName[iname]))
                                {
                                    sErrorMessage = "aeCell Value not equal " + accountName[iname] + " , but :" + cellValue;
                                    Console.WriteLine(sErrorMessage);
                                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    break;
                                }
                            }
                        }
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
                result = ConstCommon.TEST_EXCEPTION;
                sErrorMessage = ex.Message;
                Console.WriteLine(testname + " === " + sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                //Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                //	   AutomationElement.RootElement,
                //	  UIALayoutCancelButtonEventHandler);
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
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					   AutomationElement.RootElement,
					  UIALayoutNavigatorOffEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
		#region MultiLanguageCheck
		public static void MultiLanguageCheck(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;


            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 4, 4), ref sErrorMessage) == true)
            {
                result = ConstCommon.TEST_UNDEFINED;
                return;
            }

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}

            System.Windows.Point myPlacePoint = new System.Windows.Point(0, 0);
			try
			{
				string epiaDataResourceFolder = OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand + "\\Epia Server\\Data\\Resources";
				//string resourceFileName = "Epia.Modules.RnD_cn.resources";
				string[] resourceFileNames = { "Epia.Modules.RnD_cn.resources","Epia.Modules.RnD_es.resources",
												 "Epia.Modules.RnD_de.resources","Epia.Modules.RnD_el.resources",
												 "Epia.Modules.RnD_fr.resources","Epia.Modules.RnD_nl.resources",
												 "Epia.Modules.RnD_pl.resources","Epia.Modules.RnD_en.resources"};

                AutomationElement aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                if (aeWindow == null)
                {
                    sErrorMessage = "MainForm not found";
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
                }
                else
                {   // open my setting window
                    aeWindow.SetFocus();
                    string titleBarID = "_MainForm_Toolbars_Dock_Area_Top";
                    AutomationElement aeTitleBar = AUIUtilities.FindElementByID(titleBarID, aeWindow);
                    if (aeTitleBar == null)
                    {
                        sErrorMessage = titleBarID + "not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        double x = aeTitleBar.Current.BoundingRectangle.Left + 100;
                        double y = (aeTitleBar.Current.BoundingRectangle.Top + aeTitleBar.Current.BoundingRectangle.Bottom) / 2;
                        myPlacePoint = new System.Windows.Point(x, y);
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    for (int i = 0; i < resourceFileNames.Length; i++)
                    {
                        if (TestCheck == ConstCommon.TEST_PASS)
                        {
                            if (EpiaUtilities.SwitchLanguageAndFindText(epiaDataResourceFolder, resourceFileNames[i], myPlacePoint, ref sErrorMessage))
                                Epia3Common.WriteTestLogMsg(slogFilePath, resourceFileNames[i] + " OK", sOnlyUITest);
                            else
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                Console.WriteLine(sErrorMessage);
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
                result = ConstCommon.TEST_EXCEPTION;
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
		#region ShellConfigSecurity
		public static int ShellConfigSecurity(string testname, AutomationElement root, out int result)
		{
            Console.WriteLine("\n=== Test " + testname + " === ShellConfigSecurity");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                return result;
			}

            if (sOnlyUITest)
                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 60);

            // test case logon use Generic user or Windows user
			AutomationEventHandler UIAConfigSecurityEventHandler = new AutomationEventHandler(OnConfigSecurityUIAEvent);
            if ( testname.ToLower().IndexOf("inactivitytimeout") >= 0)   // these test case logon with gerneric user
               UIAConfigSecurityEventHandler = new AutomationEventHandler(OnConfigSecurityGenericUserUIAEvent);

			try
			{
                root.SetFocus();
                // Add Open shell configuration window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIAConfigSecurityEventHandler);
                
                if ( ProjBasicUI.ShellAction(root, "configuration", ref sErrorMessage) == false )
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                Thread.Sleep(9000);
                sEventEnd = false;
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);
                while (sEventEnd == false && mTime.TotalSeconds <= 120)
                {
                    Thread.Sleep(2000);
                    Console.WriteLine(" sEventEnd:" + sEventEnd);
                    mTime = DateTime.Now - mStartTime;
                }
                Console.WriteLine("Final sEventEnd:" + sEventEnd);

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					   AutomationElement.RootElement, UIAConfigSecurityEventHandler);

                if (TestCheck == ConstCommon.TEST_PASS)
				{
                    sErrorMessage = string.Empty;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
					result = ConstCommon.TEST_PASS;
				}

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_FAIL;
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
					   AutomationElement.RootElement, UIAConfigSecurityEventHandler);
			}

            return result;
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region LogonCurrentUser
		public static void LogonCurrentUser(string testname, AutomationElement root, out int result)
		{
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

			try
			{
                if (sEpiaServerStartupOK == false)
                {
                    sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                    return;
                }

				if (sOnlyUITest)
                    root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);

                // Shell log off
                if (ProjBasicUI.ShellAction(root, "logoff", ref sErrorMessage) == false)
                {
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
				
                AutomationElement aeSecurityForm = null;
                string tester = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                string myPassword = UserPassword;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    Console.WriteLine("After Logoff, wait until LogonForm displaying... : ");
                    aeSecurityForm = ProjBasicUI.GetMainWindowWithinTime("LogonForm", 120);
                    if (aeSecurityForm != null)
                    {
                        if (tester.ToLower().IndexOf("jiemin") >= 0)
                            myPassword = "tfstest2011";

                        if (ProjBasicUI.Logon(aeSecurityForm, tester, myPassword, ref sErrorMessage ) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        sErrorMessage = "Find Application aeSecurityForm : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
                    if (root == null)
                    {
                        sErrorMessage = "Application MainForm not displayed after logon : " + System.DateTime.Now;
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
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    Thread.Sleep(5000);
                }
			}
			catch (Exception ex)
			{
                result = ConstCommon.TEST_EXCEPTION;
				Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
				Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
			}
		}
		#endregion LogonCurrentUser
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region InactivityTimeoutLogout
        public static void InactivityTimeoutLogout(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 4, 10), ref sErrorMessage) == true)
            {
                result = ConstCommon.TEST_UNDEFINED;
                return;
            }


            try
            {
                if (sEpiaServerStartupOK == false)
                {
                    sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                    return;
                }

                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 60);
                if (root != null)
                {
                    if (ShellConfigSecurity(testname, root, out result) == ConstCommon.TEST_FAIL)
                    {
                        sErrorMessage = "ShellConfigSecurity failed";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                else
                {
                    sErrorMessage = "Find Application aeMainForm not found at beginning of the test: " + System.DateTime.Now;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    if (ProjBasicUI.ShellAction(root, "logoff", ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                AutomationElement aeSecurityForm = null;
                string user = "AccountLogoutInOneMinute";
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    aeSecurityForm = ProjBasicUI.GetMainWindowWithinTime("LogonForm", 120);
                    if (aeSecurityForm != null)
                    {
                        if (testname.Equals(ACCOUNT_INACTIVITY_TIMEOUT_LOGOUT))
                            user = "AccountLogoutInOneMinute";
                        else
                            user = "RoleLogoutInOneMinute";
                        if (ProjBasicUI.Logon(aeSecurityForm, user, UserPassword, ref sErrorMessage) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        sErrorMessage = "Find Application aeSecurityForm failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                Console.WriteLine(" find total mainForm :");
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    AutomationElement aeMainForm = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
                    if (aeMainForm == null)
                    {
                        sErrorMessage = "Find Application aeMainForm after logon failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }


                // wait for logon form
                aeSecurityForm = null;
                AutomationEventHandler UIWaitLogonFormEventHandler = new AutomationEventHandler(OnWaitLogonFormErrorEvent);
          
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                        AutomationElement.RootElement, TreeScope.Descendants, UIWaitLogonFormEventHandler);

                    Thread.Sleep(90000);
                    aeSecurityForm = ProjBasicUI.GetMainWindowWithinTime("LogonForm", 300);

                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement, UIWaitLogonFormEventHandler);

                    if (aeSecurityForm == null)
                    {
                        sErrorMessage = "aeSecurityForm not found after 2 min.";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        if (ProjBasicUI.Logon(aeSecurityForm, TFSQATestTools.Constants.HIDDEN_USERNAME, TFSQATestTools.Constants.HIDDEN_PASSWORD, ref sErrorMessage) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, "--- logon hidden user failed: " + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
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
                    Console.WriteLine(testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            { 
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region InactivityTimeoutLogoutXP
        public static void InactivityTimeoutLogoutXP(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 4, 10), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }


                if (sEpiaServerStartupOK == false)
                {
                    sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                    return;
                }

                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 60);
                if (root != null)
                {
                    if (ShellConfigSecurity(testname, root, out result) == ConstCommon.TEST_FAIL)
                    {
                        sErrorMessage = "ShellConfigSecurity failed";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                else
                {
                    sErrorMessage = "Find Application aeMainForm not found at beginning of the test: " + System.DateTime.Now;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    //ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    if (ProjBasicUI.ShellAction(root, "logoff", ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                AutomationElement aeSecurityForm = null;
                string user = "AccountLogoutInOneMinute";
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    aeSecurityForm = ProjBasicUI.GetMainWindowWithinTime("LogonForm", 120);
                    if (aeSecurityForm != null)
                    {
                        if (testname.Equals(ACCOUNT_INACTIVITY_TIMEOUT_LOGOUT))
                            user = "AccountLogoutInOneMinute";
                        else
                            user = "RoleLogoutInOneMinute";
                        if (ProjBasicUI.Logon(aeSecurityForm, user, UserPassword, ref sErrorMessage) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        sErrorMessage = "Find Application aeSecurityForm failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                Console.WriteLine(" find total mainForm :");
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    AutomationElement aeMainForm = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
                    if (aeMainForm == null)
                    {
                        sErrorMessage = "Find Application aeMainForm after logon failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }


                // wait for logon form
                aeSecurityForm = null;
                AutomationEventHandler UIWaitLogonFormEventHandler = new AutomationEventHandler(OnWaitLogonFormErrorEventXP);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                        AutomationElement.RootElement, TreeScope.Descendants, UIWaitLogonFormEventHandler);

                    Thread.Sleep(90000);
                    aeSecurityForm = ProjBasicUI.GetMainWindowWithinTime("LogonForm", 300);

                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement, UIWaitLogonFormEventHandler);

                    if (aeSecurityForm == null)
                    {
                        sErrorMessage = "aeSecurityForm not found after 2 min.";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        if (ProjBasicUI.Logon(aeSecurityForm, TFSQATestTools.Constants.HIDDEN_USERNAME, TFSQATestTools.Constants.HIDDEN_PASSWORD, ref sErrorMessage) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, "--- logon hidden user failed: " + sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    // try to get error screen message
                    // close shell process
                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    ProcessUtilities.CloseProcess("DW20");
                    Console.WriteLine("Close shell process: ");
                    // start shell
                    ProcessUtilities.StartProcessNoWait(OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand+ @"\Epia Shell",
                        ConstCommon.EGEMIN_EPIA_SHELL_EXE, string.Empty);

                    Console.WriteLine("Searching SplashScreen .... ");
                    AutomationElement aeSplashScreen = ProjBasicUI.GetMainWindowWithinTime("SplashScreen", 120);
                    if (aeSplashScreen != null)
                    {
                        Console.WriteLine("Searching LogonForm .... ");
                        aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", aeSplashScreen);
                        if (aeSecurityForm != null)
                        {
                            if (ProjBasicUI.Logon(aeSecurityForm, TFSQATestTools.Constants.HIDDEN_USERNAME, TFSQATestTools.Constants.HIDDEN_PASSWORD, ref sErrorMessage) == false)
                            {
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                        else
                        {
                            sErrorMessage = "Find Application aeSecurityForm failed : " + System.DateTime.Now;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        sErrorMessage = "Find aeSplashScreen failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }

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
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region InactivityTimeoutShutdown
        public static void InactivityTimeoutShutdown(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 4, 10), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }


                if (sEpiaServerStartupOK == false)
                {
                    sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                    return;
                }

                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 60);
                if (root != null)
                {
                    if (ShellConfigSecurity(testname, root, out result) == ConstCommon.TEST_FAIL)
                    {
                        sErrorMessage = "ShellConfigSecurity failed";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                else
                {
                    sErrorMessage = "Find Application aeMainForm not found at beginning of the test: " + System.DateTime.Now;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    //ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    if (ProjBasicUI.ShellAction(root, "logoff", ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                AutomationElement aeSecurityForm = null;
                string user = "AccountShutdownInOneMinute";
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    aeSecurityForm = ProjBasicUI.GetMainWindowWithinTime("LogonForm", 120);
                    if (aeSecurityForm != null)
                    {
                        if (testname.Equals(ACCOUNT_INACTIVITY_TIMEOUT_SHUTDOWN))
                            user = "AccountShutdownInOneMinute";
                        else
                            user = "RoleShutdownInOneMinute";
                        if (ProjBasicUI.Logon(aeSecurityForm, user, UserPassword, ref sErrorMessage) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        sErrorMessage = "Find Application aeSecurityForm failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                Console.WriteLine(" find total mainForm :");
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    AutomationElement aeMainForm = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
                    if (aeMainForm == null)
                    {
                        sErrorMessage = "Find Application aeMainForm after logon failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                Thread.Sleep(120000);
                AutomationElement aeForm = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeForm = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                    if (aeForm != null)
                    {
                        sErrorMessage = "aeMainForm not shutdown after 2 min.";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("After shutdown wait 15 sec .... ");
                        Thread.Sleep(15000);
                        string path = Path.Combine(m_SystemDrive + ConstCommon.EPIA_CLIENT_ROOT,
                            ConstCommon.EGEMIN_EPIA_SHELL_EXE);
                        System.Diagnostics.Process proc = System.Diagnostics.Process.Start(path);
                        Console.WriteLine("Searching SplashScreen .... " + proc.Id);

                        AutomationElement aeSplashScreen = ProjBasicUI.GetMainWindowWithinTime("SplashScreen", 120);
                        if (aeSplashScreen != null)
                        {
                            Thread.Sleep(3000);
                            aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", aeSplashScreen);
                            int k = 0;
                            while (aeSecurityForm == null && k++ < 20 )
                            {
                                Console.WriteLine("Wait until LogonForm display .... " + k);
                                Epia3Common.WriteTestLogMsg(slogFilePath, "Wait until LogonForm display .... " + k, sOnlyUITest);
                                Thread.Sleep(2000);
                                aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", aeSplashScreen);
                            }

                            if (aeSecurityForm != null)
                            {
                                Thread.Sleep(3000);
                                if (ProjBasicUI.Logon(aeSecurityForm, TFSQATestTools.Constants.HIDDEN_USERNAME, TFSQATestTools.Constants.HIDDEN_PASSWORD, ref sErrorMessage) == false)
                                {
                                    Console.WriteLine(sErrorMessage);
                                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                            }
                            else
                            {
                                sErrorMessage = "Find Application aeSecurityForm failed : " + System.DateTime.Now;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                        else
                        {
                            sErrorMessage = "Find aeSplashScreen failed : " + System.DateTime.Now;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
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
                    Console.WriteLine(testname + ": Pass");
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
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
        #region InactivityTimeoutShutdownXP
        public static void InactivityTimeoutShutdownXP(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            try
            {
                ProcessUtilities.CloseProcess("DW20");

                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 4, 10), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }


                if (sEpiaServerStartupOK == false)
                {
                    sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
                    return;
                }

                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 60);
                if (root != null)
                {
                    if (ShellConfigSecurity(testname, root, out result) == ConstCommon.TEST_FAIL)
                    {
                        sErrorMessage = "ShellConfigSecurity failed";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                else
                {
                    sErrorMessage = "Find Application aeMainForm not found at beginning of the test: " + System.DateTime.Now;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                    //ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    if (ProjBasicUI.ShellAction(root, "logoff", ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                AutomationElement aeSecurityForm = null;
                string user = "AccountShutdownInOneMinute";
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    aeSecurityForm = ProjBasicUI.GetMainWindowWithinTime("LogonForm", 120);
                    if (aeSecurityForm != null)
                    {
                        if (testname.Equals(ACCOUNT_INACTIVITY_TIMEOUT_SHUTDOWN))
                            user = "AccountShutdownInOneMinute";
                        else
                            user = "RoleShutdownInOneMinute";
                        if (ProjBasicUI.Logon(aeSecurityForm, user, UserPassword, ref sErrorMessage) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        sErrorMessage = "Find Application aeSecurityForm failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                Console.WriteLine(" find total mainForm :");
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Thread.Sleep(3000);
                    AutomationElement aeMainForm = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
                    if (aeMainForm == null)
                    {
                        sErrorMessage = "Find Application aeMainForm after logon failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                Thread.Sleep(120000);
                AutomationElement aeForm = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeForm = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
                    if (aeForm != null)
                    {
                        sErrorMessage = "aeMainForm not shutdown after 2 min.";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Console.WriteLine("After shutdown wait 15 sec .... ");
                        Thread.Sleep(15000);
                        string path = Path.Combine(m_SystemDrive + ConstCommon.EPIA_CLIENT_ROOT,
                            ConstCommon.EGEMIN_EPIA_SHELL_EXE);
                        System.Diagnostics.Process proc = System.Diagnostics.Process.Start(path);
                        Console.WriteLine("Searching SplashScreen .... " + proc.Id);

                        AutomationElement aeSplashScreen = ProjBasicUI.GetMainWindowWithinTime("SplashScreen", 120);
                        if (aeSplashScreen != null)
                        {
                            Thread.Sleep(3000);
                            aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", aeSplashScreen);
                            int k = 0;
                            while (aeSecurityForm == null && k++ < 10)
                            {
                                Thread.Sleep(2000);
                                aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", aeSplashScreen);
                            }

                            if (aeSecurityForm != null)
                            {
                                Thread.Sleep(3000);
                                if (ProjBasicUI.Logon(aeSecurityForm, TFSQATestTools.Constants.HIDDEN_USERNAME, TFSQATestTools.Constants.HIDDEN_PASSWORD, ref sErrorMessage) == false)
                                {
                                    Console.WriteLine(sErrorMessage);
                                    Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                }
                            }
                            else
                            {
                                sErrorMessage = "Find Application aeSecurityForm failed : " + System.DateTime.Now;
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                        else
                        {
                            sErrorMessage = "Find aeSplashScreen failed : " + System.DateTime.Now;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    // try to get error screen message
                    // close shell process
                    ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    ProcessUtilities.CloseProcess("DW20");
                    Console.WriteLine("Close shell process: ");
                    // start shell
                    ProcessUtilities.StartProcessNoWait(OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand+ @"\Epia Shell",
                        ConstCommon.EGEMIN_EPIA_SHELL_EXE, string.Empty);

                    Console.WriteLine("Searching SplashScreen .... ");
                    AutomationElement aeSplashScreen = ProjBasicUI.GetMainWindowWithinTime("SplashScreen", 120);
                    if (aeSplashScreen != null)
                    {
                        Console.WriteLine("Searching LogonForm .... ");
                        aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", aeSplashScreen);
                        if (aeSecurityForm != null)
                        {
                            if (ProjBasicUI.Logon(aeSecurityForm, TFSQATestTools.Constants.HIDDEN_USERNAME, TFSQATestTools.Constants.HIDDEN_PASSWORD, ref sErrorMessage) == false)
                            {
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                        else
                        {
                            sErrorMessage = "Find Application aeSecurityForm failed : " + System.DateTime.Now;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        sErrorMessage = "Find aeSplashScreen failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }

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
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine(testname + " === " + ex.Message + "------" + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
        }
        #endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region LogonEpiaAdministrator
		public static void LogonEpiaAdministrator(string testname, AutomationElement root, out int result)
		{
			ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SHELL );
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			AutomationEventHandler UIAConfigSecurityEventHandler = new AutomationEventHandler(OnLogonEpiaAdminUIAEvent);

			Thread.Sleep(7000);
			try
			{
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (ProjBasicUI.ShellAction(root, "configuration", ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					   AutomationElement.RootElement,
					  UIAConfigSecurityEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

				Thread.Sleep(5000);
				if (TestCheck == ConstCommon.TEST_FAIL)
				{
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
					result = ConstCommon.TEST_FAIL;
					return;
				}

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (ProjBasicUI.ShellAction(root, "logoff", ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

				// logon with hidden user
				Console.WriteLine("Application is started : ");

				DateTime mStartTime2 = DateTime.Now;
				TimeSpan mTime2 = DateTime.Now - mStartTime2;
				AutomationElement aeSecurityForm = null;
                while (aeSecurityForm == null && mTime2.TotalSeconds < 120)
				{
					aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", AutomationElement.RootElement);
					Thread.Sleep(2000);
                    Console.WriteLine(" time is :" + mTime2.TotalSeconds);
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
                string tester = TFSQATestTools.Constants.HIDDEN_USERNAME;

				if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester, ref sErrorMessage))
					Thread.Sleep(3000);
				else
				{
					sErrorMessage = "FindTextBoxAndChangeValue failed:" + UserNameID;
					Console.WriteLine(sErrorMessage);
					result = ConstCommon.TEST_FAIL;
					return;
				}

                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser, TFSQATestTools.Constants.HIDDEN_USERNAME, ref sErrorMessage))
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
            TestCheck = ConstCommon.TEST_PASS;

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}

			try
			{
                root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
                if (ShellConfigSecurity(testname, root, out result) == ConstCommon.TEST_FAIL)
                {
                    sErrorMessage = "ShellConfigSecurity failed";
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_FAIL;
                }

                Console.WriteLine("--- ShellConfigSecurity OK");
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (ProjBasicUI.ShellAction(root, "shutdown", ref sErrorMessage) == false)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
				
				Epia3Common.WriteTestLogMsg(slogFilePath, "Epia shutdown:", sOnlyUITest);
                DateTime mAppTime = DateTime.Now;
                TimeSpan Time = DateTime.Now - mAppTime;
                bool shutdownOK = false;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    while (shutdownOK == false && Time.TotalSeconds < 120)
                    {
                        root = ProjBasicUI.GetMainWindowWithinTime("MainForm", 5);
                        if (root == null)
                            shutdownOK = true;
                        else
                        {
                            Console.WriteLine(" find time is :" + Time.TotalSeconds);
                            Thread.Sleep(5000);
                            Time = DateTime.Now - mAppTime;
                        } 
                    }

                    if (shutdownOK == false)
                    {
                        sErrorMessage = "aeForm not shutdown after 2 min.";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
				
				if (result == ConstCommon.TEST_PASS)
				{
                    sErrorMessage = string.Empty;
					Console.WriteLine("\nTest Shell Shutdown: Pass");
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
					result = ConstCommon.TEST_PASS;
				}
				else
				{
					Console.WriteLine("\nTest Test Shell Shutdown: *FAIL*");
					result = ConstCommon.TEST_FAIL;
					Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
				}
			}
			catch (Exception ex)
			{
                result = ConstCommon.TEST_EXCEPTION;
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
            TestCheck = ConstCommon.TEST_PASS;

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}

			Thread.Sleep(5000);
			try
			{
				string path = Path.Combine(m_SystemDrive + ConstCommon.EPIA_CLIENT_ROOT,
					ConstCommon.EGEMIN_EPIA_SHELL_EXE);
				System.Diagnostics.Process proc = System.Diagnostics.Process.Start(path);
				
                string tester = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                string myPassword = UserPassword;
                AutomationElement aeSecurityForm = null;
                if(TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Searching SplashScreen .... ");
                    AutomationElement aeSplashScreen = ProjBasicUI.GetMainWindowWithinTime("SplashScreen", 120);
                    if (aeSplashScreen != null)
                    {
                        Console.WriteLine("Searching LogonForm .... ");
                        aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", aeSplashScreen);
                        if (aeSecurityForm != null)
                        {
                            if (tester.ToLower().IndexOf("jiemin") >= 0)
                                myPassword = "tfstest2011";

                            if (ProjBasicUI.Logon(aeSecurityForm, tester, myPassword, ref sErrorMessage) == false)
                            {
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                        }
                        else
                        {
                            sErrorMessage = "Find Application aeSecurityForm failed : " + System.DateTime.Now;
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        sErrorMessage = "Find aeSplashScreen failed : " + System.DateTime.Now;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                Thread.Sleep(3000);
                AutomationElement aeMainForm = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeMainForm = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
                    if (aeMainForm == null)
                    {
                        sErrorMessage = "aeForm  not found : ";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        if (ProjBasicUI.ShellAction(aeMainForm, "logoff", ref sErrorMessage) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                aeSecurityForm = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    myPassword = UserPassword;
                    aeSecurityForm = ProjBasicUI.GetMainWindowWithinTime("LogonForm", 120); 
                    if (aeSecurityForm != null)
                    {
                        if (tester.ToLower().IndexOf("jiemin") >= 0)
                            myPassword = "tfstest2011";

                        if (ProjBasicUI.Logon(aeSecurityForm, tester, myPassword, ref sErrorMessage) == false)
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        sErrorMessage = "Find Application aeSecurityForm : " + System.DateTime.Now;
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
                result = ConstCommon.TEST_EXCEPTION;
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
		#region Epia4Close
		public static void Epia4Close(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			string BtnCloseID = "Close";

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}

			try
			{
				Thread.Sleep(5000);
                AutomationElement aeMainForm = null;
				// Close the other mainForms
				System.Diagnostics.Process[] pShell = System.Diagnostics.Process.GetProcessesByName(ConstCommon.EGEMIN_EPIA_SHELL);
				for (int i = 0; i < pShell.Length; i++)
				{
                    try
                    {
                        aeMainForm = AutomationElement.FromHandle(pShell[i].MainWindowHandle);
                        aeMainForm.SetFocus();
                        Thread.Sleep(5000);
                        /*while (!ProjBasicUI.GetThisWindowIsTopMost(aeMainForm))
                        {
                            Console.WriteLine(testname + " aeMainForm is not TopMost ");
                            Thread.Sleep(5000);
                            Epia3Common.WriteTestLogMsg(slogFilePath, testname + " aeMainForm is not TopMost ", sOnlyUITest);
                        }*/
	
                      
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
                    catch (Exception ex)
                    {
                        Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
                    }
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
				int pID = ProcessUtilities.GetApplicationProcessID( ConstCommon.EGEMIN_EPIA_SHELL, out proc );
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
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ShellCloseWithinOneMinuteAfterServerDown
        public static void ShellCloseWithinOneMinuteAfterServerDown(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			//string BtnCloseID = "Close";

			if (sEpiaServerStartupOK == false)
			{
				sErrorMessage = "Epia Server startup failed, this testcase cannot be tested";
				return;
			}

			try
			{
                StartEpiaServerShell(START_EPIA_SERVER_SHELL, root, out result);
                if (result == ConstCommon.TEST_PASS)
                {
                    // stop Server 
                    Console.WriteLine("Stop EPIA SERVER as Service : ");
                    ProcessUtilities.StartProcessWaitForExit(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() + "\\" + sBrand+
						@"\Epia Server",
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
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)//,
                    //new PropertyCondition(AutomationElement.AutomationIdProperty, "Dialog - Egemin.Epia.Presentation.WinForms.LicenseRegistrationScreen")
                    );

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sStartTime = DateTime.Now;
                    sTime = DateTime.Now - sStartTime;
                    Console.WriteLine(" time is :" + sTime.TotalSeconds);
                    while (aeLicenseServiceShutdownDialogBox == null && sTime.TotalSeconds < 120)
                    {
                        aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                        if (aeWindow == null)
                        {
                            sErrorMessage = "MainForm is not found";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                            break;
                        }
                        else
                        {
                            Console.WriteLine(" wait until dialog box :" + sTime.TotalSeconds);
                            aeLicenseServiceShutdownDialogBox = aeWindow.FindFirst(TreeScope.Descendants, cWindowLicenseShutdown);
                            Thread.Sleep(2000);
                            sTime = DateTime.Now - sStartTime;
                            Console.WriteLine("wait aeLicenseServiceShutdownDialogBox displayed time is (sec) : " + sTime.TotalSeconds);
                        }
                    }

                    if (aeLicenseServiceShutdownDialogBox == null)
                    {
                        sErrorMessage = "aeLicenseServiceShutdownDialogBox not displayed after 2 min";
                        Console.WriteLine(testname + sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {

                        System.Windows.Automation.Condition c = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, "Shell shutdown"),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                        );

                        // Find the BUTTON element.
                        aeShellShutdownButton = aeLicenseServiceShutdownDialogBox.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                        if (aeShellShutdownButton != null)
                        {
                            Point pt = AUIUtilities.GetElementCenterPoint(aeShellShutdownButton);
                            Input.MoveTo(pt);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(pt);
                        }
                    }
                }

                // After one minute shell should be closed
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    sStartTime = DateTime.Now;
                    sTime = DateTime.Now - sStartTime;
                    Console.WriteLine(" time is :" + sTime.TotalSeconds);
                    aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 5);
                    while (aeWindow != null && sTime.TotalSeconds < 61)
                    {
                        aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 5);
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

                    //after one minute
                    if (aeWindow != null)
                    {
                        sErrorMessage = "Shell is still open after one minute";
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

                Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_EXCEPTION;
				Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
                ProcessUtilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
			}
		}
        #endregion ShellCloseWithinOneMinuteAfterServerDown
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Epia4CleanUninstallCheck
		public static void Epia4CleanUninstallCheck(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;
			
			string EpiaServerFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\" + sBrand+ "\\Epia Server";
			string EpiaShellFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\" + sBrand+ "\\Epia Shell";
			Thread.Sleep(5000);
			try
			{
				ProcessUtilities.CloseProcess( "Egemin.Epia.Shell" );
				ProcessUtilities.CloseProcess( "Egemin.Epia.Server" );
                Console.WriteLine("--------------- "+DeployUtilities.getThisPCOS());
                // uninstall playback if already installed:
                if (DeployUtilities.getThisPCOS().StartsWith("WindowsXP.32"))
                {
                    Console.WriteLine("---------------  UninstallApplicationXP");
                    //Thread.Sleep(20000);
                    if (ProjAppInstall.UninstallApplicationXP(EgeminApplication.EPIA, ref sErrorMessage) == false)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = EgeminApplication.EPIA + " Uninstall failed:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                    }
                }
                else
                {
                    if (ProjAppInstall.UninstallApplication(EgeminApplication.EPIA, ref sErrorMessage) == false)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = EgeminApplication.EPIA + " Uninstall failed:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                    }
                }

				if (TestCheck == ConstCommon.TEST_PASS)
				{
					if (System.IO.Directory.Exists(EpiaServerFolder)
						|| System.IO.Directory.Exists(EpiaShellFolder))
					{
						if (System.IO.Directory.Exists(EpiaServerFolder))
						{
							// get files in ServerFolder
							DirectoryInfo DirInfo = new DirectoryInfo(EpiaServerFolder);
							FileInfo[] serverFolderFiles = DirInfo.GetFiles("*.*");
							if (serverFolderFiles.Length > 1)
							{
								TestCheck = ConstCommon.TEST_FAIL;
								sErrorMessage = EpiaServerFolder + " still has some files:" + serverFolderFiles[0].FullName;
								Console.WriteLine(sErrorMessage);
								//System.Windows.Forms.MessageBox.Show(sErrorMessage);
							}
						}

                        if (TestCheck == ConstCommon.TEST_PASS)
                        {
                            if (System.IO.Directory.Exists(EpiaShellFolder))
                            {
                                // get files in ShellFolder
                                DirectoryInfo DirInfo = new DirectoryInfo(EpiaShellFolder);
                                FileInfo[] shellFolderFiles = DirInfo.GetFiles("*.*");
                                if (shellFolderFiles.Length > 0)
                                {
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    sErrorMessage = EpiaShellFolder + " still has some files:" + shellFolderFiles[0].FullName;
                                    Console.WriteLine(sErrorMessage);
                                    //System.Windows.Forms.MessageBox.Show(sErrorMessage);
                                }
                            }
                        }
					}
					else
					{
						TestCheck = ConstCommon.TEST_PASS;
					}
				}

				if (testname.ToLower().StartsWith("no"))
				{
					Console.WriteLine("do nothing, not test case:" + testname);
				}
				else
				{
					if (TestCheck == ConstCommon.TEST_FAIL)
					{
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
					}

					if (TestCheck == ConstCommon.TEST_PASS)
					{
						result = ConstCommon.TEST_PASS;
						sErrorMessage = string.Empty;
						Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
					}
				}

				Thread.Sleep(3000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_EXCEPTION;
				Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
			}
			finally
			{
                if (System.IO.Directory.Exists(EpiaServerFolder))
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(EpiaServerFolder);
                    while (System.IO.Directory.Exists(EpiaServerFolder))
                    {
                        FileManipulation.DeleteRecursiveFolder(dirInfo);
                        Thread.Sleep(2000);
                    }
                }

                if (System.IO.Directory.Exists(EpiaShellFolder))
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(EpiaShellFolder);
                    while (System.IO.Directory.Exists(EpiaShellFolder))
                    {
                        FileManipulation.DeleteRecursiveFolder(dirInfo);
                        Thread.Sleep(2000);
                    }
                }
			}
		}
		#endregion Epia4CleanUninstallCheckXP

		static private void Wait(int seconds)
		{
			System.Threading.Thread.Sleep(seconds * 1000);
		}

		static private void StartEpiaApplicationExecution()
		{
			Process Proc = new System.Diagnostics.Process();
			string installDir = OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand+ @"\AutomaticTesting\Setup\Epia\Current";
			Proc.StartInfo.FileName = Path.Combine(installDir, "Epia.msi");
			Proc.StartInfo.CreateNoWindow = false;
			Proc.Start();
		}

		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Epia4CleanShellInstallCheck
		public static void Epia4CleanShellInstallCheck(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
			try
			{
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 3, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    string path = OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand + @"\AutomaticTesting\Setup\Epia\Current";

                    if (DeployUtilities.getThisPCOS().StartsWith("Windows8.64") || DeployUtilities.getThisPCOS().StartsWith("WindowsServer2012.64") )
                    {
                        if (ProjAppInstall.InstallApplicationNet45(path, EgeminApplication.EPIA, EgeminApplication.SetupType.EpiaShellOnly, ref sErrorMessage, null/*logger*/))
                        {
                            TestCheck = ConstCommon.TEST_PASS;
                        }
                        else
                        {
                            sErrorMessage = EgeminApplication.EPIA + " Uninstall failed:" + sErrorMessage;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        if (ProjAppInstall.InstallApplication(path, EgeminApplication.EPIA, EgeminApplication.SetupType.EpiaShellOnly, ref sErrorMessage, null/*logger*/))
                        {
                            TestCheck = ConstCommon.TEST_PASS;
                        }
                        else
                        {
                            sErrorMessage = EgeminApplication.EPIA + " Uninstall failed:" + sErrorMessage;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

               
				// Add Open window Event Handler
				/*Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);

				Console.WriteLine("Start the MSI");
				Thread executableThread = new Thread(new ThreadStart(StartEpiaApplicationExecution));
				executableThread.Start();
				Wait(15);

				Console.WriteLine("Start Epia Installation -------------->");
				bool status = EpiaUtilities.InstallEpia("Shell", ref sErrorMessage);

				if (status == false)
				{
					TestCheck = ConstCommon.TEST_FAIL;
					Console.WriteLine(sErrorMessage);
					System.Windows.Forms.MessageBox.Show(sErrorMessage);
				}*/

				// validate shell installation
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					if (System.IO.Directory.Exists(sEpiaServerFolder))  // for install Shell folder, server folder should not exist
					{

						sErrorMessage = "error installed, Server folder exist,";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
						if (sOnlyUITest)
							System.Windows.Forms.MessageBox.Show(sErrorMessage);
					}
					else if (System.IO.Directory.Exists(sEpiaShellFolder))
					{
						// get files in ShellFolder
						DirectoryInfo DirInfo = new DirectoryInfo(sEpiaShellFolder);
						FileInfo[] shellFolderFiles = DirInfo.GetFiles("*.*");
						if (shellFolderFiles.Length == 0)
						{
							TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = sEpiaShellFolder + " has no installed files:"; // +shellFolderFiles[0].FullName;
							Console.WriteLine(sErrorMessage);
							if (sOnlyUITest)
								System.Windows.Forms.MessageBox.Show(sErrorMessage);
						}
					}
					else
					{
						TestCheck = ConstCommon.TEST_FAIL;
						sErrorMessage = "error installed, Shell folder not exist,";
						Console.WriteLine(sErrorMessage);
						if (sOnlyUITest)
							System.Windows.Forms.MessageBox.Show(sErrorMessage);
					}
				}

				if (TestCheck == ConstCommon.TEST_FAIL)
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
				}

				if (TestCheck == ConstCommon.TEST_PASS)
				{
					result = ConstCommon.TEST_PASS;
					sErrorMessage = string.Empty;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
				}
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_EXCEPTION;
				sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine(sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
               // System.Windows.Forms.MessageBox.Show(sErrorMessage);
			}
			finally
			{
				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					 AutomationElement.RootElement, UIAShellEventHandler);
			}
		}
		#endregion Epia4CleanShellInstallCheck
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Epia4CleanServerInstallCheck
		public static void Epia4CleanServerInstallCheck(string testname, AutomationElement root, out int result)
		{
			
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
			try
			{
                Epia4CleanUninstallCheck("NoTestcase", AutomationElement.RootElement, out result);
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 3, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    string path = OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand+ @"\AutomaticTesting\Setup\Epia\Current";
                    if (DeployUtilities.getThisPCOS().StartsWith("Windows8.64") || DeployUtilities.getThisPCOS().StartsWith("WindowsServer2012.64"))
                    {
                        if (ProjAppInstall.InstallApplicationNet45(path, EgeminApplication.EPIA, EgeminApplication.SetupType.EpiaServerOnly, ref sErrorMessage, null/*logger*/))
                        {
                            TestCheck = ConstCommon.TEST_PASS;
                        }
                        else
                        {
                            sErrorMessage = EgeminApplication.EPIA + " Uninstall failed:" + sErrorMessage;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                    else
                    {
                        if (ProjAppInstall.InstallApplication(path, EgeminApplication.EPIA, EgeminApplication.SetupType.EpiaServerOnly, ref sErrorMessage, null/*logger*/))
                        {
                            TestCheck = ConstCommon.TEST_PASS;
                        }
                        else
                        {
                            sErrorMessage = EgeminApplication.EPIA + " Uninstall failed:" + sErrorMessage;
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

				// validate shell installation
                string EpiaShellFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\" + sBrand+ "\\Epia Shell";
                if (!System.IO.Directory.Exists(EpiaShellFolder))
                {
                    Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + "EpiaShellFolder NOT EXIST OK: " + EpiaShellFolder, sOnlyUITest);
                    TestCheck = ConstCommon.TEST_PASS;
                }
                else if (TestCheck == ConstCommon.TEST_PASS)
				{
                    if (sMicrosoftVisualStudioShellDesignDllInstalled == true)   // see testcase 2
                    {
                        Console.WriteLine("sMicrosoftVisualStudioShellDesignDllInstalled == true");
                        // get files in ShellFolder
                        DirectoryInfo DirInfo = new DirectoryInfo(EpiaShellFolder);
                        FileInfo[] EpiaShellFolderFiles = DirInfo.GetFiles("*.*");
                        int extraCount = 0;
                        string ignoredDll = string.Empty;

                        for (int i = 0; i < EpiaShellFolderFiles.Length; i++)
                        {
                            if (EpiaShellFolderFiles[i].FullName.IndexOf("VisualStudio.Shell.Design") > 0)
                            {
                                ignoredDll = ignoredDll + "Microsoft.VisualStudio.Shell.Design.dll" + ", ";
                                Console.WriteLine("ignoredDll:" + ignoredDll);
                                extraCount++;
                            }
                            else if (EpiaShellFolderFiles[i].FullName.ToLower().IndexOf("stdole") > 0)
                            {
                                ignoredDll = ignoredDll + "stdole.dll" + ", ";
                                Console.WriteLine("ignoredDll:" + ignoredDll);
                                extraCount++;
                            }
                            else if (EpiaShellFolderFiles[i].FullName.IndexOf("DocumentFormat.OpenXml") > 0)
                            {
                                ignoredDll = ignoredDll + "DocumentFormat.OpenXml.dll" + ", ";
                                Console.WriteLine("ignoredDll:" + ignoredDll);
                                extraCount++;
                            }
                            else if (EpiaShellFolderFiles[i].FullName.IndexOf("System.Reactive.Core") > 0)
                            {
                                ignoredDll = ignoredDll + "System.Reactive.Core.dll" + ", ";
                                Console.WriteLine("ignoredDll:" + ignoredDll);
                                extraCount++;
                            }
                            else if (EpiaShellFolderFiles[i].FullName.IndexOf("System.Reactive.Interfaces") > 0)
                            {
                                ignoredDll = ignoredDll + "System.Reactive.Interfaces.dll" + ", ";
                                Console.WriteLine("ignoredDll:" + ignoredDll);
                                extraCount++;
                            }
                            else if (EpiaShellFolderFiles[i].FullName.IndexOf("System.Reactive.Linq") > 0)
                            {
                                ignoredDll = ignoredDll + "System.Reactive.Linq.dll" + ", ";
                                Console.WriteLine("ignoredDll:" + ignoredDll);
                                extraCount++;
                            }
                            else if (EpiaShellFolderFiles[i].FullName.IndexOf("System.Reactive.Windows.Threading") > 0)
                            {
                                ignoredDll = ignoredDll + "System.Reactive.Windows.Threading.dll" + ", ";
                                Console.WriteLine("ignoredDll:" + ignoredDll);
                                extraCount++;
                            }
                            else if (EpiaShellFolderFiles[i].FullName.IndexOf("Egemin.Epia.Foundation.SqlRptServices.Interfaces") > 0)
                            {
                                ignoredDll = ignoredDll + "Egemin.Epia.Foundation.SqlRptServices.Interfaces.dll" + ", ";
                                Console.WriteLine("ignoredDll:" + ignoredDll);
                                extraCount++;
                            }
                        }

                        Console.WriteLine("EpiaShellFolderFiles.Length:" + EpiaShellFolderFiles.Length);
                        Console.WriteLine("extraCount:" + extraCount);
                        if (EpiaShellFolderFiles.Length > extraCount)
                        {
                            string extraDlls = string.Empty;
                            for (int i = 0; i < EpiaShellFolderFiles.Length; i++)
                            {
                                if (EpiaShellFolderFiles[i].FullName.IndexOf("VisualStudioShellDesign") < 0)
                                    //&& EpiaShellFolderFiles[i].FullName.IndexOf("stdole") < 0)
                                    extraDlls = extraDlls + EpiaShellFolderFiles[i].Name + ";";
                            }
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "Shell folder has more extra file than " +ignoredDll + "--> :" + extraDlls;
                            Console.WriteLine(sErrorMessage);
                        }
                        else
                        {
                            sErrorMessage = EpiaShellFolder + " has extra dlls:" + ignoredDll;
                        }
                    }
                    else
                    {
                        Console.WriteLine("sMicrosoftVisualStudioShellDesignDllInstalled == false");
                        Thread.Sleep(10000);
                        if (System.IO.Directory.Exists(EpiaShellFolder))
                        {
                            // get files in ShellFolder
                            /*DirectoryInfo DirInfo = new DirectoryInfo(EpiaShellFolder);
                            FileInfo[] EpiaShellFolderFiles = DirInfo.GetFiles("*.*");
                            string extraDlls = string.Empty;
                            if (EpiaShellFolderFiles.Length > 1)
                            {
                                for (int i = 0; i < EpiaShellFolderFiles.Length; i++)
                                {
                                    if (EpiaShellFolderFiles[i].FullName.IndexOf("VisualStudioShellDesign") < 0
                                        && EpiaShellFolderFiles[i].FullName.IndexOf("stdole") < 0)
                                        extraDlls = extraDlls + EpiaShellFolderFiles[i].Name + ";";
                                }
                            }

                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "error installed, Shell folder exist: extra dlls" + extraDlls;*/
                            // get files in ShellFolder
                            DirectoryInfo DirInfo = new DirectoryInfo(EpiaShellFolder);
                            //System.Windows.Forms.MessageBox.Show("test1");
                            FileInfo[] EpiaShellFolderFiles = DirInfo.GetFiles("*.*");
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + "test2", sOnlyUITest);
                            //System.Windows.Forms.MessageBox.Show("test2");
                            int extraCount = 0;
                            string ignoredDll = string.Empty;
                            Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + "EpiaShellFolderFiles.Length:" + EpiaShellFolderFiles.Length, sOnlyUITest);
                            System.Windows.Forms.MessageBox.Show("EpiaShellFolderFiles.Length:" + EpiaShellFolderFiles.Length);
                            for (int i = 0; i < EpiaShellFolderFiles.Length; i++)
                            {
                                if (EpiaShellFolderFiles[i].FullName.IndexOf("VisualStudio.Shell.Design") > 0)
                                {
                                    ignoredDll = ignoredDll + "Microsoft.VisualStudio.Shell.Design.dll" + ", ";
                                    Console.WriteLine("ignoredDll:" + ignoredDll);
                                    extraCount++;
                                }
                                else if (EpiaShellFolderFiles[i].FullName.ToLower().IndexOf("stdole") > 0)
                                {
                                    ignoredDll = ignoredDll + "stdole.dll" + ", ";
                                    Console.WriteLine("ignoredDll:" + ignoredDll);
                                    extraCount++;
                                }
                                else if (EpiaShellFolderFiles[i].FullName.IndexOf("DocumentFormat.OpenXml") > 0)
                                {
                                    ignoredDll = ignoredDll + "DocumentFormat.OpenXml.dll" + ", ";
                                    Console.WriteLine("ignoredDll:" + ignoredDll);
                                    extraCount++;
                                }
                                else if (EpiaShellFolderFiles[i].FullName.IndexOf("System.Reactive.Core") > 0)
                                {
                                    ignoredDll = ignoredDll + "System.Reactive.Core.dll" + ", ";
                                    Console.WriteLine("ignoredDll:" + ignoredDll);
                                    extraCount++;
                                }
                                else if (EpiaShellFolderFiles[i].FullName.IndexOf("System.Reactive.Interfaces") > 0)
                                {
                                    ignoredDll = ignoredDll + "System.Reactive.Interfaces.dll" + ", ";
                                    Console.WriteLine("ignoredDll:" + ignoredDll);
                                    extraCount++;
                                }
                                else if (EpiaShellFolderFiles[i].FullName.IndexOf("System.Reactive.Linq") > 0)
                                {
                                    ignoredDll = ignoredDll + "System.Reactive.Linq.dll" + ", ";
                                    Console.WriteLine("ignoredDll:" + ignoredDll);
                                    extraCount++;
                                }
                                else if (EpiaShellFolderFiles[i].FullName.IndexOf("System.Reactive.Windows.Threading") > 0)
                                {
                                    ignoredDll = ignoredDll + "System.Reactive.Windows.Threading.dll" + ", ";
                                    Console.WriteLine("ignoredDll:" + ignoredDll);
                                    extraCount++;
                                }
                                else if (EpiaShellFolderFiles[i].FullName.IndexOf("Egemin.Epia.Foundation.SqlRptServices.Interfaces") > 0)
                                {
                                    ignoredDll = ignoredDll + "Egemin.Epia.Foundation.SqlRptServices.Interfaces.dll" + ", ";
                                    Console.WriteLine("ignoredDll:" + ignoredDll);
                                    extraCount++;
                                }
                            }

                            Console.WriteLine("EpiaShellFolderFiles.Length:" + EpiaShellFolderFiles.Length);
                            Console.WriteLine("extraCount:" + extraCount);
                            if (EpiaShellFolderFiles.Length > extraCount)
                            {
                                string extraDlls = string.Empty;
                                for (int i = 0; i < EpiaShellFolderFiles.Length; i++)
                                {
                                    if (EpiaShellFolderFiles[i].FullName.IndexOf("VisualStudioShellDesign") < 0)
                                        //&& EpiaShellFolderFiles[i].FullName.IndexOf("stdole") < 0)
                                        extraDlls = extraDlls + EpiaShellFolderFiles[i].Name + ";";
                                }
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = "Shell folder has more extra file than " + ignoredDll + "--> :" + extraDlls;
                                Console.WriteLine(sErrorMessage);
                            }
                            else
                            {
                                sErrorMessage = EpiaShellFolder + " has extra dlls:" + ignoredDll;
                            }

                            Console.WriteLine(sErrorMessage);
                            if (sOnlyUITest)
                                System.Windows.Forms.MessageBox.Show(sErrorMessage);
                        }
                        else if (System.IO.Directory.Exists(sEpiaServerFolder))
                        {
                            // get files in ServerFolder
                            DirectoryInfo DirInfo = new DirectoryInfo(sEpiaServerFolder);
                            FileInfo[] serverFolderFiles = DirInfo.GetFiles("*.*");
                            if (serverFolderFiles.Length == 0)
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = sEpiaServerFolder + " has no installed files:"; // +serverFolderFiles[0].FullName;
                                Console.WriteLine(sErrorMessage);
                            }
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "error installed, Server folder not exist,";
                            Console.WriteLine(sErrorMessage);
                        }
                    }
                }

				if (TestCheck == ConstCommon.TEST_FAIL)
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
				}


				if (TestCheck == ConstCommon.TEST_PASS)
				{
					result = ConstCommon.TEST_PASS;
					//sErrorMessage = string.Empty;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
				}
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_EXCEPTION;
				sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine(sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
			}
			finally
			{
				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						AutomationElement.RootElement, UIAShellEventHandler);
			}
		}
		#endregion Epia4CleanServerInstallCheck
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region UninstallEpiaResourceFileEditorIfAlreadyInstalled
		public static void UninstallEpiaResourceFileEditorIfAlreadyInstalled(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			string EpiaResourceFileEditorFolder = OSVersionInfoClass.ProgramFilesx86() + "\\Epia Resource File Editor";
			Thread.Sleep(5000);
			try
			{
				ProcessUtilities.CloseProcess( "Egemin.Epia.Shell" );
				ProcessUtilities.CloseProcess( "Egemin.Epia.Server" );

                // uninstall epia resource file editor if already installed:
                if (DeployUtilities.getThisPCOS().StartsWith("WindowsXP.32"))
                {
                    if (ProjAppInstall.UninstallApplicationXP(EgeminApplication.EPIA_RESOURCEFILEEDITOR, ref sErrorMessage) == false)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = EgeminApplication.EPIA_RESOURCEFILEEDITOR + " Uninstall failed:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                    }
                }
                else
                {
                    if (ProjAppInstall.UninstallApplication(EgeminApplication.EPIA_RESOURCEFILEEDITOR, ref sErrorMessage) == false)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = EgeminApplication.EPIA_RESOURCEFILEEDITOR + " Uninstall failed:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                    }
                }

				if (TestCheck == ConstCommon.TEST_PASS)
				{
					if (System.IO.Directory.Exists(EpiaResourceFileEditorFolder))
					{
						TestCheck = ConstCommon.TEST_FAIL;
						sErrorMessage = "Directory still exist:" + EpiaResourceFileEditorFolder;
						Console.WriteLine(sErrorMessage);
						System.Windows.Forms.MessageBox.Show(sErrorMessage);
					}
					else
					{
						TestCheck = ConstCommon.TEST_PASS;
					}
				}

				if (TestCheck == ConstCommon.TEST_FAIL)
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
				}

				if (TestCheck == ConstCommon.TEST_PASS)
				{
					result = ConstCommon.TEST_PASS;
					sErrorMessage = string.Empty;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
				}


				Thread.Sleep(3000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_EXCEPTION;
				sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine(sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
			}
			finally
			{
				if (System.IO.Directory.Exists(EpiaResourceFileEditorFolder))
					System.IO.Directory.Delete(EpiaResourceFileEditorFolder, true);
			}
		}
		#endregion UninstallEpiaResourceFileEditorIfAlreadyInstalled
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region InstallEpiaResourceFileEditor
		public static void InstallEpiaResourceFileEditor(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;
			try
			{
                string path = OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand+ @"\AutomaticTesting\Setup\Epia\Current";
                if (DeployUtilities.getThisPCOS().StartsWith("Windows8.64") || DeployUtilities.getThisPCOS().StartsWith("WindowsServer2012.64"))
                {
                    if (ProjAppInstall.InstallApplicationNet45(path, EgeminApplication.EPIA_RESOURCEFILEEDITOR, EgeminApplication.SetupType.Default, ref sErrorMessage, null/*logger*/))
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        sErrorMessage = EgeminApplication.EPIA_RESOURCEFILEEDITOR + " install failed:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                else
                {
                    if (ProjAppInstall.InstallApplication(path, EgeminApplication.EPIA_RESOURCEFILEEDITOR, EgeminApplication.SetupType.Default, ref sErrorMessage, null/*logger*/))
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                    else
                    {
                        sErrorMessage = EgeminApplication.EPIA_RESOURCEFILEEDITOR + " install failed:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    string EpiaResourceFileEditorFolder = OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand  + "\\Epia Resource File Editor";
                    string file = Path.Combine(EpiaResourceFileEditorFolder, "Egemin.Epia.Foundation.Globalization.ResourceFileEditor.exe");
                    if (System.IO.File.Exists(file))
                        TestCheck = ConstCommon.TEST_PASS;
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "File still not exist after installation:";
                        Console.WriteLine(sErrorMessage);
                        System.Windows.Forms.MessageBox.Show(sErrorMessage);
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
                    Console.WriteLine("\nInstall Epia Resource File Editor.: Pass");
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);

                }
			}
			catch (Exception ex)
			{
                result = ConstCommon.TEST_EXCEPTION;
				sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine(sErrorMessage);
				System.Windows.Forms.MessageBox.Show(sErrorMessage, testname+"-Exception", MessageBoxButtons.OK);
				Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
			}
			finally
			{
				Thread.Sleep(3000);
			}
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region LoadEpiaResourceFiles
		public static void LoadEpiaResourceFiles(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			AutomationElement aeWindow = null;
			AutomationElement aeSelectButton = null;

			try
			{
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 3, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

				Thread executableThread = new Thread(new ThreadStart(EpiaUtilities.StartEpiaResourceFileEditorExecution));
				executableThread.Start();
				Thread.Sleep(5000);

				string mainFormId = "ResourceFileEditorScreen";
				string selectButtonId = "btnSelectDirectory";
				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
				aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
				if (aeWindow == null)
				{
					sErrorMessage = "MainForm not found";
					Console.WriteLine(sErrorMessage);
					TestCheck = ConstCommon.TEST_FAIL;
				}
				else
				{

					if (TestCheck == ConstCommon.TEST_PASS)
					{
						while (aeSelectButton == null && mTime.TotalSeconds <= 120)
						{
							Console.WriteLine("aeSelectButton=");
							aeSelectButton = AUIUtilities.FindElementByID(selectButtonId, aeWindow);
							Thread.Sleep(3000);
							mTime = DateTime.Now - mStartTime;
						}

						if (aeSelectButton == null)
						{
							sErrorMessage = "Select button not found";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							Point pt = AUIUtilities.GetElementCenterPoint(aeSelectButton);
							Input.MoveToAndClick(pt);
							Thread.Sleep(3000);
						}
					}
				}

				String BrowseWindowName = "Browse For Folder";
				AutomationElement aeBrowseWindow = null;
				mStartTime = DateTime.Now;
				mTime = DateTime.Now - mStartTime;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
					while (aeBrowseWindow == null && mTime.TotalSeconds <= 120)
					{
						aeBrowseWindow = AUIUtilities.FindElementByName(BrowseWindowName, aeWindow);
						Thread.Sleep(3000);
						mTime = DateTime.Now - mStartTime;
					}

					if (aeBrowseWindow == null)
					{
						sErrorMessage = "aeBrowseWindow not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						TransformPattern tranform =
						aeBrowseWindow.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
						if (tranform != null)
							tranform.Move(10, 10);

						// find again
						aeBrowseWindow = AUIUtilities.FindElementByName(BrowseWindowName, aeWindow);


					}
				}

				AutomationElement aeTreeView = null;
				AutomationElement aeComputerNode = null;
				string treeViewName = "Tree View";

				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("Browser window is opend -------------- : " + System.DateTime.Now);

					DateTime sTime = DateTime.Now;
					EpiaUtilities.WaitUntilElementByNameFound(aeBrowseWindow, ref aeTreeView, treeViewName, sTime, 60);
					if (aeTreeView == null)
					{
						sErrorMessage = "aeTreeView not found name";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
						sEventEnd = true;
						return;
					}
					else
					{
						aeComputerNode = null;
						TreeWalker walker = TreeWalker.ControlViewWalker;
						AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
						while (elementNode != null)
						{
							Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
							if (elementNode.Current.Name.Equals("Computer"))
							{
								//Input.MoveTo(elementNode);
								aeComputerNode = elementNode;
								Console.WriteLine("Computer node name found , it is: " + aeComputerNode.Current.Name);
								TestCheck = ConstCommon.TEST_PASS;
								break;
							}
							Thread.Sleep(3000);
							elementNode = walker.GetNextSibling(elementNode);
						}
						//return aeNodeLink;
						if (aeComputerNode == null)
						{
							sErrorMessage = "aeComputerNode node not found, ";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							Point pt = AUIUtilities.GetElementCenterPoint(aeComputerNode);
							//Input.MoveToAndClick(pt);
							Thread.Sleep(3000);
						}
					}
				}

				// find C disk node
				AutomationElement aeCDisk = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find " + " (C:)" + " node ===");
					Thread.Sleep(3000);
					System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
					System.Windows.Automation.Condition condition1 = new PropertyCondition(AutomationElement.IsControlElementProperty, true);
					System.Windows.Automation.Condition condition2 = new PropertyCondition(AutomationElement.IsEnabledProperty, true);
					TreeWalker walker = new TreeWalker(new AndCondition(condition1, condition2));
					AutomationElement elementNode = walker.GetFirstChild(aeComputerNode);
					while (elementNode != null)
					{
						Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
						if (elementNode.Current.Name.IndexOf("(C:)") >= 0)
						{
							aeCDisk = elementNode;
							break;
						}
						System.Windows.Forms.TreeNode childTreeNode = treeNode.Nodes.Add(elementNode.Current.ControlType.LocalizedControlType);
						elementNode = walker.GetNextSibling(elementNode);
					}

					if (aeCDisk == null)
					{
						sErrorMessage = "\n=== " + "C disk" + " node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						try
						{
							Console.WriteLine("\n=== " + "C disk " + " node Exist ===");
							ExpandCollapsePattern ep = (ExpandCollapsePattern)aeCDisk.GetCurrentPattern(ExpandCollapsePattern.Pattern);
							ep.Expand();
							Thread.Sleep(3000);
						}
						catch (Exception ex)
						{
							Console.WriteLine("C disk Node can not expaned: " + aeCDisk.Current.Name);
                            Epia3Common.WriteTestLogMsg(slogFilePath, aeCDisk.Current.Name + " C disk Node can not expaned:"+ex.Message, sOnlyUITest);
						}
					}
				}

				// find Program files
				AutomationElement aeFrogramFilesNode = null;
				string programFilesFolderName = TestTools.OSVersionInfoClass.ProgramFilesx86FolderName();
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeFrogramFilesNode = EpiaUtilities.WalkerTreeViewNextChildNede(aeCDisk, programFilesFolderName, ref sErrorMessage);
					if (aeFrogramFilesNode == null)
					{
						sErrorMessage = "\n=== " + programFilesFolderName + " node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
				}

				// find brand
				AutomationElement aeEgeminNode = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeEgeminNode = EpiaUtilities.WalkerTreeViewNextChildNede(aeFrogramFilesNode, sBrand, ref sErrorMessage);
					if (aeEgeminNode == null)
					{
						sErrorMessage = "\n=== " + sBrand + " node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
				}

				// find Epia Server
				AutomationElement aeEpiaServerNode = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeEpiaServerNode = EpiaUtilities.WalkerTreeViewNextChildNede(aeEgeminNode, "Epia Server", ref sErrorMessage);
					if (aeEpiaServerNode == null)
					{
						sErrorMessage = "\n=== " + "Epia Server" + " node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
				}

				// find Data
				AutomationElement aeData = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeData = EpiaUtilities.WalkerTreeViewNextChildNede(aeEpiaServerNode, "Data", ref sErrorMessage);
					if (aeData == null)
					{
						sErrorMessage = "\n=== " + "Data" + " node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
				}

				// find Resource
				AutomationElement aeResource = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeResource = EpiaUtilities.WalkerTreeViewNextChildNede(aeData, "Resources", ref sErrorMessage);
					if (aeResource == null)
					{
						sErrorMessage = "\n=== " + "Resources" + " node NOT Exist ===";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
				}

				// find OK button
				AutomationElement aeOKBtn = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find " + "aeOKBtn" + " Button ===");
					Thread.Sleep(3000);
					mStartTime = DateTime.Now;
					mTime = DateTime.Now - mStartTime;
					if (TestCheck == ConstCommon.TEST_PASS)
					{
						//aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
						while (aeOKBtn == null && mTime.TotalSeconds <= 120)
						{
							aeOKBtn = AUIUtilities.FindElementByName("OK", aeBrowseWindow);
							Thread.Sleep(3000);
							mTime = DateTime.Now - mStartTime;
						}

						if (aeOKBtn == null)
						{
							sErrorMessage = "aeOKBtn not found";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							Point pt = AUIUtilities.GetElementCenterPoint(aeOKBtn);
							Input.MoveToAndClick(pt);
						}
					}
				}

				mStartTime = DateTime.Now;
				mTime = DateTime.Now - mStartTime;
				AutomationElement aeLoadBtn = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					while (aeLoadBtn == null && mTime.TotalSeconds <= 120)
					{
						Console.WriteLine("aeLoadBtn=");
						aeLoadBtn = AUIUtilities.FindElementByName("Load", aeWindow);
						Thread.Sleep(3000);
						mTime = DateTime.Now - mStartTime;
					}

					if (aeLoadBtn == null)
					{
						sErrorMessage = "aeLoadBtn button not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point pt = AUIUtilities.GetElementCenterPoint(aeLoadBtn);
						Input.MoveToAndClick(pt);
						Thread.Sleep(3000);
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
					Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
					result = ConstCommon.TEST_PASS;
				}
			}
			catch (Exception ex)
			{
                result = ConstCommon.TEST_EXCEPTION;
				sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine(sErrorMessage);
				System.Windows.Forms.MessageBox.Show(sErrorMessage, testname + "-Exception", MessageBoxButtons.OK);
				Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
			}
			finally
			{
				Thread.Sleep(3000);
			}
		}
		#endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region LoadEpiaResourceFilesXP
        public static void LoadEpiaResourceFilesXP(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            AutomationElement aeWindow = null;
            AutomationElement aeSelectButton = null;

            try
            {
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 3, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

                Thread executableThread = new Thread(new ThreadStart(EpiaUtilities.StartEpiaResourceFileEditorExecution));
                executableThread.Start();
                Thread.Sleep(5000);

                string mainFormId = "ResourceFileEditorScreen";
                string selectButtonId = "btnSelectDirectory";
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
                if (aeWindow == null)
                {
                    sErrorMessage = "MainForm not found";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {

                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        while (aeSelectButton == null && mTime.TotalSeconds <= 120)
                        {
                            Console.WriteLine("aeSelectButton=");
                            aeSelectButton = AUIUtilities.FindElementByID(selectButtonId, aeWindow);
                            Thread.Sleep(3000);
                            mTime = DateTime.Now - mStartTime;
                        }

                        if (aeSelectButton == null)
                        {
                            sErrorMessage = "Select button not found";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Point pt = AUIUtilities.GetElementCenterPoint(aeSelectButton);
                            Input.MoveToAndClick(pt);
                            Thread.Sleep(3000);
                        }
                    }
                }

                String BrowseWindowName = "Browse For Folder";
                AutomationElement aeBrowseWindow = null;
                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
                    while (aeBrowseWindow == null && mTime.TotalSeconds <= 120)
                    {
                        aeBrowseWindow = AUIUtilities.FindElementByName(BrowseWindowName, aeWindow);
                        Thread.Sleep(3000);
                        mTime = DateTime.Now - mStartTime;
                    }

                    if (aeBrowseWindow == null)
                    {
                        sErrorMessage = "aeBrowseWindow not found";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        TransformPattern tranform =
                        aeBrowseWindow.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                        if (tranform != null)
                            tranform.Move(10, 10);

                        // find again
                        aeBrowseWindow = AUIUtilities.FindElementByName(BrowseWindowName, aeWindow);


                    }
                }

                AutomationElement aeTreeView = null;
                AutomationElement aeComputerNode = null;
                //string treeViewName = "Tree View";

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("Browser window is opend -------------- : " + System.DateTime.Now);
                    System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "SysTreeView32");
                    aeTreeView = aeBrowseWindow.FindFirst(TreeScope.Descendants, condition);

                    DateTime sTime = DateTime.Now;
                    //EpiaUtilities.WaitUntilElementByNameFound(aeBrowseWindow, ref aeTreeView, treeViewName, sTime, 60);
                    if (aeTreeView == null)
                    {
                        sErrorMessage = "aeTreeView not found name";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    else
                    {
                        aeComputerNode = null;
                        TreeWalker walker = TreeWalker.ControlViewWalker;
                        AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
                        while (elementNode != null)
                        {
                            Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
                            if (elementNode.Current.Name.IndexOf("Computer")>= 0 )
                            {
                                //Input.MoveTo(elementNode);
                                aeComputerNode = elementNode;
                                Console.WriteLine("Computer node name found , it is: " + aeComputerNode.Current.Name);
                                TestCheck = ConstCommon.TEST_PASS;
                                break;
                            }
                            Thread.Sleep(3000);
                            elementNode = walker.GetNextSibling(elementNode);
                        }
                        //return aeNodeLink;
                        if (aeComputerNode == null)
                        {
                            sErrorMessage = "aeComputerNode node not found, ";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Point pt = AUIUtilities.GetElementCenterPoint(aeComputerNode);
                            //Input.MoveToAndClick(pt);
                            Thread.Sleep(3000);
                        }
                    }
                }

                // find C disk node
                AutomationElement aeCDisk = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\n=== Find " + " (C:)" + " node ===");
                    Thread.Sleep(3000);
                    System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
                    System.Windows.Automation.Condition condition1 = new PropertyCondition(AutomationElement.IsControlElementProperty, true);
                    System.Windows.Automation.Condition condition2 = new PropertyCondition(AutomationElement.IsEnabledProperty, true);
                    TreeWalker walker = new TreeWalker(new AndCondition(condition1, condition2));
                    AutomationElement elementNode = walker.GetFirstChild(aeComputerNode);
                    while (elementNode != null)
                    {
                        Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
                        if (elementNode.Current.Name.IndexOf("(C:)") >= 0)
                        {
                            aeCDisk = elementNode;
                            break;
                        }
                        System.Windows.Forms.TreeNode childTreeNode = treeNode.Nodes.Add(elementNode.Current.ControlType.LocalizedControlType);
                        elementNode = walker.GetNextSibling(elementNode);
                    }

                    if (aeCDisk == null)
                    {
                        sErrorMessage = "\n=== " + "C disk" + " node NOT Exist ===";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        try
                        {
                            Console.WriteLine("\n=== " + "C disk " + " node Exist ===");
                            ExpandCollapsePattern ep = (ExpandCollapsePattern)aeCDisk.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                            ep.Expand();
                            Thread.Sleep(3000);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("C disk Node can not expaned: " + aeCDisk.Current.Name);
                            Epia3Common.WriteTestLogMsg(slogFilePath, aeCDisk.Current.Name + " C disk Node can not expaned 222:" + ex.Message, sOnlyUITest); ;
                        }
                    }
                }

                // find Program files
                AutomationElement aeFrogramFilesNode = null;
                string programFilesFolderName = TestTools.OSVersionInfoClass.ProgramFilesx86FolderName();
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeFrogramFilesNode = EpiaUtilities.WalkerTreeViewNextChildNede(aeCDisk, programFilesFolderName, ref sErrorMessage);
                    if (aeFrogramFilesNode == null)
                    {
                        sErrorMessage = "\n=== " + programFilesFolderName + " node NOT Exist ===";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // find Brand
                AutomationElement aeEgeminNode = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeEgeminNode = EpiaUtilities.WalkerTreeViewNextChildNede(aeFrogramFilesNode, sBrand, ref sErrorMessage);
                    if (aeEgeminNode == null)
                    {
                        sErrorMessage = "\n=== " + sBrand + " node NOT Exist ===";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // find Epia Server
                AutomationElement aeEpiaServerNode = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeEpiaServerNode = EpiaUtilities.WalkerTreeViewNextChildNede(aeEgeminNode, "Epia Server", ref sErrorMessage);
                    if (aeEpiaServerNode == null)
                    {
                        sErrorMessage = "\n=== " + "Epia Server" + " node NOT Exist ===";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // find Data
                AutomationElement aeData = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeData = EpiaUtilities.WalkerTreeViewNextChildNede(aeEpiaServerNode, "Data", ref sErrorMessage);
                    if (aeData == null)
                    {
                        sErrorMessage = "\n=== " + "Data" + " node NOT Exist ===";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // find Resource
                AutomationElement aeResource = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeResource = EpiaUtilities.WalkerTreeViewNextChildNede(aeData, "Resources", ref sErrorMessage);
                    if (aeResource == null)
                    {
                        sErrorMessage = "\n=== " + "Resources" + " node NOT Exist ===";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // find OK button
                AutomationElement aeOKBtn = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\n=== Find " + "aeOKBtn" + " Button ===");
                    Thread.Sleep(3000);
                    mStartTime = DateTime.Now;
                    mTime = DateTime.Now - mStartTime;
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        //aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
                        while (aeOKBtn == null && mTime.TotalSeconds <= 120)
                        {
                            aeOKBtn = AUIUtilities.FindElementByName("OK", aeBrowseWindow);
                            Thread.Sleep(3000);
                            mTime = DateTime.Now - mStartTime;
                        }

                        if (aeOKBtn == null)
                        {
                            sErrorMessage = "aeOKBtn not found";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Point pt = AUIUtilities.GetElementCenterPoint(aeOKBtn);
                            Input.MoveToAndClick(pt);
                        }
                    }
                }

                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                AutomationElement aeLoadBtn = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    while (aeLoadBtn == null && mTime.TotalSeconds <= 120)
                    {
                        Console.WriteLine("aeLoadBtn=");
                        aeLoadBtn = AUIUtilities.FindElementByName("Load", aeWindow);
                        Thread.Sleep(3000);
                        mTime = DateTime.Now - mStartTime;
                    }

                    if (aeLoadBtn == null)
                    {
                        sErrorMessage = "aeLoadBtn button not found";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Point pt = AUIUtilities.GetElementCenterPoint(aeLoadBtn);
                        Input.MoveToAndClick(pt);
                        Thread.Sleep(3000);
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
                    Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
                    result = ConstCommon.TEST_PASS;
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine(sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
            }
        }
        #endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region FilterResourceFiles
		public static void FilterResourceFiles(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			AutomationElement aeWindow = null;
			try
			{
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 3, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;

                }
				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
				string mainFormId = "ResourceFileEditorScreen";

				//ControlType:	"ControlType.List"
				string filesListId = "lstText";
				AutomationElement aeFilesList = null;

				//ControlType:	"ControlType.ListItem"
				//string resourceFileName = "Epia.Global";


				//ControlType:	"ControlType.Button"
				string btnFilterHideId = "btnFilterHide";    //  Name:	"Filter file / text"
				AutomationElement aeBtnFilterHide = null;

				aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
				if (aeWindow == null)
				{
					sErrorMessage = "MainForm not found";
					Console.WriteLine(sErrorMessage);
					TestCheck = ConstCommon.TEST_FAIL;
				}

				AutomationElement aeRootDirectory = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeWindow.SetFocus();
					// find list
					System.Windows.Automation.Condition cPane = new PropertyCondition(
						AutomationElement.ControlTypeProperty, ControlType.Pane);

					int k = 0;
					mStartTime = DateTime.Now;
					mTime = DateTime.Now - mStartTime;
					while (aeRootDirectory == null && mTime.TotalSeconds <= 120)
					{
						Console.WriteLine("aeAllPanes[k]=" + k++);
						AutomationElementCollection aeAllPanes = aeWindow.FindAll(TreeScope.Children, cPane);

						Console.WriteLine("aeAllPanes.Count=" + aeAllPanes.Count);
						Thread.Sleep(3000);
						for (int i = 0; i < aeAllPanes.Count; i++)
						{
							Console.WriteLine("----   aeAllPanes[" + i + "]=" + aeAllPanes[i].Current.Name);
							if (aeAllPanes[i].Current.Name.Equals("Root Directory:"))
							{
								aeRootDirectory = aeAllPanes[i];
								Console.WriteLine("aeAllPanes[" + i + "]=" + aeRootDirectory.Current.Name);
								break;
							}
						}
						mTime = DateTime.Now - mStartTime;
					}

					if (aeRootDirectory == null)
					{
						sErrorMessage = "aeRootDirectory not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
						Thread.Sleep(3000);
					}
				}

				// find list
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					System.Windows.Automation.Condition cList = new PropertyCondition(
						AutomationElement.ControlTypeProperty, ControlType.List);

					int k = 0;
					mStartTime = DateTime.Now;
					mTime = DateTime.Now - mStartTime;
					while (aeFilesList == null && mTime.TotalSeconds <= 120)
					{
						Console.WriteLine("aeFilesList[k]=" + k++);
						AutomationElementCollection aeAllFilesLists = aeRootDirectory.FindAll(TreeScope.Descendants, cList);
						Console.WriteLine("aeAllFilesLists.Count=" + aeAllFilesLists.Count);

						Thread.Sleep(3000);
						for (int i = 0; i < aeAllFilesLists.Count; i++)
						{
							Console.WriteLine("----   aeAllFilesLists[" + i + "].Current.AutomationId=" + aeAllFilesLists[i].Current.AutomationId);
							if (aeAllFilesLists[i].Current.AutomationId.Equals(filesListId))
							{
								aeFilesList = aeAllFilesLists[i];
								Console.WriteLine("aeAllFilesLists[" + i + "]=" + aeFilesList.Current.Name);
								break;
							}
						}
						mTime = DateTime.Now - mStartTime;
					}

					if (aeFilesList == null)
					{
						sErrorMessage = "aeFilesList not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						while (aeFilesList.Current.IsOffscreen)
						{
							sErrorMessage = "aeFilesList.Current.IsOffscreen";
							Console.WriteLine(sErrorMessage);

							aeBtnFilterHide = AUIUtilities.FindElementByID(btnFilterHideId, aeRootDirectory);
							if (aeBtnFilterHide != null)
							{
								Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeBtnFilterHide));
								aeFilesList = AUIUtilities.FindElementByID(filesListId, aeRootDirectory);
							}
							Thread.Sleep(3000);
						}
					}
				}


				if (TestCheck == ConstCommon.TEST_PASS)
				{
					AutomationElement item = AUIUtilities.FindElementByName("Epia.Global", aeFilesList);
					if (item != null)
					{
						Console.WriteLine("Epia.Global" + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						Thread.Sleep(2000);

						SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
						itemPattern.Select();

						Thread.Sleep(2000);
						Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));
						Thread.Sleep(2000);
					}
					else
					{
						sErrorMessage = "Epia.Global not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
				}

				// find Filter button
				string btnFilterName = "Filter";    //LocalizedControlType:	"button"
				AutomationElement aeBtnFilter = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					Console.WriteLine("\n=== Find " + "btnFilterId" + " Button ===");
					Thread.Sleep(3000);

					mStartTime = DateTime.Now;
					mTime = DateTime.Now - mStartTime;
					if (TestCheck == ConstCommon.TEST_PASS)
					{
						//aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
						while (aeBtnFilter == null && mTime.TotalSeconds <= 120)
						{
							aeBtnFilter = AUIUtilities.FindElementByName(btnFilterName, aeRootDirectory);
							Thread.Sleep(3000);
							mTime = DateTime.Now - mStartTime;
						}

						if (aeBtnFilter == null)
						{
							sErrorMessage = "aeBtnFilter not found";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							Point pt = AUIUtilities.GetElementCenterPoint(aeBtnFilter);
							Input.MoveToAndClick(pt);
						}
					}
				}

				Thread.Sleep(3000);

				// validate filter TODO add ScrollBar elements --> by developer
				// resize columns
				//AutomationElement aeDataGrid = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					resizeColumn("ResourceID");
					resizeColumn("en");
					//resizeColumn("el");
					//resizeColumn("es");
					//resizeColumn("fr");
					//resizeColumn("nl");
					//resizeColumn("pl");
					//resizeColumn("cn");
					//resizeColumn("x");
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
					Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
					result = ConstCommon.TEST_PASS;
				}
			}
			catch (Exception ex)
			{
                result = ConstCommon.TEST_EXCEPTION;
				sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine(sErrorMessage);
				System.Windows.Forms.MessageBox.Show(sErrorMessage, testname + "-Exception", MessageBoxButtons.OK);
				Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
			}
			finally
			{
				Thread.Sleep(3000);
			}
		}
		#endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		static public void resizeColumn(string columnName)  // columnName = "ResourceID"
		{
			AutomationElement aeWindow = EpiaUtilities.GetMainWindow("ResourceFileEditorScreen");
			if (aeWindow == null)
			{
				sErrorMessage = "MainForm not found";
				Console.WriteLine(sErrorMessage);
				TestCheck = ConstCommon.TEST_FAIL;
			}
			else
			{
				aeWindow.SetFocus();
				AutomationElement aeDataGrid = AUIUtilities.FindElementByID("dataGrid", aeWindow);
				//string HeaderName = "ResourceID";
				AutomationElement aeHeaderResource = AUIUtilities.FindElementByName(columnName, aeDataGrid);

				if (aeHeaderResource == null)
				{
					sErrorMessage = "aeHeaderResource not found";
					Console.WriteLine(sErrorMessage);
					TestCheck = ConstCommon.TEST_FAIL;
				}
				else
				{
					Console.WriteLine(" aeHeaderResource width ===" + aeHeaderResource.Current.BoundingRectangle.Width);
					double x = aeHeaderResource.Current.BoundingRectangle.X;
					double y = aeHeaderResource.Current.BoundingRectangle.Y;
					double w = aeHeaderResource.Current.BoundingRectangle.Width;
					double h = aeHeaderResource.Current.BoundingRectangle.Height;
					Point pt = new Point(aeHeaderResource.Current.BoundingRectangle.X + w,
						aeHeaderResource.Current.BoundingRectangle.Y + (h / 2));

					Input.MoveTo(pt);
					Console.WriteLine("move to right");
					Thread.Sleep(3000);

					Input.SendMouseInput(pt.X, pt.Y, 0, SendMouseInputFlags.LeftDown | SendMouseInputFlags.Absolute);

					Console.WriteLine("move to right and click down");
					Thread.Sleep(3000);

					Input.MoveTo(new Point(x + 100, pt.Y));

					Console.WriteLine("move to left");
					Thread.Sleep(3000);
					Input.SendMouseInput(x + 100, pt.Y, 0, SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);

					Console.WriteLine("Left Button up");
					Thread.Sleep(3000);
				}
			}
		}
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region EditAndSaveResourceFiles
		public static void EditAndSaveResourceFiles(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			TestCheck = ConstCommon.TEST_PASS;

			AutomationElement aeWindow = null;
			string mainFormId = "ResourceFileEditorScreen";
			try
			{
                if (ProjBasicUI.IsThisBuildBeforeTestCaseDate(sBuildNr, "Epia 4", "Epia.Production.Release", new DateTime(2012, 3, 8), ref sErrorMessage) == true)
                {
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
				aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
				if (aeWindow == null)
				{
					sErrorMessage = "MainForm not found";
					Console.WriteLine(sErrorMessage);
					TestCheck = ConstCommon.TEST_FAIL;
				}

				AutomationElement aeTabelGridPane = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeWindow.SetFocus();
					// find list
					System.Windows.Automation.Condition cPane = new PropertyCondition(
						AutomationElement.ControlTypeProperty, ControlType.Pane);

					int k = 0;
					mStartTime = DateTime.Now;
					mTime = DateTime.Now - mStartTime;
					while (aeTabelGridPane == null && mTime.TotalSeconds <= 120)
					{
						Console.WriteLine("aeAllPanes[k]=" + k++);
						AutomationElementCollection aeAllPanes = aeWindow.FindAll(TreeScope.Children, cPane);

						Console.WriteLine("aeAllPanes.Count=" + aeAllPanes.Count);
						Thread.Sleep(3000);
						for (int i = 0; i < aeAllPanes.Count; i++)
						{
							Console.WriteLine("----   aeAllPanes[" + i + "].AutomationId=" + aeAllPanes[i].Current.AutomationId);
							if (aeAllPanes[i].Current.AutomationId.Equals("table_Grid"))
							{
								aeTabelGridPane = aeAllPanes[i];
								Console.WriteLine("aeAllPanes[" + i + "]=" + aeTabelGridPane.Current.Name);
								break;
							}
						}
						mTime = DateTime.Now - mStartTime;
					}

					if (aeTabelGridPane == null)
					{
						sErrorMessage = " aeTabelGridPane not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
						Thread.Sleep(3000);
					}
				}


				// find data grid view and update En 
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					AutomationElement aeDataGrid = AUIUtilities.FindElementByID("dataGrid", aeTabelGridPane);
					if (aeDataGrid == null)
					{
						sErrorMessage = "aeDataGrid not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Console.WriteLine("Find aeDataGrid:" + aeDataGrid.Current.AutomationId);
						Thread.Sleep(3000);
						// Construct the Grid Cell Element Name
						string Rowname = "Row 0";
						string cellname = "en Row 0";
						// Get the Element with the Row Col Coordinates
						AutomationElement aeRow0 = AUIUtilities.FindElementByName(Rowname, aeDataGrid);
						if (aeRow0 == null)
						{
							sErrorMessage = "aeRow0 not found" + Rowname;
							Console.WriteLine("aeRow0 failed:" + Rowname);
							TestCheck = ConstCommon.TEST_FAIL;
						}
						else
						{
							Thread.Sleep(2000);
							Console.WriteLine("aeRow0 name is :" + aeRow0.Current.Name);
							AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeRow0);
							if (aeCell == null)
							{
								sErrorMessage = "Find aeCell not found" + cellname;
								Console.WriteLine("Find aeCell failed:" + cellname);
								TestCheck = ConstCommon.TEST_FAIL;
							}
							else
							{
								Console.WriteLine("aeCell name is :" + aeCell.Current.Name);
								Point pt = AUIUtilities.GetElementCenterPoint(aeCell);
								Thread.Sleep(1000);
								Input.MoveTo(pt);
								Thread.Sleep(1000);
								Input.ClickAtPoint(pt);
								Thread.Sleep(1000);
								Input.ClickAtPoint(pt);
								Thread.Sleep(1000);
								Input.ClickAtPoint(pt);
                                Thread.Sleep(1000);
								System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
								Thread.Sleep(1000);
								System.Windows.Forms.SendKeys.SendWait(Constants.RESOURCE_FILE_UPDATED_TEXT);
								Thread.Sleep(1000);
								// Check Field value

								Point ptDataGrid = AUIUtilities.GetElementCenterPoint(aeDataGrid);
								Thread.Sleep(2000);
								Input.MoveTo(ptDataGrid);
								Thread.Sleep(2000);
								Input.ClickAtPoint(ptDataGrid);
								Thread.Sleep(2000);
								Input.ClickAtPoint(ptDataGrid);
								Thread.Sleep(2000);
								Input.ClickAtPoint(ptDataGrid);
							}

						}
					}
				}

				// find Save button
				AutomationElement aeBtnSave = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					while (aeBtnSave == null && mTime.TotalSeconds <= 120)
					{
						Console.WriteLine("aeBtnSave=");
						aeBtnSave = AUIUtilities.FindElementByName("Save", aeTabelGridPane);
						Thread.Sleep(3000);
						mTime = DateTime.Now - mStartTime;
					}

					if (aeBtnSave == null)
					{
						sErrorMessage = "aeBtnSave button not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						Point pt = AUIUtilities.GetElementCenterPoint(aeBtnSave);
						Input.MoveToAndClick(pt);
						Thread.Sleep(3000);
					}
				}



				//AutomationElement aeBtnSave = null;
				AutomationElement aeApplyScreen = null;
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
					if (aeWindow == null)
					{
						sErrorMessage = "MainForm not found";
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						System.Windows.Automation.Condition cWindow = new PropertyCondition(
							AutomationElement.ControlTypeProperty, ControlType.Window);

						int k = 0;
						mStartTime = DateTime.Now;
						mTime = DateTime.Now - mStartTime;
						while (aeApplyScreen == null && mTime.TotalSeconds <= 120)
						{
							Console.WriteLine("aeAllPanes[k]=" + k++);
							AutomationElementCollection aeAllWindows = aeWindow.FindAll(TreeScope.Children, cWindow);

							Console.WriteLine("aeAllWindows.Count=" + aeAllWindows.Count);
							Thread.Sleep(3000);
							for (int i = 0; i < aeAllWindows.Count; i++)
							{
								Console.WriteLine("----   aeAllWindows[" + i + "].AutomationId=" + aeAllWindows[i].Current.AutomationId);
								if (aeAllWindows[i].Current.AutomationId.Equals("ApplyScreen"))
								{
									aeApplyScreen = aeAllWindows[i];
									Console.WriteLine("aeAllWindows[" + i + "]=" + aeApplyScreen.Current.Name);
									break;
								}
							}
							mTime = DateTime.Now - mStartTime;
						}

						if (aeApplyScreen == null)
						{
							sErrorMessage = " aeApplyScreen not found";
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
							Thread.Sleep(3000);
						}
						else
						{
							aeBtnSave = null;
							mStartTime = DateTime.Now;
							mTime = DateTime.Now - mStartTime;
							while (aeBtnSave == null && mTime.TotalSeconds <= 120)
							{
								Console.WriteLine("aeBtnSave=");
								aeBtnSave = AUIUtilities.FindElementByName("Save", aeApplyScreen);
								Thread.Sleep(3000);
								mTime = DateTime.Now - mStartTime;
							}

							if (aeBtnSave == null)
							{
								sErrorMessage = "aeBtnSave button not found";
								Console.WriteLine(sErrorMessage);
								TestCheck = ConstCommon.TEST_FAIL;
							}
							else
							{
								Point pt = AUIUtilities.GetElementCenterPoint(aeBtnSave);
								Input.MoveToAndClick(pt);
								Thread.Sleep(3000);
								Input.MoveTo(pt);
								Console.WriteLine("Click button aeBtnSave=");
								Thread.Sleep(3000);
								Input.MoveToAndClick(pt);
								Thread.Sleep(3000);
							}
						}
					}
				}

				// validate update
				if (TestCheck == ConstCommon.TEST_PASS)
				{
					string epiaDataResourceFolder = OSVersionInfoClass.ProgramFilesx86() + "\\" + sBrand+ "\\Epia Server\\Data\\Resources";
					string fileName = "Epia.Global_en.resources";
					DirectoryInfo DirInfo = new DirectoryInfo(epiaDataResourceFolder);
					FileInfo[] epiaDataResourceFolderFiles = DirInfo.GetFiles(fileName);

					mStartTime = DateTime.Now;
					mTime = DateTime.Now - mStartTime;
					while (epiaDataResourceFolderFiles.Length == 0 && mTime.TotalSeconds <= 120)
					{
						Console.WriteLine(epiaDataResourceFolder + " has no saved resource file yet:" + fileName);
						Thread.Sleep(3000);
						epiaDataResourceFolderFiles = DirInfo.GetFiles(fileName);
						mTime = DateTime.Now - mStartTime;
					}


					if (epiaDataResourceFolderFiles.Length == 0)
					{
						sErrorMessage = epiaDataResourceFolder + " has no resource file after two minutes:" + fileName;
						Console.WriteLine(sErrorMessage);
						TestCheck = ConstCommon.TEST_FAIL;
					}
					else
					{
						StreamReader readerInfo = File.OpenText(Path.Combine(epiaDataResourceFolder, fileName));
						string info = readerInfo.ReadToEnd();
						readerInfo.Close();
						if (info.IndexOf(Constants.RESOURCE_FILE_UPDATED_TEXT) < 0)
						{
							sErrorMessage = epiaDataResourceFolder + " has no resource file:" + fileName;
							Console.WriteLine(sErrorMessage);
							TestCheck = ConstCommon.TEST_FAIL;
						}
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
					Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, sOnlyUITest);
					result = ConstCommon.TEST_PASS;
				}

				//Close main form
				// because save take long time should wait until save finished
				aeWindow = EpiaUtilities.GetMainWindow(mainFormId);
				if (aeWindow != null)
				{
					ProcessUtilities.CloseProcess( "Egemin.Epia.Foundation.Globalization.ResourceFileEditor" );
					//AutomationElement aeTitleBar = AUIUtilities.FindElementByType(ControlType.TitleBar, aeWindow);
					//if (aeTitleBar != null)
					//{
					//     AutomationElement aeClose = AUIUtilities.FindElementByID("Close", aeWindow);
					//    if (aeClose != null)
					//    {
					//        Input.MoveToAndClick(aeClose);
					//    }
					//}
				}
			}
			catch (Exception ex)
			{
                result = ConstCommon.TEST_EXCEPTION;
				sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine(sErrorMessage);
				System.Windows.Forms.MessageBox.Show(sErrorMessage, testname + "-Exception", MessageBoxButtons.OK);
				Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
			}
			finally
			{
				Thread.Sleep(3000);
			}
		}
		#endregion
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
				Thread.Sleep(5000);
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					  AutomationElement.RootElement,
					  UIALayoutStandardScreenEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
				while (root.Current.IsEnabled)
				{
					Console.WriteLine("re click my layout :");
					Input.MoveToAndClick(Pnt);
					Thread.Sleep(5000);
				}

				DateTime mStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalSeconds);

				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
				}

				Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
					  AutomationElement.RootElement,
					  UIALayoutValidateDefaultEventHandler);

				Console.WriteLine("time is:" + mTime.TotalSeconds);
				Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalSeconds, sOnlyUITest);

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
		#endregion Test Cases ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Event ------------------------------------------------------------------------------------------------
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnWaitLogonFormErrorEvent
        public static void OnWaitLogonFormErrorEvent(object src, AutomationEventArgs args)
		{
            Console.WriteLine("OnWaitLogonFormErrorEvent-Begin");
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
            string automationId = "";
			if (element == null)
				name = "null";
			else
			{
				name = element.GetCurrentPropertyValue(
					AutomationElement.NameProperty) as string;
                automationId = element.Current.AutomationId;
			}

			if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnWaitLogonFormErrorEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
			Console.WriteLine(str);

			Thread.Sleep(5000);
            if (name.Equals("Egemin Shell"))
			{
                if (element.Current.AutomationId.Equals("ErrorScreen"))
                {
                    AutomationElement aeCloseBtn = AUIUtilities.FindElementByID("m_BtnClose", element);
                    if (aeCloseBtn != null)
                    {
                        Console.WriteLine("Click Close Button ...");
                        AUIUtilities.ClickElement(aeCloseBtn);
                    }

                    /*AutomationElement aeBtn = AUIUtilities.FindElementByID("m_BtnDetails", element);
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
                    }*/  
                }
			}
			else
			{
				Console.WriteLine("Name is ------------:" + name);
				//AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
			}
            Console.WriteLine("OnWaitLogonFormErrorEvent-End");
		}
		#endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnWaitLogonFormErrorEventXP
        public static void OnWaitLogonFormErrorEventXP(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnWaitLogonFormErrorEventXP-Begin");
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
            string automationId = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
                automationId = element.Current.AutomationId;
            }

            //System.Windows.Forms.MessageBox.Show("automationId" + automationId, "name:" + name);

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("OnWaitLogonFormErrorEventXP:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            Thread.Sleep(5000);
            if (name.Equals("Egemin Shell"))
            {
                if (element.Current.AutomationId.Equals("ErrorScreen"))
                {
                    AutomationElement aeCloseBtn = AUIUtilities.FindElementByID("m_BtnClose", element);
                    if (aeCloseBtn != null)
                    {
                        Console.WriteLine("Click Close Button ...");
                        AUIUtilities.ClickElement(aeCloseBtn);
                    }

                    /*AutomationElement aeBtn = AUIUtilities.FindElementByID("m_BtnDetails", element);
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
                    }*/
                }
            }
            else
            {
                Console.WriteLine("Name is ------------:" + name);
                //AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
            }
            Console.WriteLine("OnWaitLogonFormErrorEventXP-End");
        }
        #endregion
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region OnUIAShellEvent
		public static void OnUIAShellEvent(object src, AutomationEventArgs args)
		{
			Console.WriteLine("OnUIAShellEvent-Begin");
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
					{
						Thread.Sleep(5000);
					}
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
						Thread.Sleep(9000);
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
					Console.WriteLine("OnLayoutResizeUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
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
                    int k = 0;
                    bool check = false;
                    while (check == false && k++ < 5)
                    {
                        check = AUIUtilities.FindElementAndToggle("fullScreenCheckBox", element, ToggleState.On);
                        Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + ChkFullScreenID, sOnlyUITest);
                        Thread.Sleep(2000);
                    }
					
					if (check)
						Thread.Sleep(3000);
					else
					{
                        Console.WriteLine("OnLayoutFullScreenUIAEvent: FindElementAndToggle failed:" + ChkFullScreenID);
                        Epia3Common.WriteTestLogFail(slogFilePath, "OnLayoutFullScreenUIAEvent: FindElementAndToggle failed:" + ChkFullScreenID, sOnlyUITest);
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
                    /*
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
					}*/

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
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region OnConfigSecurityUIAEvent
		public static void OnConfigSecurityUIAEvent(object src, AutomationEventArgs args)
		{
			Console.WriteLine("OnConfigSecurityUIAEvent-Begin");
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
            string id = "Dialog - Egemin.Epia.Modules.RnD.Screens.ShellConfigurationDetailsScreen";

			if (element.Current.AutomationId.Equals(id))
			{
                Console.WriteLine("window with id:"+id+" Found");
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
							= AUIUtilities.FindElementByName("Generic user or Windows user", aeCombo);
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
					Console.WriteLine("OnConfigSecurityUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
					sErrorMessage = ex.Message + " --- " + ex.StackTrace;
					TestCheck = ConstCommon.TEST_FAIL;
					sEventEnd = true;
				}
			}
			sEventEnd = true;
		}
		#endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OnConfigSecurityGenericUserUIAEvent
        public static void OnConfigSecurityGenericUserUIAEvent(object src, AutomationEventArgs args)
		{
            Console.WriteLine("OnConfigSecurityGenericUserUIAEvent-Begin");
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
            string str = string.Format("OnConfigSecurityGenericUserUIAEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
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
							= AUIUtilities.FindElementByName("Generic user", aeCombo);
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
					Console.WriteLine("OnConfigSecurityUIAEvent :" + ex.Message + " --- " + ex.StackTrace);
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
                    bool check = false;
                    int k = 0;
                    while (check == false && k++ < 5)
                    {
                        check = AUIUtilities.FindElementAndToggle("fullScreenCheckBox", element, ToggleState.Off);
                        Thread.Sleep(1000);
                    }
					
					if (check)
						Thread.Sleep(500);
					else
					{
                        Console.WriteLine("OnLayoutStandardScreenUIAEvent: FindElementAndToggle failed:" + ChkFullScreenID);
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
                    /*
					string ChkShowRibbonID = "showRibbonCheckBox"; //-------------- Ribbon OFF ---------------------------
                    bool checkRibbon = false;
                    int kx = 0;
                    while (check == false && kx++ < 5)
                    {
                        checkRibbon = AUIUtilities.FindElementAndToggle(ChkShowRibbonID, element, ToggleState.Off);
                        Thread.Sleep(1000);
                    }

					if (checkRibbon)
						Thread.Sleep(300);
					else
					{
                        sErrorMessage = "OnLayoutStandardScreenUIAEvent:FindElementAndToggle failed:" + ChkShowRibbonID;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
						TestCheck = ConstCommon.TEST_FAIL;
						sEventEnd = true;
						return;
					}
                    */
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
					Console.WriteLine(" time is :" + mTime.TotalSeconds);
					while (aeCombo == null && mTime.TotalSeconds <= 120)
					{
						Thread.Sleep(2000);
						mTime = DateTime.Now - mStartTime;
						Console.WriteLine("find time is:" + mTime.TotalSeconds);
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
		#endregion Event +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region TestLog ----------------------------------------------------------------------------------------------
		public static void SendEmail(string resultFile)
		{
            /*string str1 = "<html><body><b><center>Test Overview</center></b><br>" + "<br>" +
                        "<table col=" +
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
						+ "</td><td></td>	<td></td></tr></table><br>" 
                        + PCName + "<br>"
                        + sCurrentPlatform + "<br>" 
                        + "-------"+"<br></body></html>";
            */

            string str1 = DeployUtilities.GetTestReportContentString(sTotalCounter, sTotalPassed, sTotalFailed, sTotalException, sTotalUntested,
                sCurrentPlatform, sInstallMsiDir); // AnyCPU 

            /*string TextStatistics = "       Test Overview   " + System.Environment.NewLine;
            TextStatistics = TextStatistics + "Total Test Cases:     " + sTotalCounter + System.Environment.NewLine;
            TextStatistics = TextStatistics + "Pass:                 " + sTotalPassed + System.Environment.NewLine;
            TextStatistics = TextStatistics + "Fail:                 " + sTotalFailed + System.Environment.NewLine;
            TextStatistics = TextStatistics + "Exception:            " + sTotalException + System.Environment.NewLine;
            TextStatistics = TextStatistics + "Untested:             " + sTotalUntested + System.Environment.NewLine;
            TextStatistics = TextStatistics + System.Environment.NewLine;

            TextStatistics = str1;*/
            // sBuildDef = CI, Nightly... 
			// sTestApp = Epia for other Appliaction is layout
			// resultFile = xls file
			// sTotalFailed
			ProcessUtilities.SendTestResultToDevelopers( resultFile, sTestApp, sBuildDef, logger, sTotalFailed,
				sBuildNr/*used for email title*/, str1/*content*/, sSendMail);
		}
		#endregion TestLog +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	}
}
