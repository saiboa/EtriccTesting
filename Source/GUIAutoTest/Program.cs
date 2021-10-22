using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Configuration;
using System.Diagnostics;
using System.ServiceProcess;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using TestTools;
using Excel = Microsoft.Office.Interop.Excel;

namespace Epia4GUIAutoTest
{
#pragma warning disable 0162 // Disable_MainForm_Toolbars_Dock_Area_Top warning for Unreachable Code Detected.
    class Program
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
		static string[] sTestCaseName       = new string[100];
		static DateTime sTestStartUpTime    = DateTime.Now;
		static DateTime sStartTime          = DateTime.Now;
		static string sTestApp              = string.Empty;
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
		// PCinfo
		static public string PCName;
		static public string OSName;
		static public string OSVersion;
		static public string UICulture;
		static public string TimeOnPC;
		// Build info
		static string sBuildDropFolder = string.Empty;
		static string sBuildConfig  = string.Empty;
		static string sBuildDef     = string.Empty;
		static string sBuildNr      = string.Empty;
		static string sTestToolsVersion = string.Empty;
        static string m_SystemDrive = @"C:\";
        static string UserPassword      = "Egemin01";
        static string sTargetPlatform = string.Empty;
        static string sCurrentPlatform = string.Empty;
        static string sTestResultFolder = string.Empty;
		// Testcase not used =================================
		public static string sConfigurationName = string.Empty;
		static string sErrorMessage;
		static bool sEventEnd;
		static string sExcelVisible         = string.Empty;
		static bool sAutoTest               = true;
		static string sInstallMsiDir    = string.Empty;
		public static string sLayoutName    = string.Empty;
        static string sServerRunAs = "Service";
		static bool sDemo;
		static string sSendMail = "false";
		static string sTFSServer = "http://teamApplication.teamSystems.egemin.be:8080";
		
		// LOG=================================================================
		public static string slogFilePath = @"C:\";
        static string sOutFilename = "OutFilename";
		static string sOutFilePath = string.Empty;
		static StreamWriter Writer;
		// Build param ========================================================
		//static BuildStore   buildStore      = null;

		static IBuildServer m_BuildSvc;
		static bool         TFSConnected    = true;
		// excel 	--------------------------------------------------------
		static Excel.Application    xApp;
		static Excel.Workbook       xBook;
		static Excel.Workbooks      xBooks;
		static Excel.Range          xRange;
		//static Excel.Worksheet      xSheet;
        static dynamic xSheet;
		// default layout
		public const ToggleState DEFAULT_FULLSCREEN     = ToggleState.Off;
		public const ToggleState DEFAULT_MAXIMIZED      = ToggleState.Off;
		public const ToggleState DEFAULT_ALLOWRESIZE    = ToggleState.On;
		public const string DEFAULT_XPOS    = "0";
		public const string DEFAULT_YPOS    = "0";
		public const string DEFAULT_WIDTH   = "792";
		public const string DEFAULT_HEIGHT  = "606";
		public const string DEFAULT_TITLE   = "Egemin Shell";
		public const ToggleState DEFAULT_SHOW_RIBBON    = ToggleState.Off;
		public const ToggleState DEFAULT_SHOW_MAINENU   = ToggleState.On;
		public const ToggleState DEFAULT_SHOW_TOOLBARS  = ToggleState.Off;
		public const ToggleState DEFAULT_SHOW_NAVIGATOR = ToggleState.On;
		public const ToggleState DEFAULT_STACK_SCREENS  = ToggleState.Off;
		#endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
        static string sEpiaServerFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server";
        static string sEpiaShellFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell";
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
			}
			catch (Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Check PC-info");
			}

           
            sOnlyUITest = false;
            string x = System.Configuration.ConfigurationManager.AppSettings.Get("OnlyUITest");
            if (x.ToLower().StartsWith("true"))
                sOnlyUITest = true;

            UserPassword = System.Configuration.ConfigurationManager.AppSettings.Get("CurrentUserPassword");

            if (!sOnlyUITest)
            {
                try
                {
                    // validate inputs
                    if (args != null)
                    {
                        for (int i = 0; i <= 15; i++)
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

                        sTestResultFolder = sBuildDropFolder + "\\TestResults";
                        if (!System.IO.Directory.Exists(sTestResultFolder))
                            System.IO.Directory.CreateDirectory(sTestResultFolder);

                        //Epia3Common.CreateOutputFileInfo(args, PCName, ref sOutFilePath, ref sOutFilename);
                        CreateOutputFileInfo(args, sCurrentPlatform, PCName, ref sOutFilePath, ref sOutFilename);

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
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message + "---" + ex.StackTrace, "Validate command-line params");
                }
            }
            else
            {   // if only UI test, the msi folder is from the config file
                sInstallMsiDir = System.Configuration.ConfigurationManager.AppSettings.Get("EpiaInstallMsiDirectory");
                if ( ! File.Exists(Path.Combine(sInstallMsiDir, "Epia.msi")) )
                    MessageBox.Show("Please select a folder where Epia.msi is included");
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
                                TestTools.MessageBoxEx.Show("Team Foundation services are not available from server\nWill try to reconnect the Server after 10 minutes",
                                kTime++ + " During E'pia UI Testing, please not touch the screen, time :" + DateTime.Now.ToLongTimeString(), 10 * 60000);
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
						MessageBox.Show(ex.Message);
						Epia3Common.WriteTestLogMsg(slogFilePath, "Get TFS Server:" + ex.Message,sOnlyUITest);
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
            sTestCaseName[2] = START_EPIA_SERVER_SHELL;
			sTestCaseName[3] = LAYOUT_FIND_LAYOUT_PANEL;
			sTestCaseName[4] = LAYOUT_INITIAL_X_POSITION;
			sTestCaseName[5] = LAYOUT_INITIAL_Y_POSITION;
			sTestCaseName[6] = LAYOUT_INITIAL_WIDTH;
			sTestCaseName[7] = LAYOUT_INITIAL_HEIGHT;
			sTestCaseName[8] = LAYOUT_ALLOW_RESIZE;
			sTestCaseName[9] = LAYOUT_FULL_SCREEN;
			sTestCaseName[10] = LAYOUT_MAXIMIZED;
			sTestCaseName[11] = LAYOUT_RIBBON_ON;
			//sTestCaseName[9] = LAYOUT_TITLE;
			sTestCaseName[12] = LAYOUT_CANCEL_BUTTON;
            sTestCaseName[13] = SECURITY_OPEN_ACCOUNT_DETAIL_SCREEN;
            sTestCaseName[14] = SECURITY_ADD_NEW_ROLE;
            sTestCaseName[15] = SECURITY_OPEN_ROLE_DETAIL_SCREEN;
            sTestCaseName[16] = SECURITY_EDIT_ROLE;
            sTestCaseName[17] = MULTI_LANGUAGE_CHECK;
            sTestCaseName[18] = SETTING_LANGUAGE_NL;
			sTestCaseName[19] = SHELL_CONFIGURATION_SECURITY;
			sTestCaseName[20] = LOGON_CURRENT_USER;
			//sTestCaseName[14] = LOGON_EPIA_ADMINISTRATOR;
			sTestCaseName[21] = SHELL_SHUTDOWN;
			sTestCaseName[22] = SHELL_LOGOFF;
            sTestCaseName[23] = SHELL_MULTIPLE_LOGIN; ;
			sTestCaseName[24] = EPIA4_CLOSEE;
            sTestCaseName[25] = EPIA4_CLEAN_UNINSTALL_CHECK;
            sTestCaseName[26] = EPIA4_CLEAN_SHELL_INSTALL_CHECK;
            sTestCaseName[27] = EPIA4_CLEAN_SERVER_INSTALL_CHECK;
            sTestCaseName[28] = UNINSTALL_EPIA_RESOURCE_EDITOR_IF_ALREADY_INSTALLED;
            sTestCaseName[29] = INSTALL_EPIA_RESOURCE_EDITOR;
            sTestCaseName[30] = LOAD_EPIA_RESOURCE_FILES;
            sTestCaseName[31] = FILTER_EPIA_RESOURCE_FILES;
            sTestCaseName[32] = EDIT_SAVE_RESOURCE_FILES;
			//=============================================//
			try
			{
				if (!sOnlyUITest)
				{
					TestTools.Utilities.CloseProcess("EXCEL");
					TestTools.Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
					TestTools.Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
					Thread.Sleep(1000);
				}

				// Excel file not for EpiaTestPC3 and EPIATESTSERVER3
				if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
					|| PCName.ToUpper().Equals("EPIATESTSRV3V1") )
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
					xSheet.Cells[1, 2] = "EPIA UI Test Scenarios";
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
                    if (sInstallMsiDir != null)
					{
						xSheet.Cells[9, 1] = "Build Location:";
                        xSheet.Cells[9, 2] = sInstallMsiDir;
					}
				}

				// start test----------------------------------------------------------
				int sResult = ConstCommon.TEST_UNDEFINED;
				int aantal = 33;
				if ( sDemo )
					aantal = 2;

                if (sOnlyUITest)
                {
                    sTestType = System.Configuration.ConfigurationManager.AppSettings.Get("TestType");
                    if (sTestType.ToLower().StartsWith("all"))
                    {
                        aantal = 33;
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
                    }
                }
                else
                {
                    if (TFSConnected)
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
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", "Epia4+" + sCurrentPlatform+"Normal");
                            Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.EPIA, sOnlyUITest);
                        }
                    }
                }

				while (Counter < aantal)
				{
					sResult = ConstCommon.TEST_UNDEFINED;
					switch (sTestCaseName[Counter])
					{
                        case EPIA_SERVER_INSTALLATION_INTEGRITY:
                            EpiaServerIntegrityCheck(EPIA_SERVER_INSTALLATION_INTEGRITY, aeForm, out sResult);
                            break;
                        case EPIA_SHELL_INSTALLATION_INTEGRITY:
                            EpiaShellIntegrityCheck(EPIA_SHELL_INSTALLATION_INTEGRITY, aeForm, out sResult);
                            break;
                        case START_EPIA_SERVER_SHELL:
                            StartEpiaServerShell(START_EPIA_SERVER_SHELL, aeForm, out sResult);
                            break;
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
                        case SECURITY_OPEN_ACCOUNT_DETAIL_SCREEN:
                            SecurityOpenAccountDetailScreen(SECURITY_OPEN_ACCOUNT_DETAIL_SCREEN, aeForm, out sResult);
                            break;
                        case SECURITY_ADD_NEW_ROLE:
                            SecurityAddNewRole(SECURITY_OPEN_ROLE_DETAIL_SCREEN, aeForm, out sResult);
                            break;
                        case SECURITY_OPEN_ROLE_DETAIL_SCREEN:
                            SecurityOpenRoleDetailScreen(SECURITY_OPEN_ROLE_DETAIL_SCREEN, aeForm, out sResult);
                            break;
                        case SECURITY_EDIT_ROLE:
                            SecurityEditRole(SECURITY_EDIT_ROLE, aeForm, out sResult);
                            break;
						case SETTING_LANGUAGE_NL:
							LanguageSettingNL(SETTING_LANGUAGE_NL, aeForm, out sResult);
							break;
                        case MULTI_LANGUAGE_CHECK:
                            MultiLanguageCheck(MULTI_LANGUAGE_CHECK, aeForm, out sResult);
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
                        case SHELL_MULTIPLE_LOGIN:
                            ShellMultipleLogin(SHELL_MULTIPLE_LOGIN, aeForm, out sResult);
							break;
						case EPIA4_CLOSEE:
							Epia4Close(EPIA4_CLOSEE, aeForm, out sResult);
							break;
                        case EPIA4_CLEAN_UNINSTALL_CHECK:
                            Epia4CleanUninstallCheck(EPIA4_CLEAN_UNINSTALL_CHECK, aeForm, out sResult);
                            break;
                        case EPIA4_CLEAN_SHELL_INSTALL_CHECK:
                            Epia4CleanShellInstallCheck(EPIA4_CLEAN_SHELL_INSTALL_CHECK, aeForm, out sResult);
                            break;
                        case EPIA4_CLEAN_SERVER_INSTALL_CHECK:
                            Epia4CleanServerInstallCheck(EPIA4_CLEAN_SERVER_INSTALL_CHECK, aeForm, out sResult);
                            break;
                        case UNINSTALL_EPIA_RESOURCE_EDITOR_IF_ALREADY_INSTALLED:
                            UninstallEpiaResourceFileEditorIfAlreadyInstalled(UNINSTALL_EPIA_RESOURCE_EDITOR_IF_ALREADY_INSTALLED, aeForm, out sResult);
                            break;
                        case INSTALL_EPIA_RESOURCE_EDITOR:
                            InstallEpiaResourceFileEditor(INSTALL_EPIA_RESOURCE_EDITOR, aeForm, out sResult);
                            break;
                        case LOAD_EPIA_RESOURCE_FILES:
                            LoadEpiaResourceFiles(LOAD_EPIA_RESOURCE_FILES, aeForm, out sResult);
                            break;
                        case FILTER_EPIA_RESOURCE_FILES:
                            FilterResourceFiles(FILTER_EPIA_RESOURCE_FILES, aeForm, out sResult);
                            break;
                        case EDIT_SAVE_RESOURCE_FILES:
                            EditAndSaveResourceFiles(EDIT_SAVE_RESOURCE_FILES, aeForm, out sResult);
                            break;
					}

					if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
						|| PCName.ToUpper().Equals("EPIATESTSRV3V1"))
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

				Epia3Common.WriteTestLogMsg(slogFilePath, "Total Tests: " + Counter, sOnlyUITest);
				Epia3Common.WriteTestLogMsg(slogFilePath, "Total Passed: " + sTotalPassed, sOnlyUITest);
				Epia3Common.WriteTestLogMsg(slogFilePath, "Total Failed: " + sTotalFailed, sOnlyUITest);

				if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
					|| PCName.ToUpper().Equals("EPIATESTSRV3V1"))
					Epia3Common.WriteTestLogMsg(slogFilePath, "No Excel due to: " + PCName, sOnlyUITest);
				else
				{
                    xSheet.Cells[Counter + 2 + 8, 1] = "Total tests: ";
					xSheet.Cells[Counter + 3 + 8, 1] =  "Total Passes: ";
					xSheet.Cells[Counter + 4 + 8, 1] =  "Total Failed: ";

					xSheet.Cells[Counter + 2 + 8, 2] = sTotalCounter;
					xSheet.Cells[Counter + 3 + 8, 2] = sTotalPassed;
                    xSheet.Cells[Counter + 4 + 8, 2] = sTotalFailed;

					ulong TPhysicalMem = 0;
					ulong APhysicalMem = 0;
					ulong TVirtualMem = 0;
					ulong AVirtualMem = 0;

					HelpUtilities.GetMemoryInfo(out TPhysicalMem, out APhysicalMem, out TVirtualMem, out AVirtualMem);
					// Add Legende
					xSheet.Cells[Counter + 5 + 8, 2] = "Legende";
                    xRange = xApp.get_Range("B" + (Counter + 5 + 8));
                    xRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
					xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;                    
					xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

					xSheet.Cells[Counter + 6 + 8, 2] = "Pass";
                    xRange = xApp.get_Range("B" + (Counter + 6 + 8));
					xRange.Interior.ColorIndex = ConstCommon.EXCEL_GREEN;
					xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

					xSheet.Cells[Counter + 7 + 8, 2] = "Fail";
                    xRange = xApp.get_Range("B" + (Counter + 7 + 8));
					xRange.Interior.ColorIndex = ConstCommon.EXCEL_RED;
					xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

					xSheet.Cells[Counter + 8 + 8, 2] = "Exception";
                    xRange = xApp.get_Range("B" + (Counter + 8 + 8));
					xRange.Interior.ColorIndex = ConstCommon.EXCEL_PINK;
					xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

					xSheet.Cells[Counter + 9 + 8, 2] = "Untested";
                    xRange = xApp.get_Range("B" + (Counter + 9 + 8));
					xRange.Interior.ColorIndex = ConstCommon.EXCEL_YELLOW;
					xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;
			
					xSheet.Cells[Counter + 10 + 9, 2] = "TotalPhysicalMemory:" + TPhysicalMem + " MB";
					xSheet.Cells[Counter + 11 + 9, 2] = "AvailablePhysicalMemory:" + APhysicalMem + " MB";
					xSheet.Cells[Counter + 12 + 9, 2] = "TotalVirtualMemory:" + TVirtualMem + " MB";
                    xSheet.Cells[Counter + 13 + 9, 2] = "AvailableVirtualMemory:" + AVirtualMem + " MB";
				}

				if (!sOnlyUITest)
				{
					if (TFSConnected)
					{
						Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                            TestTools.TfsUtilities.GetProjectName(ConstCommon.EPIA), sBuildNr);
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
                           
                            while (infoline != null && infoline.Length > 0 )
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
                                    TestTools.TfsUtilities.GetProjectName(ConstCommon.EPIA),
                                    "GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
                            }*/
                            //---------
							Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
						}
						else
						{
							if (sTotalFailed == 0)
								TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName(ConstCommon.EPIA),
								"GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
							else
								TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName(ConstCommon.EPIA),
								"GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
						}

						// update testinfo file
						string testout = "-->" + sOutFilename + ".xls";
						if (sAutoTest)
						{
							if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
								|| PCName.ToUpper().Equals("EPIATESTSRV3V1"))
								testout = "-->" + sOutFilename + ".log";

                            if (sTotalFailed == 0)
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Passed", "Epia4+" + sCurrentPlatform + "Normal");
                                Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Passed:" + ConstCommon.EPIA, sOnlyUITest);
                            }
                            else
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Failed", "Epia4+" + sCurrentPlatform + "Normal");
                                Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Failed:" + ConstCommon.EPIA, sOnlyUITest);
                            }
						}
					}

					if (sAutoTest)
						TestTools.FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
				}

                if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
                    || PCName.ToUpper().Equals("EPIATESTSRV3V1"))
                    Epia3Common.WriteTestLogMsg(slogFilePath, "No Excel due to: " + PCName, sOnlyUITest);
                else
                {
                    xSheet.Columns.AutoFit();
                    xSheet.Rows.AutoFit();
                }

				// save Excel to Local machine
				string sXLSPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
					sOutFilename + ".xls");
				// Save the Workbook locally  --- not for PC EPIATESTPC3 and EPIATESTSERVER3
				object missing = System.Reflection.Missing.Value;
				if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
					|| PCName.ToUpper().Equals("EPIATESTSRV3V1"))
					Console.WriteLine("No Excel due to: " + PCName);
				else
				{
                    Console.WriteLine("Save1 : " + sXLSPath);
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
						string sXLSPath2 = System.IO.Path.Combine(sOutFilePath, sOutFilename + ".xls");
                        Console.WriteLine("Save2 : " + sXLSPath);
						Epia3Common.WriteTestLogMsg(slogFilePath, "sXLSPath2 =: " + sXLSPath2, sOnlyUITest);
						xBook.SaveAs(sXLSPath2, Excel.XlFileFormat.xlWorkbookNormal,
									missing, missing, missing, missing,
									Excel.XlSaveAsAccessMode.xlNoChange,
									missing, missing, missing, missing, missing);
					}
					// quit Excel.
					if (xBook != null) xBook.Close(true, sOutFilename, false);
					if (xBooks != null) xBooks.Close();
					xApp.Quit();

					// Send Result via Email
                    if (!sOnlyUITest)
					    SendEmail(sXLSPath);
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
                TestTools.Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                Thread.Sleep(10000);
				TestTools.Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
				TestTools.Utilities.CloseProcess("cmd");
				Console.WriteLine("\nEnd test run\n");
				//Console.ReadLine();
			}
			catch (Exception ex)
			{
				Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
				Thread.Sleep(2000);
				if (sAutoTest)
				{

                    FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Exception -->" + sOutFilename + ".log", "Epia4+" + sCurrentPlatform + "Normal");
                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Exception -->" + sOutFilename + ".log:" + ConstCommon.EPIA, sOnlyUITest);

					Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace, sOnlyUITest);

					Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
					Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
					Utilities.CloseProcess("cmd");
                    FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");

					if (TFSConnected)
					{
						Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
							TestTools.TfsUtilities.GetProjectName("Epia"), sBuildNr);
						string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

						if (quality.Equals("GUI Tests Failed"))
						{
							Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
						}
						else
						{
							TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
								TestTools.TfsUtilities.GetProjectName("Epia"),
								"GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
						}

					}
				}
			}
		}
		#endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region TestCase Name
        private const string EPIA_SERVER_INSTALLATION_INTEGRITY = "EpiaServerInstallerIntegrityCheck";
        private const string EPIA_SHELL_INSTALLATION_INTEGRITY = "EpiaShellInstallerIntegrityCheck";
        private const string START_EPIA_SERVER_SHELL = "StartEpiaServerShell";
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
		private const string SETTING_LANGUAGE_NL = "LanguageSettingNL";
        private const string MULTI_LANGUAGE_CHECK = "MultiLanguageCheck";
		private const string SHELL_CONFIGURATION_SECURITY = "ShellConfigSecurity";
		private const string LOGON_CURRENT_USER = "LogonCurrentUser";
		private const string LOGON_EPIA_ADMINISTRATOR = "LogonEpiaAdmin";
		private const string SHELL_SHUTDOWN = "ShellShutdown";
		private const string SHELL_LOGOFF = "ShellLogOff";
        private const string SHELL_MULTIPLE_LOGIN = "ShellMultipleLogin";
		private const string EPIA4_CLOSEE = "Epia4Close";
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
                
                // already converted to DOTNET 4 build
                string[] EpiaServerDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaServerFolderFiles.txt"));
                if (sInstallMsiDir.IndexOf("Net4") > 0
                    || sInstallMsiDir.IndexOf("Dev01") > 0
                    || sInstallMsiDir.IndexOf("Dev02") > 0
                     || sInstallMsiDir.IndexOf("Dev03.") > 0
                       || sInstallMsiDir.IndexOf("Dev08.") > 0
                    || sInstallMsiDir.IndexOf("Main") > 0
                     || sInstallMsiDir.IndexOf("Production") > 0
                    )
                    EpiaServerDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaServerFolderNET4Files.txt"));

                // already implemented Service function
                if (sInstallMsiDir.IndexOf("Dev02") > 0
                    || sInstallMsiDir.IndexOf("Dev01") > 0
                  || sInstallMsiDir.IndexOf("Dev03.") > 0
                    || sInstallMsiDir.IndexOf("Dev05") > 0
                    || sInstallMsiDir.IndexOf("Dev08") > 0
                     || sInstallMsiDir.IndexOf("Net4") > 0
                   || sInstallMsiDir.IndexOf("Production") > 0
                   || sInstallMsiDir.IndexOf("Main") > 0
                    )
                {

                    List<string> NewList = new List<string>();
                    for (int i = 0; i < EpiaServerDlls.Length; i++)
                    {
                        NewList.Add(EpiaServerDlls[i]);
                    }
                    NewList.Add("Egemin.Epia.Foundation.WindowsService.Interfaces.dll");
                    EpiaServerDlls = NewList.ToArray();

                    dlls = string.Empty;
                    for (int i = 0; i < EpiaServerDlls.Length; i++)
                    {
                        dlls = dlls + System.Environment.NewLine + EpiaServerDlls[i];
                    }
                    //System.Windows.Forms.MessageBox.Show(dlls, "installedServerDllCnt: " + EpiaServerDlls.Length);
                }

                // compare dlls in Server folder
                int installedServerDllCnt = installedDllsName.Length;
                int standardServerDllCnt = EpiaServerDlls.Length;

                if (sOnlyUITest)
                    System.Windows.Forms.MessageBox.Show("installedServerDllCnt: " + installedServerDllCnt + Environment.NewLine +dlls,
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
                string[] EpiaShellDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaShellFolderFiles.txt"));
                if (sInstallMsiDir.IndexOf("Net4") > 0
                     || sInstallMsiDir.IndexOf("Dev01") > 0
                      || sInstallMsiDir.IndexOf("Dev02") > 0
                     || sInstallMsiDir.IndexOf("Dev03.") > 0
                     || sInstallMsiDir.IndexOf("Dev08") > 0
                      || sInstallMsiDir.IndexOf("Main") > 0
                      || sInstallMsiDir.IndexOf("Production") > 0
                    )
                    EpiaShellDlls = System.IO.File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "EpiaShellFolderNET4Files.txt"));

               // alreadyimplemented Service
                if (sInstallMsiDir.IndexOf("Dev02") > 0
                    || sInstallMsiDir.IndexOf("Dev01") > 0
                     || sInstallMsiDir.IndexOf("Dev03.") > 0
                   || sInstallMsiDir.IndexOf("Dev05") > 0
                   || sInstallMsiDir.IndexOf("Dev08") > 0
                    || sInstallMsiDir.IndexOf("Net4") > 0
                   || sInstallMsiDir.IndexOf("Production") > 0
                   || sInstallMsiDir.IndexOf("Main") > 0
                   )
                {
                    List<string> NewList = new List<string>();
                    for (int i = 0; i < EpiaShellDlls.Length; i++)
                    {
                        NewList.Add(EpiaShellDlls[i]);
                    }
                    NewList.Add("Egemin.Epia.Foundation.WindowsService.Interfaces.dll");
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
                     || sInstallMsiDir.IndexOf("Main") > 0
                   )
                {
                    List<string> NewList = new List<string>();
                    for (int i = 0; i < EpiaShellDlls.Length; i++)
                    {
                        if (EpiaShellDlls[i].IndexOf("Win.UltraWinExplorerBar") < 0)
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

                // compare dlls in Server folder
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
                    sErrorMessage = "Too much files installed : extra installed dll --> " + sErrorMessage;
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
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
                //========================   SERVER =================================================
                #region SERVER
                Console.WriteLine("sServerRunAs : " + sServerRunAs);
                if (sServerRunAs.ToLower().IndexOf("service") >= 0)
                {
                    // uninstall Egemin.Epia.server Service
                    Console.WriteLine("UNINSTALL EPIA SERVER Service : ");
                    TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() +
                        @"\Egemin\Epia Server",
                        ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /u");
                    Thread.Sleep(2000);

                    Console.WriteLine("INSTALL EPIA SERVER Service : ");
                    TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() +
                        @"\Egemin\Epia Server",
                        ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /i");
                    Thread.Sleep(2000);

                    Console.WriteLine("Start EPIA SERVER as Service : ");
                    TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() +
                        @"\Egemin\Epia Server",
                        ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /start");
                    Thread.Sleep(2000);

                    ServiceController svcEpia = new ServiceController("Egemin Epia Server");
                    Console.WriteLine(svcEpia.ServiceName + " has status " + svcEpia.Status.ToString());
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

                    //svcEpia.WaitForStatus(ServiceControllerStatus.Running);
                    if (svcEpia.Status != ServiceControllerStatus.Running)
                    {
                        sErrorMessage = "Epia Service start up failed:";
                        Epia3Common.WriteTestLogMsg(slogFilePath, "Epia Service start up failed: " + epiaServiceStatus, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                        //throw new Exception("Epia service start up failed:"); //   get message from log file sErrorMessage//
                    }
                }
                else if (sServerRunAs.ToLower().IndexOf("console") >= 0)
                {
                    Console.WriteLine("Start EPIA Server as console applications : ");
                    // Start Epia SERVER as Console
                    TestTools.Utilities.StartProcessNoWait(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() +
                        @"\Egemin\Epia Server",
                        ConstCommon.EGEMIN_EPIA_SERVER_EXE, string.Empty);

                    Thread.Sleep(20000);
                    Console.WriteLine("Epia SERVER Started : ");
                }
                #endregion
                Thread.Sleep(5000);
                //========================   SHELL =================================================
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("EPIA SERVER Service Started : ");
                    Thread.Sleep(2000);

                    // Add Open window Event Handler
                    Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);
                    sEventEnd = false;
                    #region  Shell
                    TestTools.Utilities.StartProcessNoWait(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName() +
                        @"\Egemin\Epia Shell",
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

                System.Windows.Forms.MessageBox.Show("where ");
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
		#region LayoutFindLayoutPanel
		public static void LayoutFindLayoutPanel(string testname, AutomationElement root, out int result)
		{
            AutomationEventHandler UIAFindLayoutPanelEventHandler = new AutomationEventHandler(OnFindLayoutPanelUIAEvent);
			try
			{
                if (sOnlyUITest)
                    root = EpiaUtilities.GetMainWindow("MainForm");

                Console.WriteLine("\n=== Test " + testname + " === " + root.Current.Name);
                Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest);
                result = ConstCommon.TEST_UNDEFINED;
                


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
                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutXPosEventHandler);

                int k = 0;
                while ( k < 5 )
                {
                    #region
                root = EpiaUtilities.GetMainWindow("MainForm");
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
				Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
				while (sEventEnd == false && mTime.TotalSeconds <= 600)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - mStartTime;
					Console.WriteLine("wait time is :" + mTime.TotalMilliseconds);
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
					double Height = 700;
					TransformPattern tranform =
					root.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                    if (tranform != null)
                    {
                        tranform.Move(0, 0);
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
        #region SecurityOpenAccountDetailScreen
		public static void SecurityOpenAccountDetailScreen(string testname, AutomationElement rootxxx, out int result)
		{
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            AutomationElement aeWindow = null;
            AutomationElement aeSecurityAccounts = null;
            try
            {
                aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                if (aeWindow != null)
                {
                    EpiaUtilities.ClearDisplayedScreens(aeWindow);
                    aeSecurityAccounts = AUICommon.FindTreeViewNodeLevel1(testname, aeWindow, "Security", "Accounts", ref sErrorMessage);
                    if (aeSecurityAccounts == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Point Pnt = AUIUtilities.GetElementCenterPoint(aeSecurityAccounts);
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
                        Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {   // open detail screen
                        if (EpiaUtilities.WindowMenuAction(aeSelectedWindow, "Account type", 0, "Details...", ref sErrorMessage))
                            TestCheck = ConstCommon.TEST_PASS;
                        else
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                // check error screen is displayed
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

                //  check detail screen is opened
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                    AutomationElementCollection aeAllWindows = null;
                    // find child windows
                    System.Windows.Automation.Condition cWindows = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Window);
                    aeAllWindows = aeWindow.FindAll(TreeScope.Descendants, cWindows);
                    Console.WriteLine("aeAllWindows.Count:   " + aeAllWindows.Count);
                    if (aeAllWindows.Count < 2)
                    {
                        sErrorMessage = "Child windows count < 2";
                        Console.WriteLine(sErrorMessage);
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
        #region SecurityAddNewRole
        public static void SecurityAddNewRole(string testname, AutomationElement rootXXX, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
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
                            double x = aeButtonAdd.Current.BoundingRectangle.Right + 70.0;
                            double y =(aeButtonAdd.Current.BoundingRectangle.Bottom + aeButtonAdd.Current.BoundingRectangle.Top)/2.0;
                            Point pt = new Point( x, y );
                            Input.MoveTo(pt);
                            Input.ClickAtPoint(pt);
                            Thread.Sleep(3000);
                        }


                        /*if (EpiaUtilities.WindowMenuAction(aeSelectedWindow, "Account type", 0, "Details...", ref sErrorMessage))
                            TestCheck = ConstCommon.TEST_PASS;
                        else
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }*/
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
                        if (AUIUtilities.FindTextBoxAndChangeValue("nameTextBox", aeRoleAddEditDialog, out origRole, "RoleA", ref sErrorMessage))
                            Thread.Sleep(3000);
                        else
                        {
                            sErrorMessage = "FindTextBoxAndChangeValue failed:" + "nameTextBox";
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }

                        if (TestCheck == ConstCommon.TEST_PASS)
                        {
                            // ControlType:	"ControlType.Edit"
                            //AutomationId:	"descriptionTextBox"
                            //LocalizedControlType:	"edit"
                            //Name:	"Description:"

                            if (AUIUtilities.FindTextBoxAndChangeValue("descriptionTextBox", aeRoleAddEditDialog, out origRole, "DescriptionA", ref sErrorMessage))
                                Thread.Sleep(3000);
                            else
                            {
                                sErrorMessage = "FindTextBoxAndChangeValue failed:" + "descriptionTextBox";
                                Console.WriteLine(sErrorMessage);
                                Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
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
                        string cellname = "Name" + " Row " + 0;
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
                            else if (!cellValue.Equals("RoleA"))
                            {
                                sErrorMessage = "aeCell Value not equal RoleA   , but :" + cellValue;
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
                result = ConstCommon.TEST_FAIL;
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
        #region SecurityOpenRoleDetailScreen
        public static void SecurityOpenRoleDetailScreen(string testname, AutomationElement rootxxx, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            AutomationElement aeWindow = null;
            AutomationElement aeSecurityAccounts = null;
            try
            {
                aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                if (aeWindow != null)
                {
                    EpiaUtilities.ClearDisplayedScreens(aeWindow);
                    aeSecurityAccounts = AUICommon.FindTreeViewNodeLevel1(testname, aeWindow, "Security", "Roles", ref sErrorMessage);
                    if (aeSecurityAccounts == null)
                    {
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        Point Pnt = AUIUtilities.GetElementCenterPoint(aeSecurityAccounts);
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
                        Console.WriteLine("Selected Window not opened :" + sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        if (EpiaUtilities.WindowMenuAction(aeSelectedWindow, "Name", 0, "Details...", ref sErrorMessage))
                            TestCheck = ConstCommon.TEST_PASS;
                        else
                        {
                            Console.WriteLine(sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }


                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                    AutomationElementCollection aeAllWindows = null;
                    // find child windows
                    System.Windows.Automation.Condition cWindows = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Window);
                    aeAllWindows = aeWindow.FindAll(TreeScope.Descendants, cWindows);
                    Console.WriteLine("aeAllWindows.Count:   " + aeAllWindows.Count);
                    if (aeAllWindows.Count < 2)
                    {
                        sErrorMessage = "Child windows count < 2";
                        Console.WriteLine(sErrorMessage);
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
        #region SecurityEditRole
        public static void SecurityEditRole(string testname, AutomationElement rootXXX, out int result)
        {
            Console.WriteLine("=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
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
                            Console.WriteLine("aeCell not find :" + aeSelectedWindow.Current.Name+"---"+sErrorMessage);
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
                    Console.WriteLine("\nTest scenario" +testname +" : Pass");
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
		public static void LanguageSettingNL(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			AutomationEventHandler UIALanguageSettingEventHandler = new AutomationEventHandler(OnLanguageSettingNLUIAEvent);

			try
			{
                if ( sOnlyUITest)
                    root = EpiaUtilities.GetMainWindow("MainForm");
               

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

            try
            {
                string epiaDataResourceFolder = OSVersionInfoClass.ProgramFilesx86() + "\\Egemin\\Epia Server\\Data\\Resources";
                //string resourceFileName = "Epia.Modules.RnD_cn.resources";
                string[] resourceFileNames = { "Epia.Modules.RnD_cn.resources","Epia.Modules.RnD_es.resources",
                                                 "Epia.Modules.RnD_de.resources","Epia.Modules.RnD_el.resources",
                                                 "Epia.Modules.RnD_fr.resources","Epia.Modules.RnD_nl.resources",
                                                 "Epia.Modules.RnD_pl.resources","Epia.Modules.RnD_en.resources"};


                for (int i = 0; i < resourceFileNames.Length; i++ )
                {
                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        if (EpiaUtilities.SwitchLanguageAndFindText(epiaDataResourceFolder, resourceFileNames[i], ref sErrorMessage))
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
		#region ShellConfigSecurity
		public static void ShellConfigSecurity(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
			AutomationEventHandler UIAConfigSecurityEventHandler = new AutomationEventHandler(OnConfigSecurityUIAEvent);

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
                //while (root.Current.IsEnabled)
                //{
                    Console.WriteLine("re click Shell Config Security :");
                    Input.MoveToAndClick(shellPoint);
                    Thread.Sleep(5000);
                //}

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
					   AutomationElement.RootElement, UIAConfigSecurityEventHandler);

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
                //if (sOnlyUITest)
                //    root = EpiaUtilities.GetMainWindow("MainForm");

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
				string UserNameID   = "m_TextBoxUsername"; //"ControlType.Edit" Name : "with Windows credentials
				string PasswordID   = "m_TextBoxPassword";
				string BtnOKID      = "m_BtnOK";
		
				string origUser = string.Empty;
				string tester = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

				

				if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser,  UserPassword, ref sErrorMessage))
					Thread.Sleep(3000);
				else
				{
					sErrorMessage = "FindTextBoxAndChangeValue failed:" + PasswordID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname, sOnlyUITest);
					result = ConstCommon.TEST_FAIL;
					return;
				}

                if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester, ref sErrorMessage))
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
				string path = Path.Combine(m_SystemDrive+ConstCommon.EPIA_CLIENT_ROOT, 
					ConstCommon.EGEMIN_EPIA_SHELL_EXE);
				System.Diagnostics.Process proc = System.Diagnostics.Process.Start(path);
				Console.WriteLine("*****" + proc.Id);
                Thread.Sleep(3000);
				
                // Start Shell
                //TestTools.Utilities.StartProcessNoWait(m_SystemDrive + ConstCommon.EPIA_CLIENT_ROOT,
                //    ConstCommon.EGEMIN_EPIA_SHELL_EXE, string.Empty);

				

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

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIACurrentUserEventHandler);

				Console.WriteLine("Application aeSecurityForm name : " + aeSecurityForm.Current.Name);
				string UserNameID = "m_TextBoxUsername"; //"ControlType.Edit" Name : "with Windows credentials
				string PasswordID = "m_TextBoxPassword";
				string BtnOKID = "m_BtnOK";

				string origUser = string.Empty;
				string tester = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
         

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

				if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester, ref sErrorMessage))
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
                Thread.Sleep(9000);

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
            
                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser, UserPassword, ref sErrorMessage))
                    Thread.Sleep(3000);
                else
                {
                    sErrorMessage = "FindTextBoxAndChangeValue failed:" + PasswordID;
                    Console.WriteLine(sErrorMessage);
                    result = ConstCommon.TEST_FAIL;
                    return;
                }

				if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester, ref sErrorMessage))
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
		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ShellMultipleLogin
        public static void ShellMultipleLogin(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            string shellToolbarsID = "_MainForm_Toolbars_Dock_Area_Top";
            AutomationElement aeSecurityForm = null;
            AutomationElement aeShellToolbars = null;
            double x = 0;
            double y = 0;
            Point shellPoint = new Point(x, y);
            try
            {
                Random rnd = new Random(); 
                
                string loginCount = System.Configuration.ConfigurationManager.AppSettings.Get("MultipleLoginCount");
                Int64 intLoginCount = Convert.ToInt64(loginCount);
                Console.WriteLine("login count : " + intLoginCount);
                int k = 0;
                while (TestCheck == ConstCommon.TEST_PASS && k++ < intLoginCount)
                {
                    root = EpiaUtilities.GetMainWindow("MainForm");
                    Console.WriteLine("find shell toolbars -----"+k);
                    aeShellToolbars = AUIUtilities.FindElementByID(shellToolbarsID, root);
                    if (aeShellToolbars == null)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = shellToolbarsID + "not found";
                        Console.WriteLine(sErrorMessage);
                    }
                    else
                    {
                        x = aeShellToolbars.Current.BoundingRectangle.Left + 5;
                        y = (aeShellToolbars.Current.BoundingRectangle.Top + aeShellToolbars.Current.BoundingRectangle.Bottom) / 2;
                        
                        shellPoint = new Point(x, y);
                        Input.MoveTo(shellPoint);
                        Thread.Sleep(2000);
                        Input.MoveToAndClick(shellPoint);
                        Thread.Sleep(3000);
                        Point shutdownPoint = new Point(x+20, y + 75);
                        Input.MoveTo(shutdownPoint);
                        Thread.Sleep(2000);
                        Input.MoveToAndClick(shutdownPoint);
                        Thread.Sleep(3000);

                        if ( intLoginCount < 10 )
                            Thread.Sleep(30000);

                        //-----------------shell shuted down----
                        AutomationEventHandler UIACurrentUserEventHandler = new AutomationEventHandler(OnUIACurrentUserEvent);
                        // Add Open window Event Handler
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement, TreeScope.Descendants, UIACurrentUserEventHandler);
                        
                        sEventEnd = false;
                        TestCheck = ConstCommon.TEST_PASS;

                        //string path = Path.Combine(sInstallScriptsDir, Constants.SHELL_BAT);
                        string path = Path.Combine(m_SystemDrive + ConstCommon.EPIA_CLIENT_ROOT,
                            ConstCommon.EGEMIN_EPIA_SHELL_EXE);
                        System.Diagnostics.Process proc = System.Diagnostics.Process.Start(path);
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
                        aeSecurityForm = null;
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
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "LogonForm not found";
                            Console.WriteLine(sErrorMessage);
                        }

                    }

                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        //string tester = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                        string tester = "tfstest0" + rnd.Next(1, 9); ; 
                        if (EpiaUtilities.ProcessSecurityForm(aeSecurityForm, tester, UserPassword, ref sErrorMessage))
                        {
                            Thread.Sleep(5000);
                            DateTime mStartTime = DateTime.Now;
                            TimeSpan mTime = DateTime.Now - mStartTime;
                            root = EpiaUtilities.GetMainWindow("MainForm");
                            while (root == null && mTime.TotalMilliseconds < 120000)
                            {
                                Thread.Sleep(2000);
                                root = EpiaUtilities.GetMainWindow("MainForm");
                                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                                mTime = DateTime.Now - mStartTime;
                            }

                            Console.WriteLine("find shell toolbars =====");
                            aeShellToolbars = AUIUtilities.FindElementByID(shellToolbarsID, root);
                            if (aeShellToolbars == null)
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = shellToolbarsID + " not found";
                                Console.WriteLine(sErrorMessage);
                            }
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = " ProcessSecurityForm failed: "+ sErrorMessage;
                            Console.WriteLine(sErrorMessage);
                        }
                    }

                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        x = aeShellToolbars.Current.BoundingRectangle.Left + 5;
                        y = (aeShellToolbars.Current.BoundingRectangle.Top + aeShellToolbars.Current.BoundingRectangle.Bottom) / 2;
                        shellPoint = new Point(x, y);
                        Input.MoveTo(shellPoint);
                        Thread.Sleep(2000);
                        Input.MoveToAndClick(shellPoint);
                        Thread.Sleep(3000);

                        double y2 = y + 50;
                        Point logoffPoint = new Point(x, y2);
                        Input.MoveTo(logoffPoint);
                        Thread.Sleep(5000);
                        Input.MoveToAndClick(logoffPoint);
                        Thread.Sleep(3000);

                        aeSecurityForm = null;
                        DateTime mStartTime1 = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - mStartTime1;

                        while (aeSecurityForm == null && mTime.TotalMilliseconds < 120000)
                        {
                            aeSecurityForm = AUIUtilities.FindElementByID("LogonForm", AutomationElement.RootElement);
                            Thread.Sleep(2000);
                            Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                            mTime = DateTime.Now - mStartTime1;
                        }

                        if (aeSecurityForm != null)
                        {
                            //string tester = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                            string tester = "tfstest0" + rnd.Next(1, 9); ; 
                            if (EpiaUtilities.ProcessSecurityForm(aeSecurityForm, tester, UserPassword, ref sErrorMessage))
                            {
                                Thread.Sleep(3000);
                            }
                            else
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = " ProcessSecurityForm failed: " + sErrorMessage;
                                Console.WriteLine(sErrorMessage);
                            }
                        }
                        else
                        {
                            TestCheck = ConstCommon.TEST_FAIL;
                            sErrorMessage = "LogonForm not found";
                            Console.WriteLine(sErrorMessage);
                        }
                    }    
                } // end while

                if (TestCheck == ConstCommon.TEST_PASS)
                {
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
		#endregion Epia4Close
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Epia4CleanUninstallCheck
        public static void Epia4CleanUninstallCheck(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
			result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            AutomationElement aeEpia = null;
            
            string sYesButtonName = "Yes";
            string sCloseButtonName = "Close";
            string EpiaServerFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server";
            string EpiaShellFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell";
            Thread.Sleep(5000);
            try
            {
                Utilities.CloseProcess("Egemin.Epia.Shell");
                Utilities.CloseProcess("Egemin.Epia.Server");

                #region Uninstall Epia
                Thread executableThread = new Thread(new ThreadStart(EpiaUtilities.StartProgramsAndFeaturesExecution));
                executableThread.Start();
                Thread.Sleep(5000);

                Console.WriteLine("Searching for Programs and Features main window:" + EpiaUtilities.GetProgramsFeaturesScreenNaam());
                System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, EpiaUtilities.GetProgramsFeaturesScreenNaam());
                AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                if (appElement != null)
                {   // (1) Programs and Features main window
                    Console.WriteLine("Programs and Features main window opend");
                    Thread.Sleep(2000);
                    Console.WriteLine("Searching programs item button...");
                    AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(appElement, Constants.PROGRAMS_FEATURES_FOLDER_VIEW_ID);
                    if (aeGridView != null)
                        Console.WriteLine("Gridview found...");
                    Thread.Sleep(2000);
                    // Set a property condition that will be used to find the control.
                    System.Windows.Automation.Condition c = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.DataItem);

                    AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children, c);
                    Console.WriteLine("Programs count ..." + aeProgram.Count);
                    for (int i = 0; i < aeProgram.Count; i++)
                    {
                        Console.WriteLine("programs name: " + aeProgram[i].Current.Name);
                        if (aeProgram[i].Current.Name.StartsWith("E'pia Fr"))
                            aeEpia = aeProgram[i];
                    }

                    if (aeEpia == null)
                    {
                        Console.WriteLine("No Epia name: ");
                        AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(appElement, sCloseButtonName);
                        if (btnClose != null)
                        {
                            AUIUtilities.ClickElement(btnClose);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Epia name: " + aeEpia.Current.Name);
                        string x = aeEpia.Current.Name;
                        Thread.Sleep(5000);
                        // click on Epia item 
                        InvokePattern pattern = aeEpia.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        pattern.Invoke();
                        Thread.Sleep(2000);
                        #region Program and Features dialog
                        // find Program and Features dialog (in the future, do not show me this dialog box possible)
                        AutomationElement dialogElement = appElement.FindFirst(TreeScope.Children, condition);
                        if (dialogElement != null)
                        {
                            Thread.Sleep(3000);
                            AUIUtilities.MoveUIElement(dialogElement, 0, 0);
                            Thread.Sleep(3000);
                            AutomationElement btnYes = AUIUtilities.GetElementByNameProperty(appElement, sYesButtonName);
                            if (btnYes != null)
                            {
                                AUIUtilities.ClickElement(btnYes);
                            }
                        }
                        #endregion

                        #region // Window Installer section
                        //AutomationElement installerElement
                        //    = GetElementByNameProperty(rootElement, "Windows Installer");

                        //if (installerElement != null)
                        //    Console.WriteLine("Uninstaller dialog found ...");

                        // wait until application uninstalled
                        DateTime startTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - startTime;
                        bool hasApplication = EpiaUtilities.IsApplicationInstalled("Epia");
                        while (hasApplication == true && mTime.TotalMilliseconds < 120000)
                        {
                            Thread.Sleep(8000);
                            mTime = DateTime.Now - startTime;
                            if (mTime.TotalMilliseconds > 120000)
                            {
                                System.Windows.Forms.MessageBox.Show("Uninstall Epia run timeout " + mTime.TotalMilliseconds);
                                break;
                            }
                            hasApplication = EpiaUtilities.IsApplicationInstalled("Epia");
                        }
                        #endregion
                        // close Features and Programs window    
                        Console.WriteLine("close Features and Programs window----------------------");
                        appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                        Thread.Sleep(2000);
                        AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(appElement, sCloseButtonName);
                        if (btnClose != null)
                            AUIUtilities.ClickElement(btnClose);

                        Console.WriteLine("---------- Epia Uninstalled ----------");
                    }
                }
                else
                {
                    sErrorMessage = "---------- Uninstalled main window not found  ----------" + EpiaUtilities.GetProgramsFeaturesScreenNaam();
                    Console.WriteLine("---------- Uninstalled main window not found  ----------" + EpiaUtilities.GetProgramsFeaturesScreenNaam());
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                #endregion

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
                                System.Windows.Forms.MessageBox.Show(sErrorMessage);
                            }
                        }
                        else if (System.IO.Directory.Exists(EpiaShellFolder))
                        {
                            // get files in ShellFolder
                            DirectoryInfo DirInfo = new DirectoryInfo(EpiaShellFolder);
                            FileInfo[] shellFolderFiles = DirInfo.GetFiles("*.*");
                            if (shellFolderFiles.Length > 0)
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = EpiaShellFolder + " still has some files:" + shellFolderFiles[0].FullName;
                                Console.WriteLine(sErrorMessage);
                                System.Windows.Forms.MessageBox.Show(sErrorMessage);
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
                    Console.WriteLine("do nothing, not test case:"+testname);
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
                     System.IO.Directory.Delete(EpiaServerFolder, true);

                 if (System.IO.Directory.Exists(EpiaShellFolder))
                     System.IO.Directory.Delete(EpiaShellFolder, true);
            }
		}
        #endregion Epia4CleanUninstallCheck

        static private void Wait(int seconds)
        {
            System.Threading.Thread.Sleep(seconds * 1000);
        }

        static private void StartEpiaApplicationExecution()
        {
            Process Proc = new System.Diagnostics.Process();
            Proc.StartInfo.FileName = Path.Combine(sInstallMsiDir, "Epia.msi");  
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

            DateTime testCaseCreateDate = new DateTime(2012, 3, 8);
            
            if (sBuildNr.IndexOf("Hotfix") >= 0)
            {
                System.Windows.Forms.MessageBox.Show("sBuildNr:" + sBuildNr);
                //string buildnr = "Hotfix 4.3.2.1 of Epia.Production.Hotfix_20120405.1";
                string hotfixVersion = EpiaUtilities.getReleaseFromHotfixVersion(sBuildNr);
                DateTime dt = EpiaUtilities.getReleaseVersionDate(hotfixVersion);

                System.Windows.Forms.MessageBox.Show("dt:" + dt, "testCaseCreateDate:" + testCaseCreateDate);

                if (dt < testCaseCreateDate)
                {
                    sErrorMessage = "Release date of this hotfix is earlier then this test case created date, Not test";
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }
                else
                    System.Windows.Forms.MessageBox.Show("dt > testCaseCreateDate:");
            }

            AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
            // Add Open window Event Handler
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);
            

			try
			{
                Console.WriteLine("Start the MSI");
                Thread executableThread = new Thread(new ThreadStart(StartEpiaApplicationExecution));
                executableThread.Start();
                Wait(15);

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement, UIAShellEventHandler);

                Console.WriteLine("Start Epia Installation -------------->");
                bool status = EpiaUtilities.InstallEpia("Shell", ref sErrorMessage);

                if (status == false)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    System.Windows.Forms.MessageBox.Show(sErrorMessage);
                }

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
                            sErrorMessage = sEpiaShellFolder + " has no installed files:" + shellFolderFiles[0].FullName;
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
			}
		}
        #endregion Epia4CleanShellInstallCheck
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Epia4CleanServerInstallCheck
        public static void Epia4CleanServerInstallCheck(string testname, AutomationElement root, out int result)
        {

            Epia4CleanUninstallCheck("NoTestcase", AutomationElement.RootElement, out result);


            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            DateTime testCaseCreateDate = new DateTime(2012, 3, 8);
            if (sBuildNr.IndexOf("Hotfix") >= 0)
            {
                //string buildnr = "Hotfix 4.3.2.1 of Epia.Production.Hotfix_20120405.1";
                string hotfixVersion = EpiaUtilities.getReleaseFromHotfixVersion(sBuildNr);
                DateTime dt = EpiaUtilities.getReleaseVersionDate(hotfixVersion);

                if (dt < testCaseCreateDate)
                {
                    sErrorMessage = "Release date of this hotfix is earlier then this test case created date, Not test";
                    result = ConstCommon.TEST_UNDEFINED;
                    return;
                }
            }

            AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
            // Add Open window Event Handler
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);


            try
            {
                Console.WriteLine("Start the MSI");
                Thread executableThread = new Thread(new ThreadStart(StartEpiaApplicationExecution));
                executableThread.Start();
                Wait(15);

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                      AutomationElement.RootElement, UIAShellEventHandler);

                Console.WriteLine("Start Epia Installation -------------->");
                bool status = EpiaUtilities.InstallEpia("Server", ref sErrorMessage);

                if (status == false)
                {
                    TestCheck = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    if (sOnlyUITest)
                        System.Windows.Forms.MessageBox.Show(sErrorMessage);
                }

                // validate shell installation
                string EpiaShellFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell";

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (System.IO.Directory.Exists(EpiaShellFolder))
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "error installed, Shell folder exist,";
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
                            sErrorMessage = sEpiaServerFolder + " has no installed files:" + serverFolderFiles[0].FullName;
                            Console.WriteLine(sErrorMessage);
                            if (sOnlyUITest)
                                System.Windows.Forms.MessageBox.Show(sErrorMessage);
                        }
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = "error installed, Server folder not exist,";
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
            AutomationElement aeEpia = null;

            string sYesButtonName = "Yes";
            string sCloseButtonName = "Close";
            string EpiaResourceFileEditorFolder = OSVersionInfoClass.ProgramFilesx86() + "\\Epia Resource File Editor";
            Thread.Sleep(5000);
            try
            {
                Utilities.CloseProcess("Egemin.Epia.Shell");
                Utilities.CloseProcess("Egemin.Epia.Server");

                #region Uninstall Epia
                Thread executableThread = new Thread(new ThreadStart(EpiaUtilities.StartProgramsAndFeaturesExecution));
                executableThread.Start();
                Thread.Sleep(5000);

                Console.WriteLine("Searching for Programs and Features main window:" + EpiaUtilities.GetProgramsFeaturesScreenNaam());
                System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, EpiaUtilities.GetProgramsFeaturesScreenNaam());
                AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                if (appElement != null)
                {   // (1) Programs and Features main window
                    Console.WriteLine("Programs and Features main window opend");
                    Thread.Sleep(2000);
                    Console.WriteLine("Searching programs item button...");
                    AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(appElement, Constants.PROGRAMS_FEATURES_FOLDER_VIEW_ID);
                    if (aeGridView != null)
                        Console.WriteLine("Gridview found...");
                    Thread.Sleep(2000);
                    // Set a property condition that will be used to find the control.
                    System.Windows.Automation.Condition c = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.DataItem);

                    AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children, c);
                    Console.WriteLine("Programs count ..." + aeProgram.Count);
                    for (int i = 0; i < aeProgram.Count; i++)
                    {
                        Console.WriteLine("programs name: " + aeProgram[i].Current.Name);
                        if (aeProgram[i].Current.Name.StartsWith("E'pia Resource"))
                            aeEpia = aeProgram[i];
                    }

                    if (aeEpia == null)
                    {
                        Console.WriteLine("No Epia name: ");
                        AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(appElement, sCloseButtonName);
                        if (btnClose != null)
                        {
                            AUIUtilities.ClickElement(btnClose);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Epia name: " + aeEpia.Current.Name);
                        string x = aeEpia.Current.Name;
                        Thread.Sleep(5000);
                        // click on Epia item 
                        InvokePattern pattern = aeEpia.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        pattern.Invoke();
                        Thread.Sleep(2000);
                        #region Program and Features dialog
                        // find Program and Features dialog (in the future, do not show me this dialog box possible)
                        AutomationElement dialogElement = appElement.FindFirst(TreeScope.Children, condition);
                        if (dialogElement != null)
                        {
                            Thread.Sleep(3000);
                            AUIUtilities.MoveUIElement(dialogElement, 0, 0);
                            Thread.Sleep(3000);
                            AutomationElement btnYes = AUIUtilities.GetElementByNameProperty(appElement, sYesButtonName);
                            if (btnYes != null)
                            {
                                AUIUtilities.ClickElement(btnYes);
                            }
                        }
                        #endregion

                        #region // Window Installer section
                        //AutomationElement installerElement
                        //    = GetElementByNameProperty(rootElement, "Windows Installer");

                        //if (installerElement != null)
                        //    Console.WriteLine("Uninstaller dialog found ...");

                        // wait until application uninstalled
                        DateTime startTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - startTime;
                        bool hasApplication = EpiaUtilities.IsApplicationInstalled("EpiaResourceFileEditor");
                        while (hasApplication == true && mTime.TotalMilliseconds < 120000)
                        {
                            Thread.Sleep(8000);
                            mTime = DateTime.Now - startTime;
                            if (mTime.TotalMilliseconds > 120000)
                            {
                                System.Windows.Forms.MessageBox.Show("Uninstall EpiaResourceFileEditor run timeout " + mTime.TotalMilliseconds);
                                break;
                            }
                            hasApplication = EpiaUtilities.IsApplicationInstalled("EpiaResourceFileEditor");
                        }
                        #endregion
                        // close Features and Programs window    
                        Console.WriteLine("close Features and Programs window----------------------");
                        appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                        Thread.Sleep(2000);
                        AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(appElement, sCloseButtonName);
                        if (btnClose != null)
                            AUIUtilities.ClickElement(btnClose);

                        Console.WriteLine("---------- EpiaResourceFileEditor Uninstalled ----------");
                    }
                }
                else
                {
                    sErrorMessage = "---------- Uninstalled main window not found  ----------" + EpiaUtilities.GetProgramsFeaturesScreenNaam();
                    Console.WriteLine("---------- Uninstalled main window not found  ----------" + EpiaUtilities.GetProgramsFeaturesScreenNaam());
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                #endregion

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
            AutomationEventHandler UIEpiaResourceFileEditorEventHandler = new AutomationEventHandler(OnInstallEpiaResourceFileEditorUIEvent);

            try
            {
                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIEpiaResourceFileEditorEventHandler);

                string InstallerSource = Path.Combine(sInstallMsiDir, "Epia.ResourceFileEditor.msi");
                if ( sOnlyUITest )
                    InstallerSource = Path.Combine(System.Configuration.ConfigurationManager.AppSettings.Get("EpiaInstallMsiDirectory"),
                        "Epia.ResourceFileEditor.msi"); 
                Console.WriteLine("start:" + InstallerSource);
                Process Proc = new System.Diagnostics.Process();
                Proc.StartInfo.FileName = InstallerSource;
                Proc.StartInfo.CreateNoWindow = false;
                Proc.Start();
                Console.WriteLine("started:" + InstallerSource);

                sEventEnd = false;
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                while (sEventEnd == false && mTime.TotalMinutes <= 5)
                {
                    Thread.Sleep(2000);
                    mTime = DateTime.Now - mStartTime;
                    Console.WriteLine("wait time is :" + mTime.TotalSeconds + "   --------- sEventEnd:" + sEventEnd);

                }


                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    string EpiaResourceFileEditorFolder = OSVersionInfoClass.ProgramFilesx86() + "\\Egemin\\Epia Resource File Editor";
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
                    Console.WriteLine("\nInstall Epia Rwesource File Editor.: Pass");
                    result = ConstCommon.TEST_PASS;
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);

                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                sErrorMessage = "Fatal error: " + ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine(sErrorMessage);
                Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIEpiaResourceFileEditorEventHandler);

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

            Thread executableThread = new Thread(new ThreadStart(EpiaUtilities.StartEpiaResourceFileEditorExecution));
            executableThread.Start();
            Thread.Sleep(5000);
            AutomationElement aeWindow = null;
            AutomationElement aeSelectButton = null;

            try
            {
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
                            tranform.Move(10,10);

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
                            if (elementNode.Current.Name.Equals("Computer") )
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
                        }
                    }
                }

                // find Program files
                AutomationElement aeFrogramFilesNode = null;
                string programFilesFolderName = TestTools.OSVersionInfoClass.ProgramFilesx86FolderName();
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeFrogramFilesNode = EpiaUtilities.WalkerTreeViewNextChildNede(aeCDisk, programFilesFolderName,ref sErrorMessage);
                    if (aeFrogramFilesNode == null)
                    {
                        sErrorMessage = "\n=== " + programFilesFolderName + " node NOT Exist ===";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }

                // find Egemin
                AutomationElement aeEgeminNode = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    aeEgeminNode = EpiaUtilities.WalkerTreeViewNextChildNede(aeFrogramFilesNode, "Egemin", ref sErrorMessage);
                    if (aeEgeminNode == null)
                    {
                        sErrorMessage = "\n=== " + "Egemin" + " node NOT Exist ===";
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
                result = ConstCommon.TEST_FAIL;
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
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                string mainFormId = "ResourceFileEditorScreen";

                //ControlType:	"ControlType.List"
                string filesListId = "lstText";
                AutomationElement aeFilesList = null;

                //ControlType:	"ControlType.ListItem"
                string resourceFileName = "Epia.Global";


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
              
                AutomationElement aeDataGrid = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    resizeColumn("ResourceID");
                    resizeColumn("en");
                    resizeColumn("el");
                    resizeColumn("es");
                    resizeColumn("fr");
                    resizeColumn("nl");
                    resizeColumn("pl");
                    resizeColumn("cn");
                    resizeColumn("x");
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
                result = ConstCommon.TEST_FAIL;
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
                    Input.SendMouseInput(x+100, pt.Y, 0, SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);

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
                                Thread.Sleep(2000);
                                Input.MoveTo(pt);
                                Thread.Sleep(2000);
                                Input.ClickAtPoint(pt);
                                Thread.Sleep(2000);
                                Input.ClickAtPoint(pt);
                                Thread.Sleep(2000);
                                Input.ClickAtPoint(pt);
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
                    string epiaDataResourceFolder = OSVersionInfoClass.ProgramFilesx86() + "\\Egemin\\Epia Server\\Data\\Resources";
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
                    Utilities.CloseProcess("Egemin.Epia.Foundation.Globalization.ResourceFileEditor");
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
                result = ConstCommon.TEST_FAIL;
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
        #region OnInstallEpiaResourceFileEditorUIEvent
        public static void OnInstallEpiaResourceFileEditorUIEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnInstallEpiaResourceFileEditorUIEvent");
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
            string str = string.Format("OnInstallEpiaResourceFileEditorUIEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
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
            else if (name.StartsWith("E'pia Resource"))
            {
                AutomationElement rootElement = AutomationElement.RootElement;
                #region Install Epia Resource File Editor
                Console.WriteLine("Searching for main window");
                System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
                AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);

                DateTime startTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - startTime;
                while (appElement == null && mTime.TotalMilliseconds < 60000)
                {
                    Wait(2);
                    appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                    mTime = DateTime.Now - startTime;
                    if (mTime.TotalMilliseconds > 60000)
                    {
                        System.Windows.Forms.MessageBox.Show("After one minute no Installer Window Form found");
                        return;
                    }
                }

                if (appElement != null)
                {   // (1) Welcom Main window
                    Console.WriteLine("---------- Welcom to Epia Resource File Editor window opend...");
                    Console.WriteLine("Searching next button...");
                    AutomationElement btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                    if (btnNext != null)
                    {   // (2) Select Folder
                        AUIUtilities.ClickElement(btnNext);
                        appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                        Console.WriteLine("---------- Select Folder window opend ...");
                        Console.WriteLine("Searching checkbox...");
                        Wait(2);
                        Console.WriteLine("Searching next button...");
                        btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                        if (btnNext != null)
                        {
                            AUIUtilities.ClickElement(btnNext);
                            appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                            Console.WriteLine("---------- Installation Folders window opend ...");
                            Console.WriteLine("Searching next button...");
                            btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                            if (btnNext != null)
                            {   // (2) Confirm 
                                AUIUtilities.ClickElement(btnNext);
                                appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                Console.WriteLine("---------- Confirm Installation window opend ...");
                                Console.WriteLine("Searching next button...");
                                //btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                                //if (btnNext != null)
                                //{
                                    //AUIUtilities.ClickElement(btnNext);
                                    Wait(2);
                                    Console.WriteLine("---------- Installing Epia Resource File Editor window opend ...");
                                    Console.WriteLine("Searching Close button...");
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
                                        Console.WriteLine("Wait until Close button found...");
                                        appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                        aeBtnClose = appElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                                        Wait(5);
                                    }
                                    Console.WriteLine("Close button found... ---> Close Installer Window");
                                    AUIUtilities.ClickElement(aeBtnClose);
                                    Wait(3);
                               // }
                            }
                        }
                    }
                }
                #endregion
                sEventEnd = true;
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
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
                sEventEnd = false;
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
        static string uninstallWindowName = "Programs and Features";
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
                        Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                            TestTools.TfsUtilities.GetProjectName(ConstCommon.EPIA), sBuildNr);
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
                                    TestTools.TfsUtilities.GetProjectName(ConstCommon.EPIA),
                                    "GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
                            }*/
                            //---------
                            Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                        }
                        else
                        {
                            //if (sTotalFailed == 0)
                            //    TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                            //    TestTools.TfsUtilities.GetProjectName(ConstCommon.EPIA),
                            //    "GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
                            //else
                                TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName(ConstCommon.EPIA),
                                "GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
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

                            FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Failed" + testout, "Epia4+" + sCurrentPlatform + "Normal");
                                Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Failed:" + ConstCommon.EPIA, sOnlyUITest);
                            //}
                        }
                    }

                    if (sAutoTest)
                        TestTools.FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");
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
		#region OnLanguageSettingUIAEvent
		public static void OnLanguageSettingNLUIAEvent(object src, AutomationEventArgs args)
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
						Console.WriteLine("LanguageSettings failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						sErrorMessage = "LanguageSettings failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss");
						TestCheck = ConstCommon.TEST_FAIL;
						sEventEnd = true;
						return;
					}

                    ExpandCollapsePattern cP = aeCombo.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
                    cP.Expand();
                    Thread.Sleep(1000);

					bool select = false; //Utilities.SelectItemFromList("nl", aeCombo);
					if (TestCheck == ConstCommon.TEST_PASS)
					{
						//SelectionPattern selectPattern =
						//   aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

						AutomationElement item
							= AUIUtilities.FindElementByName("Nederlands", aeCombo);
						if (item != null)
						{
							Console.WriteLine("LanguageSettings item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
							Thread.Sleep(2000);
                            Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));
							//SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
							//itemPattern.Select();
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
							= AUIUtilities.FindElementByName("Windows gebruiker", aeCombo);
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
					ToggleState fstg = AUIUtilities.FindCheckBoxAndToggleState(ChkFullScreenID, element,ref sErrorMessage);
					if (  fstg == DEFAULT_FULLSCREEN )
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
        #region OnUninstallEtriccUIEvent
        public static void OnUninstallEpiaResourceFileEditorUIEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnUninstallEpiaResourceFileEditorUIEvent");
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
            string str = string.Format("OnUninstallEtriccUIEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
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
            else if (name.IndexOf("Programs and Features") >= 0)
            {
                AutomationElement rootElement = AutomationElement.RootElement;

                AutomationElement aeEpiaEditor = null;
                string sYesButtonName = "Yes";
                string sCloseButtonName = "Close";

                #region // Uninstall Etricc

                // (1) Programs and Features main window
                Console.WriteLine("Programs and Features Main Form found ... Welcom Main window");
                Console.WriteLine("Searching programs item element...");
                AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(element, Constants.PROGRAMS_FEATURES_FOLDER_VIEW_ID);
                if (aeGridView != null)
                    Console.WriteLine("Gridview found...");
                System.Threading.Thread.Sleep(2000);

                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.DataItem);

                AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children, c);
                Console.WriteLine("Programs count ..." + aeProgram.Count);
                for (int i = 0; i < aeProgram.Count; i++)
                {

                    Console.WriteLine("programs name: " + aeProgram[i].Current.Name);
                    if (aeProgram[i].Current.Name.StartsWith("Epia Resource"))
                        aeEpiaEditor = aeProgram[i];
                }

                if (aeEpiaEditor == null)    // Etricc Core not in Programs list
                {
                    Console.WriteLine("No Epia Editor name: ");
                    AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(element, sCloseButtonName);
                    if (btnClose != null)
                    {   // (2) Components
                        //UnInstalled = true;
                        AUIUtilities.ClickElement(btnClose);
                    }
                }
                else
                {
                    Console.WriteLine("Epia Editor name: " + aeEpiaEditor.Current.Name);
                    string x = aeEpiaEditor.Current.Name;

                    InvokePattern pattern = aeEpiaEditor.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    pattern.Invoke();
                    System.Threading.Thread.Sleep(2000);

                    System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName);
                    AutomationElement dialogElement = element.FindFirst(TreeScope.Children, condition);
                    if (dialogElement != null)
                    {
                        AutomationElement aeTitleBar =
                                AUIUtilities.FindElementByID("TitleBar", dialogElement);

                        Point pt1 = new Point(
                            (aeTitleBar.Current.BoundingRectangle.Left + aeTitleBar.Current.BoundingRectangle.Right) / 2,
                            (aeTitleBar.Current.BoundingRectangle.Bottom + aeTitleBar.Current.BoundingRectangle.Top) / 2);

                        Point newPt1 = new Point(pt1.X + 100, pt1.Y + 100);
                        Input.MoveTo(pt1);

                        System.Threading.Thread.Sleep(1000);
                        Input.SendMouseInput(pt1.X, pt1.Y, 0, SendMouseInputFlags.LeftDown | SendMouseInputFlags.Absolute);

                        System.Threading.Thread.Sleep(1000);
                        Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.Move | SendMouseInputFlags.Absolute);
                        //Input.MoveTo(newPt1);

                        System.Threading.Thread.Sleep(1000);
                        Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);

                        System.Threading.Thread.Sleep(3000);

                        AutomationElement btnYes = AUIUtilities.GetElementByNameProperty(element, sYesButtonName);
                        if (btnYes != null)
                        {   // (2) Components
                            AUIUtilities.ClickElement(btnYes);

                            AutomationElement installerElement
                                = AUIUtilities.GetElementByNameProperty(rootElement, "Windows Installer");

                            if (installerElement != null)
                                Console.WriteLine("Uninstaller dialog found ...");

                            //System.Windows.Forms.MessageBox.Show("TOP" + installerElement.Current.BoundingRectangle.Top, "Bottom" + installerElement.Current.BoundingRectangle.Bottom);
                            System.Threading.Thread.Sleep(3000);
                            Point pt = new Point(
                                (installerElement.Current.BoundingRectangle.Left + installerElement.Current.BoundingRectangle.Right) / 2,
                                ((installerElement.Current.BoundingRectangle.Top) / 2) - 60
                                );

                            //System.Windows.Forms.MessageBox.Show("TOP" + installerElement.Current.BoundingRectangle.Top, "Bottom" + installerElement.Current.BoundingRectangle.Bottom);
                            System.Threading.Thread.Sleep(1000);
                            Point newPt = new Point(400, 50);
                            //Input.MoveTo(pt);

                            Input.SendMouseInput(pt.X, pt.Y, 0, SendMouseInputFlags.Move | SendMouseInputFlags.Absolute);
                            System.Threading.Thread.Sleep(1000);
                            Input.SendMouseInput(pt.X, pt.Y, 0, SendMouseInputFlags.LeftDown | SendMouseInputFlags.Absolute);
                            System.Threading.Thread.Sleep(1000);
                            Input.SendMouseInput(newPt.X, newPt.Y, 0, SendMouseInputFlags.Move | SendMouseInputFlags.Absolute);
                            System.Threading.Thread.Sleep(1000);
                            Input.SendMouseInput(newPt.X, newPt.Y, 0, SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);


                            DateTime startTime = DateTime.Now;
                            TimeSpan mTime = DateTime.Now - startTime;
                            while (installerElement != null && mTime.TotalMilliseconds < 60000)
                            {
                                Console.WriteLine("----- Windows Installer window still open ...");
                                System.Threading.Thread.Sleep(5000);
                                mTime = DateTime.Now - startTime;

                                installerElement
                                   = AUIUtilities.GetElementByNameProperty(rootElement, "Windows Installer");
                            }

                            System.Threading.Thread.Sleep(3000);

                            AutomationElement aeRegisterdialog = null;
                            AutomationElement aeYesButton = null;
                            AutomationElement installer2Element
                                 = AUIUtilities.GetElementByNameProperty(rootElement, x);

                            startTime = DateTime.Now;
                            mTime = DateTime.Now - startTime;
                            while (installer2Element != null && mTime.TotalMilliseconds < 120000)
                            {
                                Console.WriteLine("----- Windows installer 2 window still open ...");
                                aeRegisterdialog =
                                   AUIUtilities.FindElementByID("FrmRemoveRegistryKeysDialog", AutomationElement.RootElement);

                                if (aeRegisterdialog != null)
                                {
                                    Console.WriteLine("aeRegisterdialog dialog found ...");
                                    aeYesButton =
                                        AUIUtilities.GetElementByNameProperty(rootElement, sYesButtonName);

                                    if (aeYesButton != null)
                                    {
                                        Console.WriteLine("Click Yes Button ...");
                                        AUIUtilities.ClickElement(aeYesButton);
                                        break;
                                    }
                                }
                                else
                                    Console.WriteLine("aeRegisterdialog dialog ----  NOT found ...");


                                System.Threading.Thread.Sleep(3000);
                                Console.WriteLine("Wait 3 sec and try to find --> Windows installer 2 window ...");
                                mTime = DateTime.Now - startTime;

                                installer2Element
                                   = AUIUtilities.GetElementByNameProperty(rootElement, x);

                                if (installer2Element == null)
                                    Console.WriteLine("uninstaller dialog closed.");

                            }

                            // wait until application uninstalled
                            startTime = DateTime.Now;
                            mTime = DateTime.Now - startTime;
                            bool hasApplication = EpiaUtilities.IsApplicationInstalled("EpiaResourceFileEditor", uninstallWindowName);
                            while (hasApplication == true && mTime.TotalMilliseconds < 120000)
                            {
                                System.Threading.Thread.Sleep(8000);
                                mTime = DateTime.Now - startTime;
                                if (mTime.TotalMilliseconds > 120000)
                                {
                                    System.Windows.Forms.MessageBox.Show("Uninstall EtriccCore run timeout " + mTime.TotalMilliseconds);
                                    break;
                                }
                                hasApplication = EpiaUtilities.IsApplicationInstalled("EpiaResourceFileEditor", uninstallWindowName);
                            }

                            System.Threading.Thread.Sleep(2000);
                            AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(element, sCloseButtonName);
                            if (btnClose != null)
                                AUIUtilities.ClickElement(btnClose);
                        }
                    }
                }
                #endregion
            }
            else
            {
                TestCheck = ConstCommon.TEST_FAIL;
                Console.WriteLine("Name is ------------:" + name);
                AUICommon.ErrorWindowHandling(element, ref sErrorMessage);
            }
            sEventEnd = true;
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

		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region Excel ------------------------------------------------------------------------------------------------
		public static void WriteResult(int result, int counter, string name,
			Excel.Worksheet sheet, string errorMSG)
		{
			string time = System.DateTime.Now.ToString("HH:mm:ss");
			xSheet.Cells[counter + 2 + 8, 1]= time;
			xSheet.Cells[counter + 2 + 8, 2]= name;
			xSheet.Cells[counter + 2 + 8, 3]= errorMSG;

            xRange = sheet.get_Range("A" + (Counter + 2 + 8), "A" + (Counter + 2 + 8));
            xRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

            xRange = sheet.get_Range("C" + (Counter + 2 + 8), "C" + (Counter + 2 + 8));
            xRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

            xRange = sheet.get_Range("B" + (Counter + 2 + 8), "B" + (Counter + 2 + 8));
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

		// ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
		#region TestLog ----------------------------------------------------------------------------------------------
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
			// sBuildDef = CI, Nightly... 
			// sTestApp = Epia for other Appliaction is layout
			// resultFile = xls file
			// sTotalFailed
			TestTools.Utilities.SendTestResultToDevelopers(resultFile, sTestApp, sBuildDef, logger, sTotalFailed,
                sBuildNr/*used for email title*/, str1/*content*/, sSendMail);
		}
		#endregion TestLog +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

		public static void CreateOutputFileInfo(string[] args, string currentPlatform, string PCName, ref string outPath, ref string outFilename)
		{
			
			// out filename 
			outFilename = System.DateTime.Now.ToString("yyyyMMdd-HH-mm-ss") + "-GUITESTS";

			if (args[0].IndexOf("Debug") > 0)
				outFilename = currentPlatform+"Debug-" + outFilename + "-" + PCName;
			else
				outFilename = currentPlatform+"Release-" + outFilename + "-" + PCName;

			if (args[10].ToLower().StartsWith("false"))
                outFilename = "Manual-" + currentPlatform+outFilename;

            outPath = args[1] + "\\TestResults";
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