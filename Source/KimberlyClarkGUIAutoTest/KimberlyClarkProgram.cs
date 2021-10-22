using System;
using System.Data.SqlClient;
using System.IO;
using System.ServiceProcess;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Xml;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using TestTools;
using Excel = Microsoft.Office.Interop.Excel;

namespace KimberlyClarkGUIAutoTest
{
	#pragma warning disable 0162 // Disable warning for Unreachable Code Detected.
	public class KimberlyClarkProgram
	{
        internal static TestTools.Logger logger;


		const string sUnitIDHeader = "Unit ID";
		const string sCarrierIDHeader = "Carrier ID";
		const string sCarrierTypeID = "Carrier type id";
		const string sProductIDHeader = "Product ID";
		const string sLocationIDHeader = "Location ID";

		// PCinfo
		static public string PCName;
		static public string OSName;
		static public string OSVersion;
		static public string UICulture;
		// Build param ========================================================
		static string sTFSServer = "http://teamApplication.teamSystems.egemin.be:8080";
		static bool TFSConnected = true;
        static IBuildServer m_BuildSvc;
		static string sInstallScriptsDir = string.Empty;
		static string sBuildBaseDir = string.Empty;
		static string sTestApp = string.Empty;
		static string sBuildDef = string.Empty;
		static string sBuildConfig = string.Empty;
		static string sBuildNr = string.Empty;
		static string sTestToolsVersion = string.Empty;
		static string m_SystemDrive = string.Empty;
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
		static bool sDemo = false;
		static bool sDropEwcsDB = false;
		static DateTime sStartTime = DateTime.Now;
		static TimeSpan sTime;
		private static int sNumAgvs = 2;
		// excel 	--------------------------------------------------------
		static Excel.Application xApp;
		static Excel.Workbook xBook;
		static Excel.Workbooks xBooks;
		static Excel.Range xRange;
		static Excel.Worksheet xSheet;

		#region TestCase Name
		//-------------------------------------------------------------------------EWCS
		private const string DATABASE_FILLER_CHECK = "DatabaseFillerCheck";
		private const string TESTDATA_ADD_CHECK = "TestDataAddCheck";
		
		private const string DISPLAY_CARRIER_TYPE = "CarrierTypeOverviewDisplay";
		private const string DISPLAY_REEL_QUALITY_REASONS = "ReelQualityReasonsOverviewDisplay";
		private const string DISPLAY_STORAGE_LOCATIONS = "StorageLocationsOverviewDisplay";
		private const string DISPLAY_PRODUCTS = "ProductsOverviewDisplay";
		private const string DISPLAY_REELS = "ReelsOverviewDisplay";
		private const string DISPLAY_CARRIERS = "CarriersOverviewDisplay";
		private const string DISPLAY_DELIVERY_REQUESTS = "DeliveryRequestOverviewDisplay";
		private const string DISPLAY_FORKLIFT_TRUCK_TRANSPORTS = "ForkliftTruckTransportsOverviewDisplay";
		private const string DISPLAY_TRANSPORTS = "TransportsOverviewDisplay";
		private const string DISPLAY_PRODUCTION_LINES = "ProductionLinesOverviewDisplay";
		private const string DISPLAY_FORKLIFT_TRUCKS = "ForkliftTrucksOverviewDisplay";

		private const string ADD_CARRIER_TYPES = "AddCarrierTypes";
		private const string ADD_REEL_QUALITY_REASONS = "AddReelQualityReasons";
		private const string ADD_PRODUCTS = "AddProducts";
		private const string ADD_REELS = "AddReels";
		private const string ADD_CARRIERS = "AddCarriers";

		private const string SEARCH_CARRIER_TYPE = "SearchCarrierTypes";
		private const string SEARCH_CARRIERS = "SearchCarriers";
		private const string SEARCH_PRODUCTS = "SearchProducts";
		private const string SEARCH_REELS = "SearchReels";
		private const string SEARCH_STORAGE_LOCATIONS = "SearchStorageLocations";

		private const string EDIT_REEL_QUALITY_REASONS = "EditReelQualityReasons";
		private const string EDIT_CARRIER_TYPE = "EditCarrierTypes";
		private const string DELETE_CARRIER_TYPE = "DeleteCarrierTypes";
		private const string ADD_10_UNITS_ON_1_CARRIER = "Add10UnitsOn1Carrier";

		private const string LOCATION_CONFIRM_MANUAL_CHANGE = "LocationConfirmManualChange";
		private const string CARRIER_CONFIRM_MANUAL_CHANGE = "CarrierConfirmManualChange";


		private const string EPIA3_CLOSE = "Epia3Close";
		//-------------------------------------------------------------------------EWCS
		private const string DISPLAY_SYSTEM_OVERVIEW = "SystemOverviewDisplay";
		private const string DISPLAY_AGV_OVERVIEW = "AgvOverviewDisplay";
		private const string DISPLAY_LOCATION_OVERVIEW = "LocationOverviewDisplay";
		private const string DISPLAY_TRANSPORT_OVERVIEW = "TransportOverviewDisplay";
		private const string LOCATION_OVERVIEW_OPEN_DETAIL = "LocationOverviewOpenDetail";
		private const string LOCATION_MODE_MANUAL = "LocationModeManual";
		private const string AGV_OVERVIEW_OPEN_DETAIL = "AgvOverviewOpenDetail";
		private const string AGV_JOB_OVERVIEW = "AgvJobsOverview";
		private const string AGV_JOB_OVERVIEW_OPEN_DETAIL = "AgvJobOverviewOpenDetail";
		private const string AGV_RESTART = "AgvRestart";
		private const string AGV_MODE_SEMIAUTOMATIC = "AgvModeSemiAutomatic";
		private const string CREATE_NEW_TRANSPORT = "TransportCreateNew";
		private const string EDIT_TRANSPORT = "TransportEdit";
		private const string CANCEL_TRANSPORT = "TransportCancel";
		private const string TRANSPORT_OVERVIEW_OPEN_DETAIL = "TransportOverviewOpenDetail";
		private const string AGV_OVERVIEW_REMOVE_ALL = "AgvsAllModeRemoved";
		private const string AGV_OVERVIEW_ID_SORTING = "AgvsIdSorting";
		#endregion TestCase Name
		//-------------------------------------------------------------------------EWCS
		private const string SYSTEM = "System";
		private const string INVENTORY = "Inventory";
		private const string CONFIGURATION = "Configuration";
		private const string LOGISTICS = "Logistics";
		private const string CARRIER_TYPES = "Carrier Types";
		private const string REEL_QUALITY_REASONS = "Reel Quality Reasons";       
		private const string STORAGE_LOCATIONS = "Storage Locations";
		private const string PRODUCTS = "Products";
		private const string REELS = "Reels";
		private const string CARRIERS = "Carriers";
		private const string DELIVERY_REQUESTS = "Delivery Requests";
		private const string FORKLIFT_TRUCK_TRANSPORTS = "Forklift Truck Transports";
		private const string TRANSPORTS = "Transports";
		private const string PRODUCTION_LINES = "Production Lines";
		private const string FORKLIFT_TRUCKS = "Forklift Trucks";
		private const string SYSTEM_OVERVIEW = "System Overview";
		private const string CARRIER_TYPES_OVERVIEW_TITLE = "Carrier Types";
		private const string REEL_QUALITY_REASONS_TITLE = "Reel Quality Reasons";
		private const string STORAGE_LOCATIONS_OVERVIEW_TITLE = "Storage Locations";
		private const string PRODUCTS_OVERVIEW_TITLE = "Products";
		private const string REELS_OVERVIEW_TITLE = "Reels";
		private const string CARRIERS_OVERVIEW_TITLE = "Carriers";
		private const string DELIVERY_REQUESTS_OVERVIEW_TITLE = "Delivery Requests";
		private const string FORKLIFT_TRUCK_TRANSPORTS_OVERVIEW_TITLE = "Forklift Truck Transports";
		private const string TRANSPORTS_OVERVIEW_TITLE = "Transports";
		private const string PRODUCTION_LINES_OVERVIEW_TITLE = "Production Lines";
		private const string FORKLIFT_TRUCKS_OVERVIEW_TITLE = "Forklift Trucks";
		private const string SYSTEM_OVERVIEW_TITLE = "System Overview";
		private const string KC_GRIDDATA_ID = "m_GridData";
		private const string INFRASTRUCTURE = "E'tricc";
		private const string AGV_OVERVIEW = "Agv overview";
		private const string LOCATION_OVERVIEW = "Location overview";
		private const string TRANSPORT_OVERVIEW = "Transport overview";
		private const string NEW_TRANSPORT = "New Transport";
		private const string AGV_OVERVIEW_TITLE = "Agvs";
		private const string LOCATION_OVERVIEW_TITLE = "Locations";
		private const string TRANSPORT_OVERVIEW_TITLE = "Transport orders";
		private const string DATAGRIDVIEW_ID = "m_GridData";
		private const string AGV_GRIDDATA_ID = "m_GridData";
		static private string sConnectionString = string.Empty;  

		[STAThread]
		static void Main(string[] args)
		{
			#region // Get test PC info======================================
			try  
			{
				HelpUtilities.SavePCInfo("y");
				HelpUtilities.GetPCInfo(out PCName, out OSName, out OSVersion, out UICulture, out TimeOnPC);
				Console.WriteLine("PCName : " + PCName);
				Console.WriteLine("OSName : " + OSName);
				Console.WriteLine("OSVersion : " + OSVersion);
				Console.WriteLine("UICulture : " + UICulture);
				Console.WriteLine("TimeOnPC : " + TimeOnPC);

				sConnectionString =
						 "Integrated Security=SSPI;"
						 + "Persist Security Info=False;"
						 + "Initial Catalog=Ewcs;"
						 + "Data Source=" + PCName;
			}
			catch (Exception ex)
			{
				MessageBox.Show("GetPCInfo:" + ex.Message);
			}
			#endregion

			if (!Constants.TEST)
			{
				#region // validate inputs
				try
				{
					if (args != null)
					{
						sInstallScriptsDir = args[0];
						Console.WriteLine("sInstallScriptsDir : " + sInstallScriptsDir);
						//MessageBox.Show(sInstallScriptsDir, "sInstallScriptsDir : ");
						sExcelVisible = args[1];
						
						if (args[2].StartsWith("true"))
							sAutoTest = true;
						else
							sAutoTest = false;

						sServerRunAs = args[3];

						if (args[4].StartsWith("true"))
							sDemo = true;
						else
							sDemo = false;

						Epia3Common.GetAllParameters(sInstallScriptsDir,
							ref sBuildBaseDir, ref sBuildNr, ref sTestApp, ref sBuildDef, ref sBuildConfig);

						sTestToolsVersion = args[6];
						sTFSServer = args[7];
						sBuildBaseDir = args[8];
						string ReportDirectory = sBuildBaseDir + "\\TestResults";
						if (!System.IO.Directory.Exists(ReportDirectory))
							System.IO.Directory.CreateDirectory(ReportDirectory);

						
						if (args[9].StartsWith("true"))
							sDropEwcsDB = true;
						else
							sDropEwcsDB = false;

						Console.WriteLine("sBuildBaseDir : " + sBuildBaseDir);
						Console.WriteLine("sBuildNr : " + sBuildNr);
						Console.WriteLine("sTestApp : " + sTestApp);
						Console.WriteLine("sBuildDef : " + sBuildDef);
						Console.WriteLine("sBuildConfig : " + sBuildConfig);

						Epia3Common.CreateOutputFileInfo(args, PCName, ref sOutFilePath, ref sOutFilename);

						sOutFilePath = Path.Combine(sBuildBaseDir, "TestResults");
						Console.WriteLine("sOutFilePath : " + sOutFilePath);
						Console.WriteLine("sOutFilename : " + sOutFilename);

						Epia3Common.CreateTestLog(ref slogFilePath, sOutFilePath, sOutFilename, ref Writer);

						Epia3Common.WriteTestLogMsg(slogFilePath, "InstallScrpts path: " + sInstallScriptsDir, Constants.TEST);
						Epia3Common.WriteTestLogMsg(slogFilePath, "Excel Visible: " + sExcelVisible, Constants.TEST);
						Epia3Common.WriteTestLogMsg(slogFilePath, "Auto test: " + sAutoTest, Constants.TEST);
						Epia3Common.WriteTestLogMsg(slogFilePath, "Server Run As: " + sServerRunAs, Constants.TEST);
						Epia3Common.WriteTestLogMsg(slogFilePath, "Demo test: " + sDemo, Constants.TEST);
						//MessageBox.Show("" + sOutFilePath, "sOutFilePath: ");

						string windir = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
						m_SystemDrive = Path.GetPathRoot(windir);
						Epia3Common.WriteTestLogMsg(slogFilePath, "m_SystemDrive: " + m_SystemDrive, Constants.TEST);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Validate arg: "+ex.Message+" : "+ex.StackTrace);
				}
				#endregion
                logger = new Logger(slogFilePath);
			}

			#region // Get TFS Server
			if (sAutoTest == true)
			{
				try
				{
                    string serverUrl = "http://team2010app.teamsystems.egemin.be:8080/tfs/Development";
                    Uri serverUri = new Uri(serverUrl);
                    System.Net.ICredentials tfsCredentials
                        = new System.Net.NetworkCredential("TfsBuild", "Egemin01", "TeamSystems.Egemin.Be");

                    TfsTeamProjectCollection tfsProjectCollection
                        = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                    //tfsProjectCollection.Connect(Microsoft.TeamFoundation.Framework.Common.ConnectOptions.IncludeServices);
                    m_BuildSvc = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));
				}
				catch (Exception ex)
				{
					MessageBox.Show("Get TFS Server:" + ex.Message);
					TFSConnected = false;
				}
			}
			else
				TFSConnected = false;
			#endregion

			Console.WriteLine("Test started:");
			sTestCaseName[0] = DATABASE_FILLER_CHECK;
			sTestCaseName[1] = TESTDATA_ADD_CHECK;
			sTestCaseName[2] = DISPLAY_CARRIER_TYPE;
			sTestCaseName[3] = DISPLAY_REEL_QUALITY_REASONS;
			sTestCaseName[4] = DISPLAY_STORAGE_LOCATIONS;
			sTestCaseName[5] = DISPLAY_PRODUCTS;
			sTestCaseName[6] = DISPLAY_REELS;
			sTestCaseName[7] = DISPLAY_CARRIERS;
			sTestCaseName[8] = DISPLAY_DELIVERY_REQUESTS;
			sTestCaseName[9] = DISPLAY_FORKLIFT_TRUCK_TRANSPORTS;
			sTestCaseName[10] = DISPLAY_TRANSPORTS;
			sTestCaseName[11] = DISPLAY_PRODUCTION_LINES;
			sTestCaseName[12] = DISPLAY_FORKLIFT_TRUCKS;
			sTestCaseName[13] = DISPLAY_SYSTEM_OVERVIEW;
			sTestCaseName[14] = ADD_CARRIER_TYPES;
			sTestCaseName[15] = ADD_REEL_QUALITY_REASONS;
			sTestCaseName[16] = ADD_CARRIERS;
			sTestCaseName[17] = ADD_PRODUCTS;
			sTestCaseName[18] = ADD_REELS;
			sTestCaseName[19] = ADD_10_UNITS_ON_1_CARRIER;
			sTestCaseName[20] = SEARCH_CARRIER_TYPE;
			sTestCaseName[21] = SEARCH_CARRIERS;
			sTestCaseName[22] = SEARCH_PRODUCTS;
			sTestCaseName[23] = SEARCH_REELS;
			sTestCaseName[24] = SEARCH_STORAGE_LOCATIONS;
			sTestCaseName[25] = EDIT_REEL_QUALITY_REASONS;
			sTestCaseName[26] = EDIT_CARRIER_TYPE;
			sTestCaseName[27] = DELETE_CARRIER_TYPE;
			sTestCaseName[28] = CARRIER_CONFIRM_MANUAL_CHANGE;
			sTestCaseName[29] = LOCATION_CONFIRM_MANUAL_CHANGE;                    
			//sTestCaseName[6] = DISPLAY_AGV_OVERVIEW;
			//sTestCaseName[2] = DISPLAY_LOCATION_OVERVIEW;
			//sTestCaseName[3] = DISPLAY_TRANSPORT_OVERVIEW;
			//sTestCaseName[4] = LOCATION_OVERVIEW_OPEN_DETAIL;
			//sTestCaseName[5] = LOCATION_MODE_MANUAL;
			//sTestCaseName[6] = AGV_OVERVIEW_OPEN_DETAIL;
			//sTestCaseName[7] = AGV_RESTART;
			//sTestCaseName[8] = AGV_MODE_SEMIAUTOMATIC;
			//sTestCaseName[9] = AGV_JOB_OVERVIEW;
			//sTestCaseName[10] = AGV_JOB_OVERVIEW_OPEN_DETAIL;
			//sTestCaseName[11] = CREATE_NEW_TRANSPORT;
			//sTestCaseName[12] = EDIT_TRANSPORT;
			//sTestCaseName[13] = CANCEL_TRANSPORT;
			//sTestCaseName[14] = TRANSPORT_OVERVIEW_OPEN_DETAIL;
			//sTestCaseName[15] = AGV_OVERVIEW_REMOVE_ALL;
			//sTestCaseName[16] = AGV_OVERVIEW_ID_SORTING;
			sTestCaseName[30] = EPIA3_CLOSE;
			//=============================================//

			try
			{
				if (!Constants.TEST)
				{
					Utilities.CloseProcess("EXCEL");
					Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
					Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_SERVER);
					Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
					Thread.Sleep(10000);
					#region // DatabaseDeployment First
					string sDeployPath = m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT + "\\DatabaseDeploy";
					if (sDropEwcsDB)
						TestTools.Utilities.StartProcessWaitForExit(sDeployPath, Constants.CREATEDB_BAT, string.Empty);
					#endregion
					Thread.Sleep(20000);

					Console.WriteLine("Start Database Filler : ");
					// Start database filler process
					TestTools.Utilities.StartProcessNoWait(m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT,
						"Egemin.Ewcs.Tools.DatabaseFiller.exe", "/c /t");
					Thread.Sleep(60000);
				   
					//========================   SERVER =================================================
					#region SERVER
					if (sServerRunAs.ToLower().IndexOf("service") >= 0)
					{
						// uninstall Egemin.Epia.server Service
						Console.WriteLine("UNINSTALL EPIA SERVER Service : ");
						TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
							ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /u");
						Thread.Sleep(2000);

						// uninstall Egemin.Ewcs.server Service
						Console.WriteLine("UNINSTALL EWCS SERVER Service : ");
						TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT,
							ConstCommon.EGEMIN_EWCS_SERVER_EXE, " /u");
						Thread.Sleep(2000);

						Console.WriteLine("INSTALL EPIA SERVER Service : ");
						TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
							ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /i");
						Thread.Sleep(2000);

						Console.WriteLine("INSTALL EWCS SERVER Service : ");
						TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT,
							ConstCommon.EGEMIN_EWCS_SERVER_EXE, " /i");
						Thread.Sleep(2000);

						Console.WriteLine("Start EPIA SERVER as Service : ");
						TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
							ConstCommon.EGEMIN_EPIA_SERVER_EXE, " /start");
						Thread.Sleep(2000);

						Console.WriteLine("Start EWCS SERVER as Service : ");
						TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT,
							ConstCommon.EGEMIN_EWCS_SERVER_EXE, " /start");
						Thread.Sleep(2000);

						ServiceController svcEpia = new ServiceController("Egemin Epia Server");
						Console.WriteLine(svcEpia.ServiceName + " has status " + svcEpia.Status.ToString());
						Thread.Sleep(2000);

						svcEpia.WaitForStatus(ServiceControllerStatus.Running);

						ServiceController svcEwcs = new ServiceController("Egemin Ewcs Service");
						Console.WriteLine(svcEwcs.ServiceName + " has status " + svcEwcs.Status.ToString());
						Thread.Sleep(2000);
						svcEwcs.WaitForStatus(ServiceControllerStatus.Running);

						Console.WriteLine("Epia and Ewcs SERVER Service Started : ");
						Thread.Sleep(2000);
					}

					if (sServerRunAs.ToLower().IndexOf("console") >= 0)
					{
						Console.WriteLine("Start EPIA and EWCS Server as console applications : ");
						// Start Epia SERVER as Console
						TestTools.Utilities.StartProcessNoWait(m_SystemDrive + ConstCommon.EPIA_SERVER_ROOT,
							ConstCommon.EGEMIN_EPIA_SERVER_EXE, string.Empty);

						// Start Ewcs SERVER as Console
						TestTools.Utilities.StartProcessNoWait(m_SystemDrive + ConstCommon.KIMBERLY_CLARK_SERVER_ROOT,
							ConstCommon.EGEMIN_EWCS_SERVER_EXE, string.Empty);

						Thread.Sleep(30000);
					}
					#endregion

					Thread.Sleep(5000);

					sStartTime = DateTime.Now;
					TimeSpan mTime = DateTime.Now - sStartTime;
					//========================   SHELL =================================================
					#region  Shell
					Console.WriteLine("Start Shell : ");

					AutomationEventHandler UIAShellEventHandler = new AutomationEventHandler(OnUIAShellEvent);
					// Add Open window Event Handler
					Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
						AutomationElement.RootElement, TreeScope.Descendants, UIAShellEventHandler);
					sEventEnd = false;
					TestCheck = ConstCommon.TEST_PASS;

					Thread.Sleep(20000);

					// Start Shell
					TestTools.Utilities.StartProcessNoWait(m_SystemDrive+ConstCommon.EPIA_CLIENT_ROOT,
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

					#endregion

					Console.WriteLine("Shell started after (sec) : " + mTime.Seconds);

					#region // Wait until Egemin Shell is opened
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
						Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
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
					#endregion
				}

				#region // Excel Header
				// Excel file not for EpiaTestPC3
				if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
						|| PCName.ToUpper().Equals("EPIATESTSRV3V1"))
					Epia3Common.WriteTestLogMsg(slogFilePath, "No Excel due to: " + PCName, Constants.TEST);
				else
				{
					xApp = new Excel.Application();
					xBooks = xApp.Workbooks;
					xBook = xBooks.Add(Type.Missing);
					xSheet = (Excel.Worksheet)xBook.Worksheets[1];
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
					xSheet.Cells[1, 2] = "EWcs Kimberly Clark UI Test Scenarios";

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
				}
				#endregion

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
				int aantal = 31;

				if (sDemo)
					aantal = 30;

				if (Constants.TEST)
					aantal = 31;
				else
				{
					#region update build quality to GUI Tests Started
                    if (TFSConnected)
                    {
                        Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                            TestTools.TfsUtilities.GetProjectName(ConstCommon.EWCS_PROJECTS), sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                        if (quality.Equals("GUI Tests Failed"))
                        {
                            Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has build quality: " + quality + " , no update needed", false);
                        }
                        else
                        {
                            TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName(ConstCommon.EWCS_PROJECTS),
                                "GUI Tests Started", m_BuildSvc, sDemo ? "true" : "false");
                        }

                        if (sAutoTest)
                        {
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(sBuildBaseDir, "GUI Tests Started", ConstCommon.EWCS_PROJECTS);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.EWCS_PROJECTS, false);
                        }
                    }
					#endregion
				}

				while (Counter < aantal)
				{
					sResult = ConstCommon.TEST_UNDEFINED;
					switch (sTestCaseName[Counter])
					{
						case DATABASE_FILLER_CHECK:
							DatabaseFillerCheck(DATABASE_FILLER_CHECK, aeForm, out sResult);
							break;
						case TESTDATA_ADD_CHECK:
							TestDataAddCheck(TESTDATA_ADD_CHECK, aeForm, out sResult);
							break;
						case DISPLAY_CARRIER_TYPE:
							DisplayCarrierType(DISPLAY_CARRIER_TYPE, aeForm, out sResult);
							break;
						case DISPLAY_REEL_QUALITY_REASONS:
							DisplayReelQualityReasons(DISPLAY_REEL_QUALITY_REASONS, aeForm, out sResult);
							break;
						case DISPLAY_STORAGE_LOCATIONS:
							DisplayStorageLocations(DISPLAY_STORAGE_LOCATIONS, aeForm, out sResult);
							break;
						case DISPLAY_PRODUCTS:
							DisplayProducts(DISPLAY_PRODUCTS, aeForm, out sResult);
							break;
						case DISPLAY_REELS:
							DisplayReels(DISPLAY_REELS, aeForm, out sResult);
							break;
						case DISPLAY_CARRIERS:
							DisplayCarriers(DISPLAY_CARRIERS, aeForm, out sResult);
							break;
						case DISPLAY_DELIVERY_REQUESTS:
							DisplayDeliveryRequests(DISPLAY_DELIVERY_REQUESTS, aeForm, out sResult);
							break;
						case DISPLAY_FORKLIFT_TRUCK_TRANSPORTS:
							DisplayForkliftTruckTransports(DISPLAY_FORKLIFT_TRUCK_TRANSPORTS, aeForm, out sResult);
							break;
						case DISPLAY_TRANSPORTS:
							DisplayTransports(DISPLAY_TRANSPORTS, aeForm, out sResult);
							break;
						case DISPLAY_PRODUCTION_LINES:
							DisplayProductionLines(DISPLAY_PRODUCTION_LINES, aeForm, out sResult);
							break;
						case DISPLAY_FORKLIFT_TRUCKS:
							DisplayForkliftTrucks(DISPLAY_FORKLIFT_TRUCKS, aeForm, out sResult);
							break;
						case DISPLAY_SYSTEM_OVERVIEW:
							SystemOverviewDisplay(DISPLAY_SYSTEM_OVERVIEW, aeForm, out sResult);
							break;
						case ADD_CARRIER_TYPES:
							AddCarrierType(ADD_CARRIER_TYPES, aeForm, out sResult);
							break;
						case ADD_REEL_QUALITY_REASONS:
							AddReelQualityReasons(ADD_REEL_QUALITY_REASONS, aeForm, out sResult);
							break;
						case ADD_PRODUCTS:
							AddProduct(ADD_PRODUCTS, aeForm, out sResult);
							break;
						case ADD_REELS:
							AddReel(ADD_REELS, aeForm, out sResult);
							break;
						case ADD_CARRIERS:
							AddCarrier(ADD_CARRIERS, aeForm, out sResult);
							break;
						case SEARCH_CARRIER_TYPE:
							SearchCarrierTypes(SEARCH_CARRIER_TYPE, aeForm, out sResult);
							break;
						case SEARCH_CARRIERS:
							SearchCarriers(SEARCH_CARRIERS, aeForm, out sResult);
							break;
						case SEARCH_PRODUCTS:
							SearchProducts(SEARCH_PRODUCTS, aeForm, out sResult);
							break;
						case SEARCH_REELS:
							SearchReels(SEARCH_REELS, aeForm, out sResult);
							break;
						case SEARCH_STORAGE_LOCATIONS:
							SearchStorageLocations(SEARCH_STORAGE_LOCATIONS, aeForm, out sResult);
							break;
						case EDIT_REEL_QUALITY_REASONS:
							EditReelQualityReasons(EDIT_REEL_QUALITY_REASONS, aeForm, out sResult);
							break;
						case EDIT_CARRIER_TYPE:
							EditCarrierTypes(EDIT_CARRIER_TYPE, aeForm, out sResult);
							break;
						case DELETE_CARRIER_TYPE:
							DeleteCarrierTypes(DELETE_CARRIER_TYPE, aeForm, out sResult);
							break;
						case ADD_10_UNITS_ON_1_CARRIER:
							Add10UnitsOn1Carrier(ADD_10_UNITS_ON_1_CARRIER, aeForm, out sResult);
							break;
						case LOCATION_CONFIRM_MANUAL_CHANGE:
							LocationConfirmManualChange(LOCATION_CONFIRM_MANUAL_CHANGE, aeForm, out sResult);
							break;
						case CARRIER_CONFIRM_MANUAL_CHANGE:
							CarrierConfirmManualChange(CARRIER_CONFIRM_MANUAL_CHANGE, aeForm, out sResult);
							break;
						//---------------------------------------------------------------------------------------
					  
						case DISPLAY_AGV_OVERVIEW:
							AgvOverviewDisplay(DISPLAY_AGV_OVERVIEW, aeForm, out sResult);
							break;
						case DISPLAY_LOCATION_OVERVIEW:
							LocationOverviewDisplay(DISPLAY_LOCATION_OVERVIEW, aeForm, out sResult);
							break;
						case DISPLAY_TRANSPORT_OVERVIEW:
							TransportOverviewDisplay(DISPLAY_TRANSPORT_OVERVIEW, aeForm, out sResult);
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
						case AGV_MODE_SEMIAUTOMATIC:
							AgvModeSemiAutomatic(AGV_MODE_SEMIAUTOMATIC, aeForm, out sResult);
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
						case CANCEL_TRANSPORT:
							CancelTransport(CANCEL_TRANSPORT, aeForm, out sResult);
							break;
						case TRANSPORT_OVERVIEW_OPEN_DETAIL:
							TransportOverviewOpenDetail(TRANSPORT_OVERVIEW_OPEN_DETAIL, aeForm, out sResult);
							break;
						case EPIA3_CLOSE:
							Epia3Close(EPIA3_CLOSE, aeForm, out sResult);
							break;
						default:
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

				if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
					|| PCName.ToUpper().Equals("EPIATESTSRV3V1"))
					Epia3Common.WriteTestLogMsg(slogFilePath, "No Excel due to: " + PCName, Constants.TEST);
				else
				{
					xSheet.Cells.set_Item(Counter + 2 + 8, 1, "Total tests: ");
					xSheet.Cells.set_Item(Counter + 3 + 8, 1, "Total Passes: ");
					xSheet.Cells.set_Item(Counter + 4 + 8, 1, "Total Failed: ");

					xSheet.Cells.set_Item(Counter + 2 + 8, 2, sTotalCounter);
					xSheet.Cells.set_Item(Counter + 3 + 8, 2, sTotalPassed);
					xSheet.Cells.set_Item(Counter + 4 + 8, 2, sTotalFailed);

					//xSheet.Cells.set_Item(Counter + 5 + 8, 2, "Project is: " + sProjectFile);

					// Add Legende
					xSheet.Cells.set_Item(Counter + 6 + 8, 2, "Legende");
					xRange = xSheet.get_Range("B" + (Counter + 6 + 8), "B" + (Counter + 6 + 7));
					xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

					xSheet.Cells.set_Item(Counter + 7 + 8, 2, "Pass");
					xRange = xSheet.get_Range("B" + (Counter + 7 + 8), "B" + (Counter + 7 + 7));
					xRange.Interior.ColorIndex = ConstCommon.EXCEL_GREEN;
					xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

					xSheet.Cells.set_Item(Counter + 8 + 8, 2, "Fail");
					xRange = xSheet.get_Range("B" + (Counter + 8 + 8), "B" + (Counter + 8 + 7));
					xRange.Interior.ColorIndex = ConstCommon.EXCEL_RED;
					xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

					xSheet.Cells.set_Item(Counter + 9 + 7, 2, "Exception");
					xRange = xSheet.get_Range("B" + (Counter + 9 + 7), "B" + (Counter + 9 + 7));
					xRange.Interior.ColorIndex = ConstCommon.EXCEL_PINK;
					xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

					xSheet.Cells.set_Item(Counter + 10 + 7, 2, "Untested");
					xRange = xSheet.get_Range("B" + (Counter + 10 + 7), "B" + (Counter + 10 + 7));
					xRange.Interior.ColorIndex = ConstCommon.EXCEL_YELLOW;
					xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
					xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;
				}

				if (!Constants.TEST)
				{
					if (TFSConnected)
					{
                        Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                            TestTools.TfsUtilities.GetProjectName(ConstCommon.EWCS_PROJECTS), sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

						if (quality.Equals("GUI Tests Failed"))
						{
							string msg = sBuildNr + " has failed quality, no update needed :" + quality;
							Epia3Common.WriteTestLogMsg(slogFilePath, msg, Constants.TEST);
						}
						else
						{
                            if (sTotalFailed == 0)
                                TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName(ConstCommon.EWCS_PROJECTS),
                                "GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
                            else
                                TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName(ConstCommon.EWCS_PROJECTS),
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
                                FileManipulation.UpdateStatusInTestInfoFile(sBuildBaseDir, "GUI Tests Passed" + testout, ConstCommon.KC);
							else
                                FileManipulation.UpdateStatusInTestInfoFile(sBuildBaseDir, "GUI Tests Failed" + testout, ConstCommon.KC);
						}
					}

					if (sAutoTest)
						FileManipulation.UpdateTestWorkingFile(sBuildBaseDir, "false");
				}

				if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
					|| PCName.ToUpper().Equals("EPIATESTSRV3V1"))
					Epia3Common.WriteTestLogMsg(slogFilePath, "No Excel due to: " + PCName, Constants.TEST);
				else
					xSheet.Columns.AutoFit();

				// save Excel to Local machine
				string sXLSPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
					sOutFilename + ".xls");
				// Save the Workbook locally  --- not for PC EPIATESTPC3
				object missing = System.Reflection.Missing.Value;
				if (PCName.ToUpper().Equals("EPIATESTPC3") || PCName.ToUpper().Equals("EPIATESTSERVER3")
					|| PCName.ToUpper().Equals("EPIATESTSRV3V1") || Constants.TEST)
				{
					Console.WriteLine("No Excel due to: " + PCName);
					Console.WriteLine("sOutFilename: " + sOutFilename);
				}
				else
				{
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
					SendEmail(sXLSPath);
				}

			}
			catch (Exception ex)
			{
				Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
				Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace, Constants.TEST);
				MessageBox.Show("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
				try
				{
					if (sAutoTest)
					{
                        FileManipulation.UpdateStatusInTestInfoFile(sBuildBaseDir, "GUI Tests Exception -->" + sOutFilename + ".log", ConstCommon.KC);
						Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace, Constants.TEST);

						Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
						Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
						Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_SERVER);
						Utilities.CloseProcess("cmd");
						FileManipulation.UpdateTestWorkingFile(sBuildBaseDir, "false");

						if (TFSConnected)
						{
                            Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                                    TestTools.TfsUtilities.GetProjectName(ConstCommon.EWCS_PROJECTS), sBuildNr);
                            string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;
                            
							if (quality.Equals("GUI Tests Failed"))
							{
								Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, Constants.TEST);
							}
							else
							{
                                TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName(ConstCommon.EWCS_PROJECTS),
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
				// Close LogFile
				Epia3Common.CloseTestLog(slogFilePath, Constants.TEST);

				Console.WriteLine("\nClosing application in 10 seconds");
				if (Constants.TEST)
					Thread.Sleep(10000000);
				else
					Thread.Sleep(10000);

				// close CommandHost
				Thread.Sleep(10000);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_SERVER);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
				Utilities.CloseProcess(ConstCommon.EGEMIN_EWCS_TOOLS_DATABASE_FILLER);
				
				Console.WriteLine("\nEnd test run\n");
			}
			catch (Exception ex)
			{
				Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
				if (sAutoTest)
				{

                    FileManipulation.UpdateStatusInTestInfoFile(sBuildBaseDir, "FUNCTIONAL Tests Exception -->" + sOutFilename + ".log", ConstCommon.KC);
					Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace, Constants.TEST);

					Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
					Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
					Utilities.CloseProcess("cmd");
					FileManipulation.UpdateTestWorkingFile(sBuildBaseDir, "false");

					if (TFSConnected)
					{
                        Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                                   TestTools.TfsUtilities.GetProjectName(ConstCommon.EWCS_PROJECTS), sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;
                        

						if (quality.Equals("Functional Tests Failed"))
						{
							Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, Constants.TEST);
						}
						else
						{
                            TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                       TestTools.TfsUtilities.GetProjectName(ConstCommon.EWCS_PROJECTS),
                                       "Functional Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
						}
					}
				}
			}
		}

		private static void XmlServerConfigUpdate(string WorkerPathValue)
		{
			var xDoc = new XmlDocument();
			xDoc.Load("C:\\Etricc\\Server\\Egemin.Etricc.Server.exe.config");

			var xPathNav = xDoc.CreateNavigator();

			xPathNav.MoveToFirstChild();
			xPathNav.MoveToNext();
			xPathNav.MoveToFirstChild();
			xPathNav.MoveToNext();
			xPathNav.MoveToFirstChild();
			xPathNav.MoveToFirstChild();
			xPathNav.MoveToFirstChild();
			xPathNav.MoveToFirstChild();   //  parameter
			xPathNav.DeleteSelf();
			xPathNav.AppendChild("<parameter name=\"XmlFile\" value=" + WorkerPathValue + " />");
			//MessageBox.Show("node name is : " + xPathNav.Name);
			xDoc.Save("C:\\Etricc\\Server\\Egemin.Etricc.Server.exe.config");

			return;
		}

		#region Excel ------------------------------------------------------------------------------------------------
		public static void WriteResult(int result, int counter, string name,
			Excel.Worksheet sheet, string errorMSG)
		{
			string time = System.DateTime.Now.ToString("HH:mm:ss");
			xSheet.Cells.set_Item(counter + 2 + 8, 1, time);
			sheet.Cells.set_Item(counter + 2 + 8, 2, name);
			sheet.Cells.set_Item(counter + 2 + 8, 3, errorMSG);
			xRange = sheet.get_Range("B" + (Counter + 2 + 8), "B" + (Counter + 2 + 8));
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
		
		
		#region DatabaseFillerCheck
		public static void DatabaseFillerCheck(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;
		   
			try
			{
				SqlConnection myConnection = new SqlConnection(sConnectionString);

				try
				{
					myConnection.Open();
				}
				catch (Exception e)
				{
					string exc = e.Message + System.Environment.NewLine + e.StackTrace;
					Console.WriteLine(e.ToString());
					MessageBox.Show(e.Message + System.Environment.NewLine + e.StackTrace);
					Epia3Common.WriteTestLogPass(slogFilePath, exc, Constants.TEST);
					return;
				}

				string sqlCommand = "select COUNT(*) from [Ewcs].[dbo].[Locations]";
				SqlCommand myCommand = new SqlCommand(sqlCommand, myConnection);
				int count = (int)myCommand.ExecuteScalar();

				int nrlocations = 854;
				string mMsg = "sql command:" + sqlCommand + System.Environment.NewLine;
				if (count == nrlocations )
				{
					result = ConstCommon.TEST_PASS;
					Console.WriteLine(mMsg+" LocationID Count = "+count);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					sErrorMessage = mMsg + " LocationID Count is not " + nrlocations + "  but is:" + count;
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}

				Thread.Sleep(2000);
				myConnection.Close();
			}
			catch (Exception ex)
			{
				sErrorMessage = sErrorMessage + " CONN:" + sConnectionString;
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion

		#region TestDataAddCheck
		public static void TestDataAddCheck(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			TestTools.Utilities.StartProcessWaitForExit(@"C:\KC", "DBTestData.bat", string.Empty);

			Thread.Sleep(5000);

			try
			{
				SqlConnection myConnection = new SqlConnection(sConnectionString);

				try
				{
					myConnection.Open();
				}
				catch (Exception e)
				{
					string exc = e.Message + System.Environment.NewLine + e.StackTrace;
					Console.WriteLine(e.ToString());
					MessageBox.Show(e.Message + System.Environment.NewLine + e.StackTrace);
					Epia3Common.WriteTestLogPass(slogFilePath, exc, Constants.TEST);
					return;
				}

				string sqlCommand = "select COUNT(*) from [Ewcs].[dbo].[ForkliftTruckTransports]";
				SqlCommand myCommand = new SqlCommand(sqlCommand, myConnection);
				int count = (int)myCommand.ExecuteScalar();

				string mMsg = "sql command:" + sqlCommand + System.Environment.NewLine;
				if (count == 0)
				{
					result = ConstCommon.TEST_PASS;
					Console.WriteLine(mMsg + " ForkliftTruckTransportID Count = " + count);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					sErrorMessage = mMsg + " ForkliftTruckTransportID Count is not 16 but is:" + count;
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}

				Thread.Sleep(2000);
				myConnection.Close();
			}
			catch (Exception ex)
			{
				sErrorMessage = sErrorMessage + " CONN:" + sConnectionString;
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplayCarrierType
		public static void DisplayCarrierType(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 12);
			  
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, CONFIGURATION, CARRIER_TYPES, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Carrier Types Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIER_TYPES_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the CarrierTypesOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = CARRIER_TYPES + " Window not found";
					Console.WriteLine("FindElementByID failed:" + CARRIER_TYPES);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = CARRIER_TYPES + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplayReelQualityReasons
		public static void DisplayReelQualityReasons(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, CONFIGURATION, REEL_QUALITY_REASONS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Reel Quality Reasons Types Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, REEL_QUALITY_REASONS_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the Reel Quality Reasons Overview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = REEL_QUALITY_REASONS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + REEL_QUALITY_REASONS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = REEL_QUALITY_REASONS + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion

		#region DisplayStorageLocations
		public static void DisplayStorageLocations(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, STORAGE_LOCATIONS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Storage Locations Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, STORAGE_LOCATIONS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the Storage Locations Overview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = STORAGE_LOCATIONS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + STORAGE_LOCATIONS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = STORAGE_LOCATIONS + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplayProducts
		public static void DisplayProducts(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root,2);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, PRODUCTS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Products Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, PRODUCTS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ProductsOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = PRODUCTS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + PRODUCTS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = PRODUCTS + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplayReels
		public static void DisplayReels(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, REELS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Products Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, REELS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ReelsOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = REELS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + REELS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = CARRIER_TYPES + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplayCarriers
		public static void DisplayCarriers(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, CARRIERS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Carriers Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIERS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the CARRIERS Overview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = CARRIERS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + CARRIERS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = CARRIER_TYPES + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplayDeliveryRequests
		public static void DisplayDeliveryRequests(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, LOGISTICS, DELIVERY_REQUESTS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);
				
				// Find Delivery Requests Overview Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, DELIVERY_REQUESTS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the Delivery Requests Overview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = DELIVERY_REQUESTS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + DELIVERY_REQUESTS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = DELIVERY_REQUESTS + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplayForkliftTruckTransports
		public static void DisplayForkliftTruckTransports(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, LOGISTICS, FORKLIFT_TRUCK_TRANSPORTS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Forklift Truck Transports Overview Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, FORKLIFT_TRUCK_TRANSPORTS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the Forklift Truck Transports Overview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = FORKLIFT_TRUCK_TRANSPORTS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + FORKLIFT_TRUCK_TRANSPORTS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = FORKLIFT_TRUCK_TRANSPORTS + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplayTransports
		public static void DisplayTransports(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, LOGISTICS, TRANSPORTS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Transports Overview Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, TRANSPORTS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the Transports Overview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = TRANSPORTS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + TRANSPORTS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = TRANSPORTS + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplayProductionLines
		public static void DisplayProductionLines(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, SYSTEM, PRODUCTION_LINES, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Production Lines Overview Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, PRODUCTION_LINES_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the Production Lines Overview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = PRODUCTION_LINES + " Window not found";
					Console.WriteLine("FindElementByID failed:" + PRODUCTION_LINES);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = PRODUCTION_LINES + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplayForkliftTrucks
		public static void DisplayForkliftTrucks(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, SYSTEM, FORKLIFT_TRUCKS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Production Lines Overview Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, FORKLIFT_TRUCKS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the Production Lines Overview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = FORKLIFT_TRUCKS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + FORKLIFT_TRUCKS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = FORKLIFT_TRUCKS + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region DisplaySystemOverview
		public static void DisplaySystemOverview(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, SYSTEM, SYSTEM_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Production Lines Overview Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, SYSTEM_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the Production Lines Overview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = SYSTEM_OVERVIEW + " Window not found";
					Console.WriteLine("FindElementByID failed:" + SYSTEM_OVERVIEW);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = SYSTEM_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region AddCarrierType
		public static void AddCarrierType(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			string PALLET_TYPE = "PalletType";
			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);

				#region// Open and Find Carrier Type Screen
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, CONFIGURATION, CARRIER_TYPES, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Carrier Types Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIER_TYPES_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );
				
				// Find the CarrierTypesOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = CARRIER_TYPES + " Window not found";
					Console.WriteLine("FindElementByID failed:" + CARRIER_TYPES);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				#endregion

				#region  Add CarrierTypes
				
				// Find tool bar 
				AutomationElement aeToolBar = AUIUtilities.FindElementByType(ControlType.ToolBar, aeOverview);
				if (aeToolBar == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "ToolBar " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("ToolBar " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				// Find Edit textbox 
				AutomationElement aeEdit = AUIUtilities.FindElementByType(ControlType.Edit, aeToolBar);
				if (aeEdit == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "EditTextBox " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("EditTextBox " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Double Bottom = aeEdit.Current.BoundingRectangle.Bottom;
				Double Left = aeEdit.Current.BoundingRectangle.Left;
				Double Right = aeEdit.Current.BoundingRectangle.Right;
				Double Width = aeEdit.Current.BoundingRectangle.Width;
				Double Height = aeEdit.Current.BoundingRectangle.Height;
				Double X = aeEdit.Current.BoundingRectangle.X;
				Double Y = aeEdit.Current.BoundingRectangle.Y;
				Double Top = aeEdit.Current.BoundingRectangle.Top;
				
				double xclick = Right+45;
				Point NewBtnPoint = new Point(xclick, (Bottom + Top) / 2);
				
				Input.MoveTo(NewBtnPoint);
				Thread.Sleep(2000);

				for (int i = 1; i < 5; i++)
				{
					Thread.Sleep(3000);

					Input.MoveToAndDoubleClick(NewBtnPoint);
					Thread.Sleep(2000);
					Input.MoveToAndDoubleClick(NewBtnPoint);
					Thread.Sleep(2000);

					 //find new Egemin Shell Window screen
					System.Windows.Automation.Condition c2x = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					// Find the New CarrierType Screen element.
					AutomationElement aeNewCT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2x);
					if (aeNewCT == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "New Carrier Type Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}

					string origCTId = "";
					// "m_TxtCarrierTypeId"
					if (AUIUtilities.FindTextBoxAndChangeValue("m_TxtCarrierTypeId", aeNewCT, out origCTId, (i==4)? "RackType" :PALLET_TYPE+i, ref sErrorMessage))
							Thread.Sleep(3000);
					else
					{
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "m_TxtCarrierTypeId");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "m_TxtCarrierTypeId";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				
					// "carrierIdMaskTextBox"
					if (AUIUtilities.FindTextBoxAndChangeValue("carrierIdMaskTextBox", aeNewCT, out origCTId, (i == 4) ? "R0000" : "P0000", ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "carrierIdMaskTextBox");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "carrierIdMaskTextBox";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
					
					// "widthTextBox"
					if (AUIUtilities.FindTextBoxAndChangeValue("widthTextBox", aeNewCT, out origCTId, "1800", ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "widthTextBox");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "widthTextBox";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				
					// "depthTextBox"
					if (AUIUtilities.FindTextBoxAndChangeValue("depthTextBox", aeNewCT, out origCTId, "1800", ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "depthTextBox");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "depthTextBox";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
					
					// "heightTextBox"
					if (AUIUtilities.FindTextBoxAndChangeValue("heightTextBox", aeNewCT, out origCTId, ""+i*100, ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "heightTextBox");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "heightTextBox";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				
					string SaveID = "m_btnSave";
					// Find Save Button  element. 
					AutomationElement aeSaveBtn = AUIUtilities.FindElementByID(SaveID, aeNewCT);
					if (aeSaveBtn == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Save Button aeSaveBtn not found";
						Console.WriteLine("FindElementByID failed:" + SaveID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeSaveBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
					}

					Thread.Sleep(4000);
				}
				#endregion

				#region  // Validation results

				AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find CarrierTypeDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("CarrierTypeDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

				for (int i = 1; i < 5; i++)   // row 0 is carrier type virtual
				{
					
					Thread.Sleep(3000);
					string header = sCarrierTypeID;
					string cellValue = "PalletType"+i;
					int cellRow = -1;
					int startRow = 3;
					//AutomationElement aeCell = FindCellFromGrid(aeGrid, header, ( i==4 )? "RackType":cellValue, ref cellRow); 
					AutomationElement aeCell = FindCellFromGridStartedAtRow(aeGrid,
						header, (i == 4) ? "RackType" : cellValue, startRow + i, ref cellRow); 
				  
				   
					if (aeCell == null)
					{
						sErrorMessage = "Find CarrierTypeDataGridView aeCell failed:" + cellValue;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell CarrierTypeDataGridView found: " + cellValue);
					}
				}

			   
				result = ConstCommon.TEST_PASS;
				Console.WriteLine(testname + " ---pass --- ");
				Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
			   
				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region AddReelQualityReasons
		public static void AddReelQualityReasons(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 5);
				#region // Find and open overview screen
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, CONFIGURATION, REEL_QUALITY_REASONS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Reel Quality Reasons Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, REEL_QUALITY_REASONS_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ReelQualityReasonsOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = REEL_QUALITY_REASONS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + REEL_QUALITY_REASONS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				#endregion

				#region  Add Reel Quality Reasons
				Thread.Sleep(2000);

				// Find New... Button
				AutomationElement aeToolBar = AUIUtilities.FindElementByID("m_ToolStripInfo", aeOverview);
				if (aeToolBar == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "ToolBar " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("ToolBar " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				// Move to New.. button Point
				Double Bottom = aeToolBar.Current.BoundingRectangle.Bottom;
				Double Left = aeToolBar.Current.BoundingRectangle.Left;
				Double Right = aeToolBar.Current.BoundingRectangle.Right;
				Double Width = aeToolBar.Current.BoundingRectangle.Width;
				Double Height = aeToolBar.Current.BoundingRectangle.Height;
				Double X = aeToolBar.Current.BoundingRectangle.X;
				Double Y = aeToolBar.Current.BoundingRectangle.Y;
				Double Top = aeToolBar.Current.BoundingRectangle.Top;
				Point NewBtnPoint = new Point(Left + 8, (Bottom + Top) / 2);
				Input.MoveTo(NewBtnPoint);

				string DescriptionId = "descriptionTextBox";
				string CheckBoxBlueStarId = "validForBlueStarCheckBox";
				string CheckBoxGoldStarId = "validForGoldStarCheckBox";
				string CheckBoxProdStateId = "validForProductionStateCheckBox";
				string CheckBoxQualityStateId = "validForQualityStateCheckBox";

				for (int i = 1; i < 6; i++)
				{
					Thread.Sleep(3000);
					Input.MoveToAndClick(NewBtnPoint);
					Thread.Sleep(2000);
					Input.ClickAtPoint(NewBtnPoint);
					//Console.WriteLine("button clicked ----------------:" + ":::::::::::");
					Thread.Sleep(3000);

					//find new Egemin Shell Window screen
					System.Windows.Automation.Condition c2x = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					// Find the New Reel Quality Reasons Screen element.
					AutomationElement aeNewCT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2x);
					if (aeNewCT == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "New Reel Quality Reasons Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}

					string origCTId = "";
					// Description
					if (AUIUtilities.FindTextBoxAndChangeValue(DescriptionId, aeNewCT, out origCTId, "ReelQualityReason"+i, ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + DescriptionId);
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + DescriptionId;
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}

					if (i == 1)    // CheckBoxBlueStarId   checked
					{
						#region // blue star checkbox
						try
						{
							bool check = AUIUtilities.FindElementAndToggle(CheckBoxBlueStarId, aeNewCT, ToggleState.On);
							if (check)
								Thread.Sleep(3000);
							else
							{
								Console.WriteLine("FindElementAndToggle failed:" + CheckBoxBlueStarId);
								Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxBlueStarId, Constants.TEST);
								TestCheck = ConstCommon.TEST_FAIL;
								sErrorMessage = "FindElementAndToggle failed:" + CheckBoxBlueStarId;
								return;
							}
						}
						catch (Exception ex)
						{
							TestCheck = ConstCommon.TEST_FAIL;
							Console.WriteLine("Blue Star CheckBox :" + ex.Message + " --- " + ex.StackTrace);
							sErrorMessage = ex.Message + " --- " + ex.StackTrace;
						}
						#endregion
					}

					if (i == 2)    // CheckBoxGoldStarId   checked
					{
						#region // gold star checkbox
						try
						{
							bool check = AUIUtilities.FindElementAndToggle(CheckBoxGoldStarId, aeNewCT, ToggleState.On);
							if (check)
								Thread.Sleep(3000);
							else
							{
								Console.WriteLine("FindElementAndToggle failed:" + CheckBoxGoldStarId);
								Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxGoldStarId, Constants.TEST);
								TestCheck = ConstCommon.TEST_FAIL;
								sErrorMessage = "FindElementAndToggle failed:" + CheckBoxGoldStarId;
								return;
							}
						}
						catch (Exception ex)
						{
							TestCheck = ConstCommon.TEST_FAIL;
							Console.WriteLine("Gold Star CheckBox :" + ex.Message + " --- " + ex.StackTrace);
							sErrorMessage = ex.Message + " --- " + ex.StackTrace;
						}
						#endregion
					}

					if (i == 3)    // CheckBoxProdStateId   checked
					{
						#region // product state checkbox
						try
						{
							bool check = AUIUtilities.FindElementAndToggle(CheckBoxProdStateId, aeNewCT, ToggleState.On);
							if (check)
								Thread.Sleep(3000);
							else
							{
								Console.WriteLine("FindElementAndToggle failed:" + CheckBoxProdStateId);
								Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxProdStateId, Constants.TEST);
								TestCheck = ConstCommon.TEST_FAIL;
								sErrorMessage = "FindElementAndToggle failed:" + CheckBoxProdStateId;
								return;
							}
						}
						catch (Exception ex)
						{
							TestCheck = ConstCommon.TEST_FAIL;
							Console.WriteLine("product state checkbox :" + ex.Message + " --- " + ex.StackTrace);
							sErrorMessage = ex.Message + " --- " + ex.StackTrace;
						}
						#endregion
					}

					if (i == 4)    // CheckBoxQualityStateId   checked
					{
						#region // quality state checkbox
						try
						{
							bool check = AUIUtilities.FindElementAndToggle(CheckBoxQualityStateId, aeNewCT, ToggleState.On);
							if (check)
								Thread.Sleep(3000);
							else
							{
								Console.WriteLine("FindElementAndToggle failed:" + CheckBoxQualityStateId);
								Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxQualityStateId, Constants.TEST);
								TestCheck = ConstCommon.TEST_FAIL;
								sErrorMessage = "FindElementAndToggle failed:" + CheckBoxQualityStateId;
								return;
							}
						}
						catch (Exception ex)
						{
							TestCheck = ConstCommon.TEST_FAIL;
							Console.WriteLine("quality state checkbox :" + ex.Message + " --- " + ex.StackTrace);
							sErrorMessage = ex.Message + " --- " + ex.StackTrace;
						}
						#endregion
					}

					if (i == 5)    // All checked
					{
						#region // blue star checkbox
						try
						{
							bool check = AUIUtilities.FindElementAndToggle(CheckBoxBlueStarId, aeNewCT, ToggleState.On);
							if (check)
								Thread.Sleep(3000);
							else
							{
								Console.WriteLine("FindElementAndToggle failed:" + CheckBoxBlueStarId);
								Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxBlueStarId, Constants.TEST);
								TestCheck = ConstCommon.TEST_FAIL;
								sErrorMessage = "FindElementAndToggle failed:" + CheckBoxBlueStarId;
								return;
							}
						}
						catch (Exception ex)
						{
							TestCheck = ConstCommon.TEST_FAIL;
							Console.WriteLine("Blue Star CheckBox :" + ex.Message + " --- " + ex.StackTrace);
							sErrorMessage = ex.Message + " --- " + ex.StackTrace;
						}
						#endregion

						#region // gold star checkbox
						try
						{
							bool check = AUIUtilities.FindElementAndToggle(CheckBoxGoldStarId, aeNewCT, ToggleState.On);
							if (check)
								Thread.Sleep(3000);
							else
							{
								Console.WriteLine("FindElementAndToggle failed:" + CheckBoxGoldStarId);
								Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxGoldStarId, Constants.TEST);
								TestCheck = ConstCommon.TEST_FAIL;
								sErrorMessage = "FindElementAndToggle failed:" + CheckBoxGoldStarId;
								return;
							}
						}
						catch (Exception ex)
						{
							TestCheck = ConstCommon.TEST_FAIL;
							Console.WriteLine("Gold Star CheckBox :" + ex.Message + " --- " + ex.StackTrace);
							sErrorMessage = ex.Message + " --- " + ex.StackTrace;
						}
						#endregion

						#region // product state checkbox
						try
						{
							bool check = AUIUtilities.FindElementAndToggle(CheckBoxProdStateId, aeNewCT, ToggleState.On);
							if (check)
								Thread.Sleep(3000);
							else
							{
								Console.WriteLine("FindElementAndToggle failed:" + CheckBoxProdStateId);
								Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxProdStateId, Constants.TEST);
								TestCheck = ConstCommon.TEST_FAIL;
								sErrorMessage = "FindElementAndToggle failed:" + CheckBoxProdStateId;
								return;
							}
						}
						catch (Exception ex)
						{
							TestCheck = ConstCommon.TEST_FAIL;
							Console.WriteLine("product state checkbox :" + ex.Message + " --- " + ex.StackTrace);
							sErrorMessage = ex.Message + " --- " + ex.StackTrace;
						}
						#endregion

						#region // quality state checkbox
						try
						{
							bool check = AUIUtilities.FindElementAndToggle(CheckBoxQualityStateId, aeNewCT, ToggleState.On);
							if (check)
								Thread.Sleep(3000);
							else
							{
								Console.WriteLine("FindElementAndToggle failed:" + CheckBoxQualityStateId);
								Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxQualityStateId, Constants.TEST);
								TestCheck = ConstCommon.TEST_FAIL;
								sErrorMessage = "FindElementAndToggle failed:" + CheckBoxQualityStateId;
								return;
							}
						}
						catch (Exception ex)
						{
							TestCheck = ConstCommon.TEST_FAIL;
							Console.WriteLine("quality state checkbox :" + ex.Message + " --- " + ex.StackTrace);
							sErrorMessage = ex.Message + " --- " + ex.StackTrace;
						}
						#endregion
					}

					string SaveID = "m_btnSave";
					// Find Save Button  element. 
					AutomationElement aeSaveBtn = AUIUtilities.FindElementByID(SaveID, aeNewCT);
					if (aeSaveBtn == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Save Button aeSaveBtn not found";
						Console.WriteLine("FindElementByID failed:" + SaveID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeSaveBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
					}
				}
				#endregion

				#region  // Validation results

				 AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find QualityReasonsDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("QualityReasonsDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

				#region // calculate startRow
				int startRow = 0;
				SqlConnection myConnection = new SqlConnection(sConnectionString);
				try
				{
					try
					{
						myConnection.Open();
					}
					catch (Exception e)
					{
						string exc = e.Message + System.Environment.NewLine + e.StackTrace;
						Console.WriteLine(e.ToString());
						MessageBox.Show(e.Message + System.Environment.NewLine + e.StackTrace);
						Epia3Common.WriteTestLogPass(slogFilePath, exc, Constants.TEST);
						return;
					}

					string sqlCommand = "select COUNT(*) from [Ewcs].[dbo].[ReelQualityReasons]";
					SqlCommand myCommand = new SqlCommand(sqlCommand, myConnection);
					int count = (int)myCommand.ExecuteScalar();

					string mMsg = "sql command:" + sqlCommand + System.Environment.NewLine;
					if (count < 20 )
					{
						startRow = 0;
					}
					else
					{
						startRow = count - 6;   
					}

					myConnection.Close();
				}
				catch (Exception ex)
				{
					myConnection.Close();
					sErrorMessage = sErrorMessage + " CONN:" + sConnectionString;
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = ex.Message + "----: " + ex.StackTrace;
					Console.WriteLine("Fatal error: " + sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					return;
				}
				#endregion

				for (int i = 1; i < 6; i++)   // 
				{
					Thread.Sleep(3000); 
					string header = "Description";
					string cellValue = "ReelQualityReason"+i;
					int cellRow = -1;
					
					AutomationElement aeCell = FindCellFromGridStartedAtRow(aeGrid, header, cellValue, startRow+i, ref cellRow); 
					
					if (aeCell == null)
					{
						sErrorMessage = "Find QualityReasonsDataGridView aeCell failed:" + cellValue;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell QualityReasonsDataGridView found: " + cellValue);
					}
				}

				result = ConstCommon.TEST_PASS;
				Console.WriteLine(testname + " ---pass --- ");
				Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);

				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region AddCarrier
		public static void AddCarrier(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;
		   
			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);
				#region Find and open overview screen and click New button
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, CARRIERS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);
				
				// Find Carriers Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIERS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the CarriersOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = REELS_OVERVIEW_TITLE + " Window not found";
					Console.WriteLine("FindElementByID failed:" + REELS_OVERVIEW_TITLE);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find tool bar 
				AutomationElement aeToolBar = AUIUtilities.FindElementByType(ControlType.ToolBar, aeOverview);
				if (aeToolBar == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "ToolBar " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("ToolBar " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				// Find Edit textbox 
				AutomationElement aeEdit = AUIUtilities.FindElementByType(ControlType.Edit, aeToolBar);
				if (aeEdit == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "EditTextBox " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("EditTextBox " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Double Bottom = aeEdit.Current.BoundingRectangle.Bottom;
				Double Left = aeEdit.Current.BoundingRectangle.Left;
				Double Right = aeEdit.Current.BoundingRectangle.Right;
				Double Width = aeEdit.Current.BoundingRectangle.Width;
				Double Height = aeEdit.Current.BoundingRectangle.Height;
				Double X = aeEdit.Current.BoundingRectangle.X;
				Double Y = aeEdit.Current.BoundingRectangle.Y;
				Double Top = aeEdit.Current.BoundingRectangle.Top;

				double xclick = Right + 66; 
				Point NewBtnPoint = new Point(xclick, (Bottom + Top) / 2);
				Input.MoveTo(NewBtnPoint);
				#endregion

				// Find search field and move to search field
				AutomationElement aeSF = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSF == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSF);
				Thread.Sleep(3000);
				Input.MoveToAndClick(aeSF);
				Thread.Sleep(1000);
				ValuePattern vpsx = (ValuePattern)aeSF.GetCurrentPattern(ValuePattern.Pattern);
				Thread.Sleep(1000);
				string getValuex = vpsx.Current.Value;
				Thread.Sleep(2000);
				

				#region // Add New Carriers
				for (int i = 0; i < 12; i++)
				{
					Thread.Sleep(3000);
					vpsx.SetValue(          ((i >= 10) ? "R000" + (i - 10) : "P000" + i)           );
					

					Thread.Sleep(3000);
					Input.MoveToAndClick(NewBtnPoint);
					Thread.Sleep(3000);

					//find new Egemin Shell Window screen
					System.Windows.Automation.Condition c2p = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					// Find the New ReelInput Screen element.
					AutomationElement aeNewCarrierInput = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2p);
					if (aeNewCarrierInput == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "New Reels Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}

					// select Carrier Type

					AutomationElement aeCombo = AUIUtilities.FindElementByID("m_ComboCarrierType", aeNewCarrierInput);
					if (aeCombo == null)
					{
						Console.WriteLine("failed to find CarrierType aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						sErrorMessage = "failed to find Carriertype aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss");
						TestCheck = ConstCommon.TEST_FAIL;
						sEventEnd = true;
						return;
					}

					SelectionPattern selectPattern =
					   aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

					Thread.Sleep(1000);
					
					AutomationElement item
						= AUIUtilities.FindElementByName( (i >=10) ? "RackType" : "PalletType1", aeCombo);
					if (item != null)
					{
						Console.WriteLine("PalletType1 item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						Thread.Sleep(2000);

						SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
						itemPattern.Select();
					}
					else
					{
						Console.WriteLine("Finding Language item nl failed");
						sErrorMessage = "Finding Language item nl failed";
						TestCheck = ConstCommon.TEST_FAIL;
						sEventEnd = true;
						return;
					}

					Thread.Sleep(3000);
					string CarrierAutoID = "m_TxtCarrierId";
					// "m_TxtCarrierId"   
					// "m_TxtUnitId"  This is Edit Control, should use setFocus + sendKeys
					if (AUIUtilities.FindDocumentAndSendText(CarrierAutoID, aeNewCarrierInput, (i >=10) ? "R000"+(i-10) : "P000"+i, ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						MessageBox.Show(sErrorMessage);
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxCarrierID failed:" + CarrierAutoID);
						sErrorMessage = "FindTextBoxCarrierID failed:" + CarrierAutoID;
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				   
					
					Thread.Sleep(3000);
					string EditLocationID = "m_TxtLocationId";
					if (AUIUtilities.FindDocumentAndSendText(EditLocationID, aeNewCarrierInput, (i >= 10) ? "RAT.STOCK" : "LostAndFound", ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						MessageBox.Show(sErrorMessage);
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxLocationID failed:" + EditLocationID);
						sErrorMessage = "FindTextBoxLocationID failed:" + EditLocationID;
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}

					/*
					//----------------------------------------------------------------
					//  "m_BtnSearchLocation"
					string BtnSearchLocationID = "m_BtnSearchLocation";
					// Find Save Button  element. 
					AutomationElement aeSearchLocationBtn = AUIUtilities.FindElementByID(BtnSearchLocationID, aeNewCarrierInput);
					if (aeSearchLocationBtn == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Search Location Button aeSearchLocationBtn not found";
						Console.WriteLine("FindElementByID failed:" + BtnSearchLocationID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeSearchLocationBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSearchLocationBtn));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSearchLocationBtn));
					}
					// This is visiable Screen
					//find new Egemin Shell Window screen
					System.Windows.Automation.Condition c2s = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					// Find the New LocationID Screen element. ps not from root should be 
					AutomationElement aeNewLocSearch = aeNewCarrierInput.FindFirst(TreeScope.Element | TreeScope.Descendants, c2s);
					if (aeNewLocSearch == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "New Location Search Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeNewLocSearch Screen : Now find Grid ");
					}

					Thread.Sleep(3000);

					AutomationElement aeSearchLocGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeNewLocSearch);
					if (aeSearchLocGrid == null)
					{
						sErrorMessage = "Find LocationSearchGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("LocationSearchGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					Thread.Sleep(3000);
					// todo try to input a predefined location 
					// Construct the Grid Cell Element Name
					string LocCellname = sLocationIDHeader+" Row " + i;
					// Get the Element with the Row Col Coordinates        
					AutomationElement aeCell = AUIUtilities.FindElementByName(LocCellname, aeSearchLocGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Find LocationsDataGridView aeCell failed:" + LocCellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell LocationsDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}
					Thread.Sleep(2000);

					Input.MoveToAndClick(aeCell);
				   
					Thread.Sleep(2000);

					string OKID2 = "m_btnSave";
					// Find Save Button  element. 
					AutomationElement aeOK2 = AUIUtilities.FindElementByID(OKID2, aeNewLocSearch);
					if (aeOK2 == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "OK2 Button aeOK2 not found";
						Console.WriteLine("FindElementByID failed:" + OKID2);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeOK2:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeOK2));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeOK2));
					}

				  
					Thread.Sleep(5000);

					Console.WriteLine("Save............" + ((i >= 10) ? "R000" + (i - 10) : "P000" + i)  );
					Thread.Sleep(5000);
					*/
					string SaveID = "m_btnSave";
					// Find Save Button  element. 
					AutomationElement aeSave = AUIUtilities.FindElementByID(SaveID, aeNewCarrierInput);
					if (aeSave == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Save Button aeSave not found";
						Console.WriteLine("FindElementByID failed:" + SaveID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeSave:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSave));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSave));
					}

					Thread.Sleep(8000);
					
					//   Check Error Screen ----------------------------

					// find Error screen
					string ErrorScreenID = "ErrorScreen";
					Console.WriteLine("Check Error Screen");
					Epia3Common.WriteTestLogMsg(slogFilePath, "Check Error Screen", Constants.TEST);
					AutomationElement aeErrorScreen = AUIUtilities.FindElementByID(ErrorScreenID, root);
					if (aeErrorScreen != null)
					{
						string ErrorTextID = "m_LblCaption";
						Thread.Sleep(2000);
						// Find  error text  element. 
						AutomationElement aeErrorText = AUIUtilities.FindElementByID(ErrorTextID, aeErrorScreen);
						if (aeErrorText == null)
						{
							result = ConstCommon.TEST_FAIL;
							sErrorMessage = "error text aeErrorText not found";
							Console.WriteLine("FindElementByID failed:" + ErrorTextID);
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
							return;
						}
						else
						{
							sErrorMessage = aeErrorText.Current.Name;
							result = ConstCommon.TEST_FAIL;
							Console.WriteLine("FindElementByID failed:" + SaveID);
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);

							// close Error Screen
							string CloseBtnID = "m_BtnClose";
							Thread.Sleep(2000);
							// Find  close btn element. 
							AutomationElement aeCloseBtn = AUIUtilities.FindElementByID(CloseBtnID, aeErrorScreen);
							if (aeCloseBtn == null)
							{
								result = ConstCommon.TEST_FAIL;
								sErrorMessage = "Close button aeCloseBtn not found";
								Console.WriteLine("FindElementByID failed:" + CloseBtnID);
								Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
								return;
							}
							Console.WriteLine("aeCloseBtn:");
							Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCloseBtn));
							Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCloseBtn));
						}

						return;
						Thread.Sleep(2000);

					}
					else
					{
						Epia3Common.WriteTestLogMsg(slogFilePath, "No Error Screen found 0", Constants.TEST);
					}

					Console.WriteLine("No Error Screen found");
					Epia3Common.WriteTestLogMsg(slogFilePath, "No Error Screen found 1", Constants.TEST);
					//----------------------------
				}
				#endregion
				Thread.Sleep(5000);
				//----------------------------------------------------------------
				#region // Validation results
				
				// Find Search Field
				AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField);
				Thread.Sleep(3000);
				Input.MoveToAndClick(aeSearchField);
				Thread.Sleep(1000);
				ValuePattern vps = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
				Thread.Sleep(1000);
				string getValue = vps.Current.Value;
				Thread.Sleep(2000);
				vps.SetValue("P000");
				Thread.Sleep(1000);  

				AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find CarrierTypeDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("CarrierTypeDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));


				for (int i = 0; i < 12; i++)
				{
					if (i >= 10)
					{
						Input.MoveTo(aeSearchField);
						Thread.Sleep(3000);
						Input.MoveToAndClick(aeSearchField);
						Thread.Sleep(1000);
						ValuePattern vp2 = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
						Thread.Sleep(1000);
						getValue = vps.Current.Value;
						Thread.Sleep(2000);
						vp2.SetValue("R000");
						Thread.Sleep(1000);
					}
					Thread.Sleep(3000);
					// Construct the Grid Cell Element Name
					string cellname = sCarrierIDHeader+" Row " + ( (i>=10)? (i-10) :i);
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Validation: Find CarriersDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell CarrierDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					// find cell value
					string ProductValue = string.Empty;
					try
					{
						ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						ProductValue = vp.Current.Value;
						Console.WriteLine("Get element.Current Value:" + ProductValue);
					}
					catch (System.NullReferenceException)
					{
						ProductValue = string.Empty;
					}

					if (ProductValue == null || ProductValue == string.Empty)
					{
						sErrorMessage = "Validation: CarriersDataGridView aeCell Value not found:" + cellname;
						Console.WriteLine("CarriersDataGridView aeCell Value not found:" + cellname);
						Epia3Common.WriteTestLogMsg(slogFilePath, "CarriersDataGridView cell value not found:" + cellname, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					if (ProductValue.Equals((i >=10) ? "R000"+(i-10) : "P000"+i))
					{
						result = ConstCommon.TEST_PASS;
						Console.WriteLine(testname + " ---pass --- " + ProductValue);
						//Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					}
					else
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "CarriersDataGridView aeCell Value not correct:" + ProductValue;
						Console.WriteLine(testname + " ---fail --- " + ProductValue);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				}

				Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);

				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region AddProduct
		public static void AddProduct(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);
				#region Find and open overview screen and move to New button
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, PRODUCTS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Products Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, PRODUCTS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ProductsOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = PRODUCTS_OVERVIEW_TITLE + " Window not found";
					Console.WriteLine("FindElementByID failed:" + PRODUCTS_OVERVIEW_TITLE);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find New... Button
				// Find tool bar 
				AutomationElement aeToolBar = AUIUtilities.FindElementByType(ControlType.ToolBar, aeOverview);
				if (aeToolBar == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "ToolBar " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("ToolBar " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				// Find Edit textbox 
				AutomationElement aeEdit = AUIUtilities.FindElementByType(ControlType.Edit, aeToolBar);
				if (aeEdit == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "EditTextBox " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("EditTextBox " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Double Bottom = aeEdit.Current.BoundingRectangle.Bottom;
				Double Left = aeEdit.Current.BoundingRectangle.Left;
				Double Right = aeEdit.Current.BoundingRectangle.Right;
				Double Width = aeEdit.Current.BoundingRectangle.Width;
				Double Height = aeEdit.Current.BoundingRectangle.Height;
				Double X = aeEdit.Current.BoundingRectangle.X;
				Double Y = aeEdit.Current.BoundingRectangle.Y;
				Double Top = aeEdit.Current.BoundingRectangle.Top;

				double xclick = Right + 62;
				Point NewBtnPoint = new Point(xclick, (Bottom + Top) / 2);
				Input.MoveTo(NewBtnPoint);
				#endregion

				#region // Add New Products
				for (int i = 1; i < 5; i++)
				{
					Thread.Sleep(3000);

					Input.MoveToAndClick(NewBtnPoint);

					//find new Egemin Shell Window screen
					System.Windows.Automation.Condition c2p = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					// Find the New ProductID Screen element.
					AutomationElement aeNewPro = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2p);
					if (aeNewPro == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "New Product Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}

					Thread.Sleep(3000);

					string origUnitIdMask = "";
					// "unitIdMaskTextBox"
					if (AUIUtilities.FindTextBoxAndChangeValue("unitIdMaskTextBox", 
						aeNewPro, out origUnitIdMask,(i>=3)?"00000000": "AAAAA_00000000-000A", ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "unitIdMaskTextBox");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "unitIdMaskTextBox";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}

					Thread.Sleep(3000);

					string origProdId = "";
					// "m_TxtProductId"
					if (AUIUtilities.FindTextBoxAndChangeValue("m_TxtProductId",
						aeNewPro, out origProdId, (i >= 3) ? "AutoTestOldProd"+(i-2) : "AutoTestProduct" + i, ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "m_TxtProductId");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "m_TxtProductId";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}

					//Thread.Sleep(3000);
					//string origDesc = "";
					// "descriptionTextBox"  This is document Control temp not execute
					if (AUIUtilities.FindDocumentAndSendText("descriptionTextBox", aeNewPro, "AutoTestDescription"+i, ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						MessageBox.Show(sErrorMessage);
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "descriptionTextBox");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "descriptionTextBox";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}

					Thread.Sleep(3000);

					string SaveID = "m_btnSave";
					// Find Save Button  element. 
					AutomationElement aeSaveBtn = AUIUtilities.FindElementByID(SaveID, aeNewPro);
					if (aeSaveBtn == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Save Button aeSaveBtn not found";
						Console.WriteLine("FindElementByID failed:" + SaveID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeSaveBtn:" + ((i >= 3) ? "AutoTestOldProd" + (i - 2) : "AutoTestProduct" + i)  );
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
					}
					Thread.Sleep(3000);

				}
				#endregion

				#region // Validation results

				// Find Search Field
				AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField);
			   
				#region  // Search and Validate Products
				string productID = "AutoTestProduct";
				for (int i = 1; i < 5; i++)
				{
					productID = (i < 3) ? "AutoTestProduct" + i : "AutoTestOldProd" + (i - 2);
					Thread.Sleep(3000);

					Input.MoveToAndClick(aeSearchField);
					Thread.Sleep(1000);
					ValuePattern vp = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
					Thread.Sleep(1000);
					string getValue = vp.Current.Value;
					Thread.Sleep(2000);
					vp.SetValue(productID);

					AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
					if (aeGrid == null)
					{
						sErrorMessage = "Find ProductsDataGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
						Console.WriteLine("ProductsDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

					Thread.Sleep(3000);
					// Construct the Grid Cell Element Name
					string cellname = sProductIDHeader+" Row " + 0;
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Find ProductsDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell ProductsDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					// find cell value
					string ProductValue = string.Empty;
					try
					{
						ValuePattern vpx = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						ProductValue = vpx.Current.Value;
						Console.WriteLine("Get element.Current Value:" + ProductValue);
					}
					catch (System.NullReferenceException)
					{
						ProductValue = string.Empty;
					}

					if (ProductValue == null || ProductValue == string.Empty)
					{
						sErrorMessage = "ProductsDataGridView aeCell Value not found:" + cellname;
						Console.WriteLine("ProductsDataGridView aeCell Value not found:" + cellname);
						Epia3Common.WriteTestLogMsg(slogFilePath, "ProductsDataGridView cell value not found:" + cellname, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					if (ProductValue.Equals(productID))
					{
						result = ConstCommon.TEST_PASS;
						Console.WriteLine(testname + " ---pass --- " + ProductValue);
						Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					}
					else
					{
						sErrorMessage = "Validate failed:"+ProductValue;
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine(testname + " ---fail --- " + ProductValue);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				}
				#endregion

				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region AddReel
		public static void AddReel(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);

				#region // Open Overview screen
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, REELS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Reels Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty,REELS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ReelsOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = REELS_OVERVIEW_TITLE + " Window not found";
					Console.WriteLine("FindElementByID failed:" + REELS_OVERVIEW_TITLE);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				#endregion

				#region // Find and Move to New.. button

				// Find tool bar 
				AutomationElement aeToolBar = AUIUtilities.FindElementByType(ControlType.ToolBar, aeOverview);
				if (aeToolBar == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "ToolBar " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("ToolBar " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				// Find Edit textbox 
				AutomationElement aeEdit = AUIUtilities.FindElementByType(ControlType.Edit, aeToolBar);
				if (aeEdit == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "EditTextBox " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("EditTextBox " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Double Bottom = aeEdit.Current.BoundingRectangle.Bottom;
				Double Left = aeEdit.Current.BoundingRectangle.Left;
				Double Right = aeEdit.Current.BoundingRectangle.Right;
				Double Width = aeEdit.Current.BoundingRectangle.Width;
				Double Height = aeEdit.Current.BoundingRectangle.Height;
				Double X = aeEdit.Current.BoundingRectangle.X;
				Double Y = aeEdit.Current.BoundingRectangle.Y;
				Double Top = aeEdit.Current.BoundingRectangle.Top;

				double xclick = Right + 60;
				Point NewBtnPoint = new Point(xclick, (Bottom + Top) / 2);
				Input.MoveTo(NewBtnPoint);
				#endregion

				#region // Add New Reel
				//----------------------------------
				 // Find search field and move to search field
				AutomationElement aeSF = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSF == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}
				
				Input.MoveTo(aeSF);
				Thread.Sleep(3000);
				Input.MoveToAndClick(aeSF);
				Thread.Sleep(1000);
				ValuePattern vpsx = (ValuePattern)aeSF.GetCurrentPattern(ValuePattern.Pattern);
				Thread.Sleep(1000);
				string getValuex = vpsx.Current.Value;
				Thread.Sleep(2000);
				
				//-----------------------------------
				for (int i = 0; i < 10; i++)
				{

					Thread.Sleep(3000);
					vpsx.SetValue("NFPM1_20090403-00"+i+"A"); 
					Thread.Sleep(3000);

					Input.MoveToAndClick(NewBtnPoint);
					Thread.Sleep(3000);
					Input.MoveToAndClick(NewBtnPoint);
					Thread.Sleep(3000);

					//find new Egemin Shell Window screen
					System.Windows.Automation.Condition c2p = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					// Find the New ReelInput Screen element.
					AutomationElement aeNewReelInput = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2p);
					if (aeNewReelInput == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "New Reels Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}

					Thread.Sleep(2000);

					// Find the Reel button.
					AutomationElement aeReelButton = AUIUtilities.FindElementByID("m_btnCatReel", aeNewReelInput);
					if (aeReelButton == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Cat Reels button not found";
						Console.WriteLine("FindElementByID failed:" + "m_btnCatReel");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					Thread.Sleep(3000);

					string origUnitIdMask = "";
					// "batchTextBox"
					// if first time not found, than click reel butto  and try again
					bool found = false;
					int myTry = 0;
					while ( found == false && myTry < 2)
					{
						if (AUIUtilities.FindTextBoxAndChangeValue("batchTextBox", aeNewReelInput, out origUnitIdMask, "2", ref sErrorMessage))
						{
							found = true;
							
						}
						else
						{
							Input.MoveToAndClick(aeReelButton);
							myTry++;
						}
						Thread.Sleep(2000);
					}

					if (found == false )
					{
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "batchTextBox");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "batchTextBox";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						sErrorMessage = string.Empty;
						//MessageBox.Show(sErrorMessage);
						return;
					}
					
					//----------------------------------------------------------------
					//  "m_BtnSearchProduct"
					string BtnSearchProductID = "m_BtnSearchProduct";
					// Find Save Button  element. 
					AutomationElement aeSearchProductBtn = AUIUtilities.FindElementByID(BtnSearchProductID, aeNewReelInput);
					if (aeSearchProductBtn == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Search Product Button aeSearchProductBtn not found";
						Console.WriteLine("FindElementByID failed:" + BtnSearchProductID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeSearchProductBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSearchProductBtn));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSearchProductBtn));
					}

					//find new Egemin Shell Window screen for searching Products
					System.Windows.Automation.Condition c2l = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					// Find the New ProductID Screen element.
					AutomationElement aeNewprodSearchscreen = aeNewReelInput.FindFirst(TreeScope.Element | TreeScope.Descendants, c2l);
					if (aeNewprodSearchscreen == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "New Product Search Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}

					Thread.Sleep(2000);
					Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeNewprodSearchscreen));
					Thread.Sleep(2000);

					// Search AutoTestProduct1
					 // Find Search Field
					AutomationElement aeSearchField1 = AUIUtilities.FindElementByType(ControlType.Edit, aeNewprodSearchscreen);
					if (aeSearchField1 == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
						Console.WriteLine(sErrorMessage);
						return;
					}
					else
					{
						Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						//aeSearchField1.SetFocus();
					}

					Thread.Sleep(1000);
					Input.MoveTo(aeSearchField1);
					Thread.Sleep(2000);
					
					Input.MoveToAndClick(aeSearchField1);
					Thread.Sleep(2000);
					ValuePattern vp = (ValuePattern)aeSearchField1.GetCurrentPattern(ValuePattern.Pattern);
					Thread.Sleep(1000);
					string value = vp.Current.Value;
					Console.WriteLine("set value AutoTestProduct1 to search field");
					Thread.Sleep(1000);
					vp.SetValue("AutoTestProduct1");
					//System.Windows.Forms.SendKeys.SendWait("AutoTestProduct1");
					Thread.Sleep(1000);

					AutomationElement aeSearchGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeNewprodSearchscreen);
					if (aeSearchGrid == null)
					{
						sErrorMessage = "Find ProductSearchGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
						Console.WriteLine("ProductSearchGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

					Thread.Sleep(2000);
					Console.WriteLine("search product: " + "AutoTestProduct1");
					int row = -1;
					AutomationElement aeProd = FindCellFromGrid(aeSearchGrid, sProductIDHeader, "AutoTestProduct1", ref row);

					if (aeProd == null)
					{
						sErrorMessage = "Find Product  aeCell failed:" + "AutoTestProduct1";
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;

						#region // Cancel screen
						string CancelId = "m_btnCancel";
						// Find Cancel Button  element. 
						AutomationElement aeCancelBtn = AUIUtilities.FindElementByID(CancelId, aeNewprodSearchscreen);
						if (aeCancelBtn == null)
						{
							result = ConstCommon.TEST_FAIL;
							sErrorMessage = "Cancel Button aeCancelBtn not found";
							Console.WriteLine("FindElementByID failed:" + CancelId);
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
							return;
						}
						else
						{
							Console.WriteLine(i + ": found aeCancelBtn:");
							Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn));
							Thread.Sleep(2000);
							Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCancelBtn));
							Thread.Sleep(2000);

							AutomationElement aeCancelBtn2 = AUIUtilities.FindElementByID(CancelId, aeNewReelInput);
							if (aeCancelBtn2 == null)
							{
								result = ConstCommon.TEST_FAIL;
								sErrorMessage = "Cancel Button aeCancelBtn2 not found";
								Console.WriteLine("FindElementByID failed 2 :" + CancelId);
								Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
								return;
							}
							else
							{
								Console.WriteLine(i + ": found aeCancelBtn2:");
								Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn2));
								Thread.Sleep(2000);
								Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCancelBtn2));
								Thread.Sleep(2000);
							}

						}
						#endregion
						return;
					}
					else
					{
						Console.WriteLine(i + ": found row at: " + row);
						Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeProd));

					}
					Thread.Sleep(3000);
					
					string OKID = "m_btnSave";
					// Find Save Button  element. 
					AutomationElement aeOKBtn = AUIUtilities.FindElementByID(OKID, aeNewprodSearchscreen);
					if (aeOKBtn == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "OK Button aeOKBtn not found";
						Console.WriteLine("FindElementByID failed:" + OKID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine(i + ": found aeOKBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeOKBtn));
						Thread.Sleep(2000);
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeOKBtn));
						Thread.Sleep(2000);
					}
				   
					// "m_TxtUnitId"  This is Edit Control, should use setFocus + sendKeys
					if (AUIUtilities.FindDocumentAndSendText("m_TxtUnitId", aeNewReelInput, "NFPM1_20090403-00"+i+"A", ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						MessageBox.Show(sErrorMessage);
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "m_TxtUnitId");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "m_TxtUnitId";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}

					// "m_TxtPaperBulk"  This is Edit Control, should use setFocus + sendKeys
					if (AUIUtilities.FindDocumentAndSendText("m_TxtPaperBulk", aeNewReelInput, "20", ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						MessageBox.Show(sErrorMessage);
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "m_TxtPaperBulk");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "m_TxtPaperBulk";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}

					// "m_TxtDiameterMeter"  This is Edit Control, should use setFocus + sendKeys
					if (AUIUtilities.FindDocumentAndSendText("m_TxtDiameterMeter", aeNewReelInput, "3", ref sErrorMessage))
						Thread.Sleep(3000);   //999
					else
					{
						MessageBox.Show(sErrorMessage);
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "m_TxtDiameterMeter");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "m_TxtDiameterMeter";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}                    

					// "m_TxtPaperWidthMeter"  This is Edit Control, should use setFocus + sendKeys
					if (AUIUtilities.FindDocumentAndSendText("m_TxtPaperWidthMeter", aeNewReelInput, "2", ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						MessageBox.Show(sErrorMessage);
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + "m_TxtPaperWidthMeter");
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + "m_TxtPaperWidthMeter";
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}

					//**********************************
					//  "m_BtnSearchLocation"
					string BtnSearchLocationID = "m_BtnSearchLocation";
					// Find Save Button  element. 
					AutomationElement aeSearchLocationBtn = AUIUtilities.FindElementByID(BtnSearchLocationID, aeNewReelInput);
					if (aeSearchLocationBtn == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Search Location Button aeSearchLocationBtn not found";
						Console.WriteLine("FindElementByID failed:" + BtnSearchLocationID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine(i + ": aeSearchLocationBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSearchLocationBtn));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSearchLocationBtn));
					}
					// This is visiable Screen
					//find new Egemin Shell Window screen
					System.Windows.Automation.Condition c2s = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					// Find the New LocationID Screen element. ps not from root should be 
					AutomationElement aeNewLocSearch = aeNewReelInput.FindFirst(TreeScope.Element | TreeScope.Descendants, c2s);
					if (aeNewLocSearch == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "New Location Search Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine(i+": aeNewLocSearch Screen : Now find Grid ");
					}

					Thread.Sleep(1000);

					AutomationElement aeSearchLocGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeNewLocSearch);
					if (aeSearchLocGrid == null)
					{
						sErrorMessage = "Find LocationSearchGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine(i + ": LocationSearchGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					Thread.Sleep(1000);

					Double BottomL = aeSearchLocGrid.Current.BoundingRectangle.Top;
					Double LeftL = aeSearchLocGrid.Current.BoundingRectangle.Left;

					Console.WriteLine(i + ": Move to Grid leftcorner");
					Input.MoveToAndClick(new Point(LeftL, BottomL));
					Thread.Sleep(2000);

					// Construct the Grid Cell Element Name
					string LocCellname = sLocationIDHeader+" Row " + i;
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCell = AUIUtilities.FindElementByName(LocCellname, aeSearchLocGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Find LocationsDataGridView aeCell failed:" + LocCellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine(i + ": cell LocationsDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}
					Thread.Sleep(1000);

					Input.MoveToAndClick(aeCell);

				   
					Thread.Sleep(2000);

					string OKID2 = "m_btnSave";
					// Find Save Button  element. 
					AutomationElement aeOK2 = AUIUtilities.FindElementByID(OKID2, aeNewLocSearch);
					if (aeOK2 == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "OK2 Button aeOK2 not found";
						Console.WriteLine("FindElementByID failed:" + OKID2);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine(i + ": aeOK2:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeOK2));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeOK2));
					}

					Console.WriteLine(i + ": OK2............OK");
					Console.WriteLine(i + ": Save............");
					Thread.Sleep(1000);

					string SaveID = "m_btnSave";
					string CancelID = "m_btnCancel";
					// Find Save Button  element. 
					AutomationElement aeSave = AUIUtilities.FindElementByID(SaveID, aeNewReelInput);
					if (aeSave == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Save Button aeSave not found";
						Console.WriteLine("FindElementByID failed:" + SaveID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine(i + ": aeSave:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSave));
						if (aeSave.Current.IsEnabled )
						{
							Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSave));
						}
						else   
						{
							result = ConstCommon.TEST_FAIL;
							sErrorMessage = i+":aeSave button is unabled";
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
							Thread.Sleep(5000);
							AutomationElement aeCancel = AUIUtilities.FindElementByID(CancelID, aeNewReelInput);
							if (aeCancel == null)
							{
								sErrorMessage = i+":aeSave button is unabled"+ "and Cancel Button aeCancel not found";
								Console.WriteLine("FindElementByID failed:" + CancelID);
								Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
								return;
							}
							else
							{
								Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancel));
								Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCancel));
							}
							return;
						}

					}

					Thread.Sleep(3000);

					//   Check Error Screen ----------------------------

					// find Error screen
					string ErrorScreenID = "ErrorScreen";
					AutomationElement aeErrorScreen = AUIUtilities.FindElementByID(ErrorScreenID, root);
					if (aeErrorScreen != null)
					{
						string ErrorTextID = "m_LblCaption";
						Thread.Sleep(2000);
						// Find  error text  element. 
						AutomationElement aeErrorText = AUIUtilities.FindElementByID(ErrorTextID, aeErrorScreen);
						if (aeErrorText == null)
						{
							result = ConstCommon.TEST_FAIL;
							sErrorMessage = "error text aeErrorText not found";
							Console.WriteLine("FindElementByID failed:" + ErrorTextID);
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
							return;
						}
						else
						{
							sErrorMessage = aeErrorText.Current.Name;
							result = ConstCommon.TEST_FAIL;
							Console.WriteLine("FindElementByID failed:" + SaveID);
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);

							// close Error Screen
							string CloseBtnID = "m_BtnClose";
							Thread.Sleep(2000);
							// Find  close btn element. 
							AutomationElement aeCloseBtn = AUIUtilities.FindElementByID(CloseBtnID, aeErrorScreen);
							if (aeCloseBtn == null)
							{
								result = ConstCommon.TEST_FAIL;
								sErrorMessage = "Close button aeCloseBtn not found";
								Console.WriteLine("FindElementByID failed:" + CloseBtnID);
								Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
								return;
							}
							Console.WriteLine("aeCloseBtn:");
							Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCloseBtn));
							Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCloseBtn));
						}

						return;
						Thread.Sleep(2000);

					}
					//----------------------------
				}
				#endregion

				Thread.Sleep(3000);
			   
				#region // Validation results

				// Find Search Field
				AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Validation: Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField);
				Thread.Sleep(3000);
				Input.MoveToAndClick(aeSearchField);
				Thread.Sleep(1000);
				ValuePattern vps = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
				Thread.Sleep(1000);
				string getValue = vps.Current.Value;
				Thread.Sleep(2000);
				vps.SetValue("20090403");
				Thread.Sleep(1000);

				AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Validation: Find ReelsDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("ReelsDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

				for (int i = 0; i < 10; i++)
				{                    
					Thread.Sleep(3000);
					// Construct the Grid Cell Element Name
					string cellname = sUnitIDHeader+" Row "+i;
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Validation: Find ReelsDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine(i+": cell ReelsDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					// find cell value
					string ProductValue = string.Empty;
					try
					{
						ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						ProductValue = vp.Current.Value;
						Console.WriteLine("Get element.Current Value:" + ProductValue);
					}
					catch (System.NullReferenceException)
					{
						ProductValue = string.Empty;
					}

					if (ProductValue == null || ProductValue == string.Empty)
					{
						sErrorMessage = "Validation: ReelsDataGridView aeCell Value not found:" + cellname;
						Console.WriteLine("ProductDataGridView aeCell Value not found:" + cellname);
						Epia3Common.WriteTestLogMsg(slogFilePath, "ReelsDataGridView cell value not found:" + cellname, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					if (ProductValue.Equals("NFPM1_20090403-00"+i+"A"))
					{
						result = ConstCommon.TEST_PASS;
						Console.WriteLine(testname + " ---pass --- " + ProductValue);
						Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					}
					else
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Validation: ReelsDataGridView aeCell Value not correct:" + ProductValue;
						Console.WriteLine(testname + " ---fail --- " + ProductValue);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				}
				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion

		#region SearchCarrierTypes
		public static void SearchCarrierTypes(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			string PALLET_TYPE = "PalletType";
			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, CONFIGURATION, CARRIER_TYPES, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Carrier Types Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIER_TYPES_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the CarrierTypesOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = CARRIER_TYPES + " Window not found";
					Console.WriteLine("FindElementByID failed:" + CARRIER_TYPES);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find Search Field
				AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField);

				#region  Search CarrierTypes
				for (int i = 1; i < 4; i++)
				{
					Thread.Sleep(3000);

					Input.MoveToAndClick(aeSearchField);
					Thread.Sleep(1000);
					ValuePattern vp = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
					Thread.Sleep(1000);
					string getValue = vp.Current.Value;
					Thread.Sleep(2000);
					vp.SetValue("Type"+i);

					AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
					if (aeGrid == null)
					{
						sErrorMessage = "Find CarrierTypeDataGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
						Console.WriteLine("CarrierTypeDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

					Thread.Sleep(3000);
					// Construct the Grid Cell Element Name
					string cellname = sCarrierTypeID+" Row " + 0;
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Find CarrierTypeDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell CarrierTypeDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					// find cell value
					string CarrierTypeValue = string.Empty;
					try
					{
						ValuePattern vpx = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						CarrierTypeValue = vpx.Current.Value;
						Console.WriteLine("Get element.Current Value:" + CarrierTypeValue);
					}
					catch (System.NullReferenceException)
					{
						CarrierTypeValue = string.Empty;
					}

					if (CarrierTypeValue == null || CarrierTypeValue == string.Empty)
					{
						sErrorMessage = "CarrierTypeDataGridView aeCell Value not found:" + cellname;
						Console.WriteLine("CarrierTypeDataGridView aeCell Value not found:" + cellname);
						Epia3Common.WriteTestLogMsg(slogFilePath, "CarrierTypeDataGridView cell value not found:" + cellname, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					if (CarrierTypeValue.Equals(PALLET_TYPE + i))
					{
						result = ConstCommon.TEST_PASS;
						Console.WriteLine(testname + " ---pass --- " + CarrierTypeValue);
						Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					}
					else
					{
						sErrorMessage = "CarrierType value shouldbe:" + PALLET_TYPE + i + ", but now is  " + CarrierTypeValue;
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine(testname + " ---fail --- " + CarrierTypeValue);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				}
				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region SearchCarriers
		public static void SearchCarriers(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			string CARRIER = "C1999";
			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, CARRIERS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Carrier Types Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIERS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the CarriersOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = CARRIER_TYPES + " Window not found";
					Console.WriteLine("FindElementByID failed:" + CARRIERS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find Search Field
				AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField);

				#region  Search CarrierTypes
				for (int i = 1; i < 2; i++)
				{
					Thread.Sleep(3000);

					Input.MoveToAndClick(aeSearchField);
					Thread.Sleep(1000);
					ValuePattern vp = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
					Thread.Sleep(1000);
					string getValue = vp.Current.Value;
					Thread.Sleep(2000);
					vp.SetValue(CARRIER);

					AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
					if (aeGrid == null)
					{
						sErrorMessage = "Find CarriersDataGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
						Console.WriteLine("CarriersDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

					Thread.Sleep(3000);
					// Construct the Grid Cell Element Name
					string cellname = sCarrierIDHeader+" Row " + 0;
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Find CarriersDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell CarriersDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					// find cell value
					string CarrierValue = string.Empty;
					try
					{
						ValuePattern vpx = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						CarrierValue = vpx.Current.Value;
						Console.WriteLine("Get element.Current Value:" + CarrierValue);
					}
					catch (System.NullReferenceException)
					{
						CarrierValue = string.Empty;
					}

					if (CarrierValue == null || CarrierValue == string.Empty)
					{
						sErrorMessage = "CarriersDataGridView aeCell Value not found:" + cellname;
						Console.WriteLine("CarriersDataGridView aeCell Value not found:" + cellname);
						Epia3Common.WriteTestLogMsg(slogFilePath, "CarriersDataGridView cell value not found:" + cellname, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					if (CarrierValue.Equals(CARRIER))
					{
						result = ConstCommon.TEST_PASS;
						Console.WriteLine(testname + " ---pass --- " + CarrierValue);
						Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					}
					else
					{
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine(testname + " ---fail --- " + CarrierValue);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				}
				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region SearchProducts
		public static void SearchProducts(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			string PRODUCT20 = "Product";
			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);
				#region // Open Products Overviez screen
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, PRODUCTS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Products Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, PRODUCTS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ProductsOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = PRODUCTS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + PRODUCTS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				#endregion

				// Find Search Field
				AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField);

				#region  Search Products
				for (int i = 1; i < 2; i++)
				{
					PRODUCT20 = "Product"+i;
					Thread.Sleep(3000);

					Input.MoveToAndClick(aeSearchField);
					Thread.Sleep(1000);
					ValuePattern vp = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
					Thread.Sleep(1000);
					string getValue = vp.Current.Value;
					Thread.Sleep(2000);
					vp.SetValue(PRODUCT20);

					AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
					if (aeGrid == null)
					{
						sErrorMessage = "Find ProductsDataGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
						Console.WriteLine("ProductsDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

					Thread.Sleep(3000);
					// Construct the Grid Cell Element Name
					string cellname = sProductIDHeader+" Row " + 0;
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Find ProductsDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell ProductsDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					// find cell value
					string ProductValue = string.Empty;
					try
					{
						ValuePattern vpx = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						ProductValue = vpx.Current.Value;
						Console.WriteLine("Get element.Current Value:" + ProductValue);
					}
					catch (System.NullReferenceException)
					{
						ProductValue = string.Empty;
					}

					if (ProductValue == null || ProductValue == string.Empty)
					{
						sErrorMessage = "ProductsDataGridView aeCell Value not found:" + cellname;
						Console.WriteLine("ProductsDataGridView aeCell Value not found:" + cellname);
						Epia3Common.WriteTestLogMsg(slogFilePath, "ProductsDataGridView cell value not found:" + cellname, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					if (ProductValue.Equals("AutoTest" + PRODUCT20))
					{
						result = ConstCommon.TEST_PASS;
						Console.WriteLine(testname + " ---pass --- " + ProductValue);
						Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					}
					else
					{
						sErrorMessage = ProductValue;
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine(testname + " ---fail --- " + ProductValue);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				}
				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region SearchReels
		public static void SearchReels(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			string SUFFIX = "-199A";
			#region // Open ReelOverview screen 
			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, REELS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find REELs Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, REELS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ReelsOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = REELS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + REELS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
			#endregion
				// Find Search Field
				AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField);

				#region  Search Reels
				for (int i = 1; i < 3; i++)
				{
					Thread.Sleep(3000);

					Input.MoveToAndClick(aeSearchField);
					Thread.Sleep(1000);
					ValuePattern vp = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
					Thread.Sleep(1000);
					string getValue = vp.Current.Value;
					Thread.Sleep(2000);
					vp.SetValue((i>1)?"12340099":SUFFIX);

					AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
					if (aeGrid == null)
					{
						sErrorMessage = "Find ReelsDataGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
						Console.WriteLine("ReelsDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

					Thread.Sleep(3000);
					// Construct the Grid Cell Element Name
					string cellname = sUnitIDHeader+" Row " + 0;
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Find ReelsDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell ReelsDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					// find cell value
					string ReelValue = string.Empty;
					try
					{
						ValuePattern vpx = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						ReelValue = vpx.Current.Value;
						Console.WriteLine("Get element.Current Value:" + ReelValue);
					}
					catch (System.NullReferenceException)
					{
						ReelValue = string.Empty;
					}

					if (ReelValue == null || ReelValue == string.Empty)
					{
						sErrorMessage = "CarrierTypeDataGridView aeCell Value not found:" + cellname;
						Console.WriteLine("CarrierTypeDataGridView aeCell Value not found:" + cellname);
						Epia3Common.WriteTestLogMsg(slogFilePath, "CarrierTypeDataGridView cell value not found:" + cellname, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					if (ReelValue.Equals((i>1)?"12340099":"NFPM1_20090430" + SUFFIX))
					{
						result = ConstCommon.TEST_PASS;
						Console.WriteLine(testname + " ---pass --- " + ReelValue);
						Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					}
					else
					{
						sErrorMessage = ReelValue;
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine(testname + " ---fail --- " + ReelValue);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				}
				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region SearchStorageLocations
		public static void SearchStorageLocations(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			string LOCATION = "AT48.12.08";
			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, STORAGE_LOCATIONS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Locations Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, STORAGE_LOCATIONS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the LocationsOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = STORAGE_LOCATIONS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + STORAGE_LOCATIONS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find Search Field
				AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField);

				#region  Search Locations
				for (int i = 1; i < 2; i++)
				{
					Thread.Sleep(3000);

					Input.MoveToAndClick(aeSearchField);
					Thread.Sleep(1000);
					ValuePattern vp = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
					Thread.Sleep(1000);
					string getValue = vp.Current.Value;
					Thread.Sleep(2000);
					vp.SetValue(LOCATION);

					AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
					if (aeGrid == null)
					{
						sErrorMessage = "Find LocationsDataGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
						Console.WriteLine("LocationsDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

					Thread.Sleep(3000);
					// Construct the Grid Cell Element Name
					string cellname = sLocationIDHeader+" Row " + 0;
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Find LocationDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell LocationDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					// find cell value
					string LocationValue = string.Empty;
					try
					{
						ValuePattern vpx = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						LocationValue = vpx.Current.Value;
						Console.WriteLine("Get element.Current Value:" + LocationValue);
					}
					catch (System.NullReferenceException)
					{
						LocationValue = string.Empty;
					}

					if (LocationValue == null || LocationValue == string.Empty)
					{
						sErrorMessage = "CarrierTypeDataGridView aeCell Value not found:" + cellname;
						Console.WriteLine("CarrierTypeDataGridView aeCell Value not found:" + cellname);
						Epia3Common.WriteTestLogMsg(slogFilePath, "CarrierTypeDataGridView cell value not found:" + cellname, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					if (LocationValue.Equals(LOCATION))
					{
						result = ConstCommon.TEST_PASS;
						Console.WriteLine(testname + " ---pass --- " + LocationValue);
						Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					}
					else
					{
						sErrorMessage = LocationValue;
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine(testname + " ---fail --- " + LocationValue);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				}
				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region EditReelQualityReasons
		public static void EditReelQualityReasons(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);
				#region // Find and open overview screen
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, CONFIGURATION, REEL_QUALITY_REASONS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Reel Quality Reasons Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, REEL_QUALITY_REASONS_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ReelQualityReasonsOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = REEL_QUALITY_REASONS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + REEL_QUALITY_REASONS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				#endregion

				#region  Edit Reel Quality Reasons
				Thread.Sleep(2000);

				AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find QualityReasonsDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("QualityReasonsDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				// Find aeCell of Bobine Humide

				string header = "Description";
				string cellValue = "Bobine Humide";
				int cellRow = -1;
				AutomationElement aeCell = FindCellFromGrid(aeGrid, header, cellValue, ref cellRow);

				if (aeCell == null)
				{
					sErrorMessage = "Find QualityReasonsDataGridView aeCell failed:" + cellValue;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("cell QualityReasonsGridView found: " + cellValue);
					Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeCell));
				}

				Thread.Sleep(3000);

				//find new Egemin Shell Window screen
				System.Windows.Automation.Condition c2x = new AndCondition(
				  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
				  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
				);

				// Find the New Reel Quality Reasons Screen element.
				AutomationElement aeNewCT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2x);
				if (aeNewCT == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "New Reel Quality Reasons Window not found";
					Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				
				string DescriptionId = "descriptionTextBox";
				string CheckBoxBlueStarId = "validForBlueStarCheckBox";
				string CheckBoxGoldStarId = "validForGoldStarCheckBox";
				string CheckBoxProdStateId = "validForProductionStateCheckBox";
				string CheckBoxQualityStateId = "validForQualityStateCheckBox";
				string description = AUIUtilities.FindTextBoxAndValue(DescriptionId, aeNewCT, ref sErrorMessage);
				if (description == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "New Reel Quality Reasons Window not found";
					Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				else
				{
					if (!description.Equals(cellValue))
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = " description should be:" + cellValue + " not:" + description;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
				}
				   
				#region // blue star checkbox
				try
				{
					bool check = AUIUtilities.FindElementAndToggle(CheckBoxBlueStarId, aeNewCT, ToggleState.Off);
					if (check)
						Thread.Sleep(3000);
					else
					{
						Console.WriteLine("FindElementAndToggle failed:" + CheckBoxBlueStarId);
						Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxBlueStarId, Constants.TEST);
						TestCheck = ConstCommon.TEST_FAIL;
						sErrorMessage = "FindElementAndToggle failed:" + CheckBoxBlueStarId;
						return;
					}
				}
				catch (Exception ex)
				{
					TestCheck = ConstCommon.TEST_FAIL;
					Console.WriteLine("Blue Star CheckBox :" + ex.Message + " --- " + ex.StackTrace);
					sErrorMessage = ex.Message + " --- " + ex.StackTrace;
				}
				#endregion
			
				#region // gold star checkbox
				try
				{
					bool check = AUIUtilities.FindElementAndToggle(CheckBoxGoldStarId, aeNewCT, ToggleState.Off);
					if (check)
						Thread.Sleep(3000);
					else
					{
						Console.WriteLine("FindElementAndToggle failed:" + CheckBoxGoldStarId);
						Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxGoldStarId, Constants.TEST);
						TestCheck = ConstCommon.TEST_FAIL;
						sErrorMessage = "FindElementAndToggle failed:" + CheckBoxGoldStarId;
						return;
					}
				}
				catch (Exception ex)
				{
					TestCheck = ConstCommon.TEST_FAIL;
					Console.WriteLine("Gold Star CheckBox :" + ex.Message + " --- " + ex.StackTrace);
					sErrorMessage = ex.Message + " --- " + ex.StackTrace;
				}
				#endregion

				#region // product state checkbox
				try
				{
					bool check = AUIUtilities.FindElementAndToggle(CheckBoxProdStateId, aeNewCT, ToggleState.Off);
					if (check)
						Thread.Sleep(3000);
					else
					{
						Console.WriteLine("FindElementAndToggle failed:" + CheckBoxProdStateId);
						Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxProdStateId, Constants.TEST);
						TestCheck = ConstCommon.TEST_FAIL;
						sErrorMessage = "FindElementAndToggle failed:" + CheckBoxProdStateId;
						return;
					}
				}
				catch (Exception ex)
				{
					TestCheck = ConstCommon.TEST_FAIL;
					Console.WriteLine("product state checkbox :" + ex.Message + " --- " + ex.StackTrace);
					sErrorMessage = ex.Message + " --- " + ex.StackTrace;
				}
				#endregion
			
				#region // quality state checkbox
				try
				{
					bool check = AUIUtilities.FindElementAndToggle(CheckBoxQualityStateId, aeNewCT, ToggleState.Off);
					if (check)
						Thread.Sleep(3000);
					else
					{
						Console.WriteLine("FindElementAndToggle failed:" + CheckBoxQualityStateId);
						Epia3Common.WriteTestLogFail(slogFilePath, "FindElementAndToggle failed:" + CheckBoxQualityStateId, Constants.TEST);
						TestCheck = ConstCommon.TEST_FAIL;
						sErrorMessage = "FindElementAndToggle failed:" + CheckBoxQualityStateId;
						return;
					}
				}
				catch (Exception ex)
				{
					TestCheck = ConstCommon.TEST_FAIL;
					Console.WriteLine("quality state checkbox :" + ex.Message + " --- " + ex.StackTrace);
					sErrorMessage = ex.Message + " --- " + ex.StackTrace;
				}
				#endregion
			
					string SaveID = "m_btnSave";
					// Find Save Button  element. 
					AutomationElement aeSaveBtn = AUIUtilities.FindElementByID(SaveID, aeNewCT);
					if (aeSaveBtn == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Save Button aeSaveBtn not found";
						Console.WriteLine("FindElementByID failed:" + SaveID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeSaveBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
					}
				
				#endregion

				#region  // Validation results

				AutomationElement aeGrid2 = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGrid2 == null)
				{
					sErrorMessage = "Find QualityReasonsDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("QualityReasonsDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));


				string[] headers = new string[5];
				headers[0] = "Description";
				headers[1] = "Valid for blue star";
				headers[2] = "Valid for gold star";
				headers[3] = "Valid for production state";
				headers[4] = "Valid for quality state";

				for (int i = 0; i < 5; i++)   // 
				{
					Thread.Sleep(3000);

					string cellname = headers[i] + " Row " + cellRow;
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCellv = AUIUtilities.FindElementByName(cellname, aeGrid2);

					if (aeCellv == null)
					{
						sErrorMessage = "Find QualityReasonsDataGridView aeCell failed:" + headers[i] + " at row " + cellRow;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell found at row: " + cellRow);
						Thread.Sleep(2000);

						ValuePattern vp = aeCellv.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						string Value = vp.Current.Value;
						Console.WriteLine("Get element.Current Value:" + Value);
					 
						if (i == 0 )
						{  
							if (!Value.Equals(cellValue))
							{
								sErrorMessage = "aeCell value should be:" + cellValue + " but not " + Value;
								Console.WriteLine(sErrorMessage);
								Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
								result = ConstCommon.TEST_FAIL;
								return;
							}
						}
						else
						{
							if (!Value.Equals("False"))
							{
								sErrorMessage = "aeCell value should be:" + cellValue + " but not " + Value;
								Console.WriteLine(sErrorMessage);
								Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
								result = ConstCommon.TEST_FAIL;
								return;
							}
							/*TogglePattern tg = aeCellv.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
							if (tg.Current.ToggleState == ToggleState.On)
							{
								sErrorMessage = "aeCell CHECK BOX SHOULD OFF:";
								Console.WriteLine(sErrorMessage);
								Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
								result = ConstCommon.TEST_FAIL;
								return;
							}
							 * */
						}
						
					}
				}

				result = ConstCommon.TEST_PASS;
				Console.WriteLine(testname + " ---pass --- ");
				Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);

				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region EditCarrierTypes
		public static void EditCarrierTypes(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			string PALLET_TYPE = "PalletType";
			AutomationElement aePanelLink = null;
			AutomationElement aeCell = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);
				#region // Open and Find Overview screen 
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, CONFIGURATION, CARRIER_TYPES, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Carrier Types Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIER_TYPES_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the CarrierTypesOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = CARRIER_TYPES + " Window not found";
					Console.WriteLine("FindElementByID failed:" + CARRIER_TYPES);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				#endregion

				AutomationElement aePalletType3 = null;
				int editRow = -1;
				#region // FInd CarrierType element of PalletType3 AND PalletType3 row position
				AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find CarrierTypeDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("CarrierTypeDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

				for (int i = 0; i < 10; i++)
				{
					// Construct the Grid Cell Element Name
					string cellname = sCarrierTypeID+" Row " + i;
					// Get the Element with the Row Col Coordinates
					aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Find CarrierTypeDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell CarrierTypeDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					// find cell value
					string CarrierTypeValue = string.Empty;
					try
					{
						ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						CarrierTypeValue = vp.Current.Value;
						Console.WriteLine("Get element.Current Value:" + CarrierTypeValue);
					}
					catch (System.NullReferenceException)
					{
						CarrierTypeValue = string.Empty;
					}

					if (CarrierTypeValue == null || CarrierTypeValue == string.Empty)
					{
						if (aePalletType3 == null)
						{
							sErrorMessage = "CarrierTypeDataGridView aeCell Value not found:" + cellname;
							Console.WriteLine("CarrierTypeDataGridView aeCell Value not found:" + cellname);
							Epia3Common.WriteTestLogMsg(slogFilePath, "CarrierTypeDataGridView cell value not found:" + cellname, Constants.TEST);
							result = ConstCommon.TEST_FAIL;
							return;
						}
					}

					if (CarrierTypeValue.Equals(PALLET_TYPE+"3"))
					{
						aePalletType3 = aeCell;
						editRow = i;
						Thread.Sleep(2000);
						Input.MoveToAndClick(aePalletType3);
						Thread.Sleep(2000);
						Console.WriteLine(testname + " ---cell found --- " + CarrierTypeValue);
						Epia3Common.WriteTestLogPass(slogFilePath, testname+" ---cell found --- "+CarrierTypeValue, Constants.TEST);
						break;
					}
				}
				#endregion

				Input.MoveToAndRightClick(aeCell.GetClickablePoint());



				#region // Find Edit... action

				Double X = aeCell.GetClickablePoint().X;
				Double Y = aeCell.GetClickablePoint().Y;
				
				Point EditBtnPoint = new Point(X + 45, Y+20);

				Input.MoveToAndClick(EditBtnPoint);
				#endregion

				#region  Edit CarrierTypes
				Thread.Sleep(2000);
				//find new Egemin Shell Window screen
				System.Windows.Automation.Condition c2x = new AndCondition(
				  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
				  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
				);

				// Find the Edit CarrierType Screen element.
				AutomationElement aeNewCT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2x);
				if (aeNewCT == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "New Carrier Type Window not found";
					Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				string origCTId = "";
				// "widthTextBox"
				if (AUIUtilities.FindTextBoxAndChangeValue("widthTextBox", aeNewCT, out origCTId, "1000", ref sErrorMessage))
					Thread.Sleep(3000);
				else
				{
					Console.WriteLine("FindTextBoxAndChangeValue failed:" + "widthTextBox");
					sErrorMessage = "FindTextBoxAndChangeValue failed:" + "widthTextBox";
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					return;
				}

				// "depthTextBox"
				if (AUIUtilities.FindTextBoxAndChangeValue("depthTextBox", aeNewCT, out origCTId, "1000", ref sErrorMessage))
					Thread.Sleep(3000);
				else
				{
					Console.WriteLine("FindTextBoxAndChangeValue failed:" + "depthTextBox");
					sErrorMessage = "FindTextBoxAndChangeValue failed:" + "depthTextBox";
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					return;
				}

				// "heightTextBox"
				if (AUIUtilities.FindTextBoxAndChangeValue("heightTextBox", aeNewCT, out origCTId, "1000", ref sErrorMessage))
					Thread.Sleep(3000);
				else
				{
					Console.WriteLine("FindTextBoxAndChangeValue failed:" + "heightTextBox");
					sErrorMessage = "FindTextBoxAndChangeValue failed:" + "heightTextBox";
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					return;
				}

				string SaveID = "m_btnSave";
				// Find Save Button  element. 
				AutomationElement aeSaveBtn = AUIUtilities.FindElementByID(SaveID, aeNewCT);
				if (aeSaveBtn == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Save Button aeSaveBtn not found";
					Console.WriteLine("FindElementByID failed:" + SaveID);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				else
				{
					Console.WriteLine("aeSaveBtn:");
					Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
					Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
				}
				#endregion

				#region  // Validation results         
				AutomationElement aeGridx = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGridx == null)
				{
					sErrorMessage = "Find CarrierTypeDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("CarrierTypeDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

				Thread.Sleep(3000);
				// Construct the Grid Cell Element Name
				string cellnameX = "Volume Row " + editRow;
				// Get the Element with the Row Col Coordinates
				AutomationElement aeCellX = AUIUtilities.FindElementByName(cellnameX, aeGridx);

				if (aeCellX == null)
				{
					sErrorMessage = "Find CarrierTypeDataGridView aeCell failed:" + cellnameX;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("cell CarrierTypeDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				// find cell value
				string CarrierSizeValue = string.Empty;
				try
				{
					ValuePattern vp = aeCellX.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
					CarrierSizeValue = vp.Current.Value;
					Console.WriteLine("Get element.Current Value:" + CarrierSizeValue);
				}
				catch (System.NullReferenceException)
				{
					CarrierSizeValue = string.Empty;
				}

				if (CarrierSizeValue == null || CarrierSizeValue == string.Empty)
				{
					sErrorMessage = "CarrierTypeDataGridView aeCell Value not found:" + cellnameX;
					Console.WriteLine("CarrierTypeDataGridView aeCell Value not found:" + cellnameX);
					Epia3Common.WriteTestLogMsg(slogFilePath, "CarrierTypeDataGridView cell value not found:" + cellnameX, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				if (CarrierSizeValue.Equals("1000 x 1000 x 1000 mm"))
				{
					result = ConstCommon.TEST_PASS;
					Console.WriteLine(testname + " ---pass --- " + CarrierSizeValue);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(testname + " ---fail --- " + CarrierSizeValue);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					return;
				}
			
				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region DeleteCarrierTypes
		public static void DeleteCarrierTypes(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			string PALLET_TYPE = "PalletType";
			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);
				#region // Open and Find Overview screen
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, CONFIGURATION, CARRIER_TYPES, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Carrier Types Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIER_TYPES_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the CarrierTypesOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = CARRIER_TYPES + " Window not found";
					Console.WriteLine("FindElementByID failed:" + CARRIER_TYPES);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				#endregion

				AutomationElement aePalletType3 = null;
				int deleteRow = -1;
				#region // FInd CarrierType element of PalletType3 AND PalletType3 row position and Move to and Click Element
				AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				AutomationElement aeCell = null;
				if (aeGrid == null)
				{
					sErrorMessage = "Find CarrierTypeDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("CarrierTypeDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

				for (int i = 0; i < 10; i++)
				{
					// Construct the Grid Cell Element Name
					string cellname = sCarrierTypeID+" Row " + i;
					// Get the Element with the Row Col Coordinates
					aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

					if (aeCell == null)
					{
						sErrorMessage = "Find CarrierTypeDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell CarrierTypeDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					// find cell value
					string CarrierTypeValue = string.Empty;
					try
					{
						ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						CarrierTypeValue = vp.Current.Value;
						Console.WriteLine("Get element.Current Value:" + CarrierTypeValue);
					}
					catch (System.NullReferenceException)
					{
						CarrierTypeValue = string.Empty;
					}

					if (CarrierTypeValue == null || CarrierTypeValue == string.Empty)
					{
						if (aePalletType3 == null)
						{
							sErrorMessage = "CarrierTypeDataGridView aeCell Value not found:" + cellname;
							Console.WriteLine("CarrierTypeDataGridView aeCell Value not found:" + cellname);
							Epia3Common.WriteTestLogMsg(slogFilePath, "CarrierTypeDataGridView cell value not found:" + cellname, Constants.TEST);
							result = ConstCommon.TEST_FAIL;
							return;
						}
					}

					if (CarrierTypeValue.Equals(PALLET_TYPE + "3"))
					{
						aePalletType3 = aeCell;
						deleteRow = i;
						Thread.Sleep(2000);
						Input.MoveToAndClick(aePalletType3);
						Thread.Sleep(2000);
						Console.WriteLine(testname + " ---cell found --- " + CarrierTypeValue);
						Epia3Common.WriteTestLogPass(slogFilePath, testname + " ---cell found --- " + CarrierTypeValue, Constants.TEST);
						break;
					}
				}
				#endregion

				#region // Find Delete... Button and Move to and click Delete... button
				Input.MoveToAndRightClick(aeCell.GetClickablePoint());
				Double X = aeCell.GetClickablePoint().X; 
				Double Y = aeCell.GetClickablePoint().Y;

				Point DeleteBtnPoint = new Point(X + 45, Y + 50);

				Input.MoveToAndClick(DeleteBtnPoint);
				#endregion

				#region  Delete CarrierTypes
				Thread.Sleep(2000);
				//find new Egemin Shell Window screen
				System.Windows.Automation.Condition c2x = new AndCondition(
				  new PropertyCondition(AutomationElement.NameProperty, "Delete Carrier Types?"),
				  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
				);

				// Find the Delete CarrierType Screen element.
				AutomationElement aeNewCT = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2x);
				if (aeNewCT == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Delete Carrier Types Window not found";
					Console.WriteLine("FindElementByID failed:" + "Delete Carrier Types?");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Check text include PalletType3 
				AutomationElement aeText = AUIUtilities.FindElementByID("m_LblInfo", aeNewCT);
				if (aeText == null)
				{
					sErrorMessage = "Find Delete CarrierType window Text field failed:";
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					string textMsg = aeText.Current.Name; ;
					//TextPattern vp = (TextPattern)aeText.GetCurrentPattern(TextPattern.Pattern);
					Thread.Sleep(1000);
					//string getText = vp.DocumentRange.GetText(-1).Trim();
					Thread.Sleep(2000);

					Console.WriteLine("Text is : " + textMsg);
					if (textMsg.IndexOf("") >= 0)
					{
						Console.WriteLine("Delete text message include : " + PALLET_TYPE+"3");
					}
					else
					{
						sErrorMessage = "Find Delete CarrierType window Text include no id:" + PALLET_TYPE + "3";
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath,sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
				}

				string yesID = "m_btn1";
				// Find Yes Button  element. 
				AutomationElement aeYesBtn = AUIUtilities.FindElementByID(yesID, aeNewCT);
				if (aeYesBtn == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Yes Button aeYesBtn not found";
					Console.WriteLine("FindElementByID failed:" + yesID);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				else
				{
					Console.WriteLine("aeYesBtn:");
					Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeYesBtn));
					Thread.Sleep(3000);
					Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeYesBtn));
				}
				#endregion

				#region  // Validation results

				AutomationElement aeGridx = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGridx == null)
				{
					sErrorMessage = "Find CarrierTypeDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("CarrierTypeDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

				Thread.Sleep(3000);
				// Construct the Grid Cell Element Name
				string cellnameX = sCarrierTypeID+" Row " + deleteRow;
				// Get the Element with the Row Col Coordinates
				AutomationElement aeCellX = AUIUtilities.FindElementByName(cellnameX, aeGridx);

				if (aeCellX == null)
				{
					result = ConstCommon.TEST_PASS;
					Console.WriteLine(testname + " ---pass ---  cell is null ");
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					Console.WriteLine("cell CarrierTypeDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					// find cell value
					string CarrierTypeValue = string.Empty;
					try
					{
						ValuePattern vp = aeCellX.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						CarrierTypeValue = vp.Current.Value;
						Console.WriteLine("Get element.Current Value:" + CarrierTypeValue);
					}
					catch (System.NullReferenceException)
					{
						CarrierTypeValue = string.Empty;
					}

					if (CarrierTypeValue.Equals(PALLET_TYPE+"3"))
					{
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine(testname + " ---fail --- " + CarrierTypeValue+" not deleted" );
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						
					}
					else
					{
						result = ConstCommon.TEST_PASS;
						Console.WriteLine(testname + " ---pass --- " + CarrierTypeValue);
						Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					}
				}
				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region Add10UnitsOn1Carrier
		public static void Add10UnitsOn1Carrier(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);
				#region // Open and Find Overview screen 
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, CARRIERS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Carriers Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIERS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the CarriersOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = CARRIERS + " Window not found";
					Console.WriteLine("FindElementByID failed:" + CARRIERS);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				#endregion

				#region  // Find Search Field and search P0000
				AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField);
				Thread.Sleep(3000);
				Input.MoveToAndClick(aeSearchField);
				Thread.Sleep(1000);
				ValuePattern vps = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
				Thread.Sleep(1000);
				string getValue = vps.Current.Value;
				Thread.Sleep(2000);
				vps.SetValue("P0000");
				Thread.Sleep(1000);

				AutomationElement aeP0000 = null;
				 // FInd CarrierID element of P0000 AND P00000 row position
				AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find CarrierTypeDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("CarrierTypeDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				
				// Find Cell aeP00000
				string CellValue = "P0000";
				int CellRow = -1;
				aeP0000 = FindCellFromGrid(aeGrid, sCarrierIDHeader, CellValue, ref CellRow);

				if (aeP0000 == null)
				{
					sErrorMessage = "Find Carriers  aeCell failed:" + CellValue;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("found row at: " + CellRow);
					Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeP0000));

				}
				#endregion

				#region // Add Units
				Thread.Sleep(2000);
				//find new Egemin Shell Window screen
				System.Windows.Automation.Condition c2 = new AndCondition(
				  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
				  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
				);

				// Find the  Edit Carrier Screen element.
				AutomationElement aeEditC = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
				if (aeEditC == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Edit Carrier Window not found";
					Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find Unit Add button
				string addButtinId = "m_BtnUnitAdd";
				AutomationElement aeUnitAdd = AUIUtilities.FindElementByID(addButtinId, aeEditC);
				if (aeUnitAdd == null)
				{
					sErrorMessage = "Find Unit Add button failed:" + addButtinId;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				Thread.Sleep(2000);

				for (int i = 0; i < 10; i++)
				{
					Thread.Sleep(2000);
					Input.MoveToAndClick(aeUnitAdd);
					Thread.Sleep(2000);

					// Find the  Unit Screen.
					//find new Egemin Shell Window screen
					System.Windows.Automation.Condition cu = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					AutomationElement aeUnitScreen = aeEditC.FindFirst(TreeScope.Element | TreeScope.Descendants, cu);
					if (aeUnitScreen == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Unit Screen Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}

					//-------------------------------
					 // Search unit
					 // Find Search Field
					AutomationElement aeSearchField1 = AUIUtilities.FindElementByType(ControlType.Edit, aeUnitScreen);
					if (aeSearchField1 == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
						Console.WriteLine(sErrorMessage);
						return;
					}
					else
					{
						Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					Input.MoveTo(aeSearchField1);
					string searchValue = "NFPM1_20090403-00"+i+"A";
					Thread.Sleep(3000);

					Input.MoveToAndClick(aeSearchField1);
					Thread.Sleep(1000);
					ValuePattern vp = (ValuePattern)aeSearchField1.GetCurrentPattern(ValuePattern.Pattern);
					Thread.Sleep(1000);
					vp.SetValue(searchValue);
					Thread.Sleep(1000);

					AutomationElement aeSearchGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeUnitScreen);
					if (aeSearchGrid == null)
					{   
						sErrorMessage = "Find UnitSearchGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
						Console.WriteLine("UnitSearchGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

					Thread.Sleep(2000);
					Console.WriteLine("search unit: " + searchValue);
					int row = 0;
					AutomationElement aeUnit = FindCellFromGridAtHeaderRow(aeSearchGrid, sUnitIDHeader, row);
				 
					if (aeUnit == null)
					{
						sErrorMessage = "Find Unit  aeCell failed:" + searchValue;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;

						#region // Cancel screen
						string CancelId = "m_btnCancel";
						// Find Cancel Button  element. 
						AutomationElement aeCancelBtn = AUIUtilities.FindElementByID(CancelId, aeUnitScreen);
						if (aeCancelBtn == null)
						{
							sErrorMessage = sErrorMessage + "Cancel Button aeCancelBtn not found";
							Console.WriteLine("FindElementByID failed:" + CancelId);
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						}
						else
						{
							Console.WriteLine(i + ": found aeCancelBtn:");
							Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn));
							Thread.Sleep(2000);
							Input.MoveToAndClick(aeCancelBtn);
							Thread.Sleep(2000);
							Input.MoveToAndClick(aeCancelBtn);
							Thread.Sleep(2000);

							AutomationElement aeCancelBtn2 = AUIUtilities.FindElementByID(CancelId, aeEditC);
							if (aeCancelBtn2 == null)
							{
								sErrorMessage = sErrorMessage + "Cancel Button aeCancelBtn2 not found";
								Console.WriteLine("FindElementByID failed 2 :" + CancelId);
								Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
							}
							else
							{
								Console.WriteLine(i + ": found aeCancelBtn2:");
								Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn2));
								Thread.Sleep(2000);
								Input.MoveToAndClick(aeCancelBtn2);
								Thread.Sleep(2000);
							}
						}
						#endregion

						return;
					}
					else
					{
						// Check thsi UnitID equal to searchValue
						ValuePattern vpUnit = aeUnit.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
						string UnitValue = vpUnit.Current.Value;
						if (UnitValue.Equals(searchValue))
						{
							Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeUnit));
					   
							Console.WriteLine("Unit "+searchValue +"  added");
							Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						}
						else
						{

							result = ConstCommon.TEST_FAIL;
							sErrorMessage = searchValue + " is not found during Searching, but:" + UnitValue;
							Console.WriteLine(sErrorMessage);
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);

							#region // Cancel screen
							string CancelId = "m_btnCancel";
							// Find Cancel Button  element. 
							AutomationElement aeCancelBtn = AUIUtilities.FindElementByID(CancelId, aeUnitScreen);
							if (aeCancelBtn == null)
							{
								sErrorMessage = sErrorMessage + "Cancel Button aeCancelBtn not found";
								Console.WriteLine("FindElementByID failed:" + CancelId);
								Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
							}
							else
							{
								Console.WriteLine(i + ": found aeCancelBtn:");
								Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn));
								Thread.Sleep(2000);
								Input.MoveToAndClick(aeCancelBtn);
								Thread.Sleep(2000);
								Input.MoveToAndClick(aeCancelBtn);
								Thread.Sleep(2000);

								AutomationElement aeCancelBtn2 = AUIUtilities.FindElementByID(CancelId, aeEditC);
								if (aeCancelBtn2 == null)
								{
									sErrorMessage = sErrorMessage + "Cancel Button aeCancelBtn2 not found";
									Console.WriteLine("FindElementByID failed 2 :" + CancelId);
									Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
								}
								else
								{
									Console.WriteLine(i + ": found aeCancelBtn2:");
									Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn2));
									Thread.Sleep(2000);
									Input.MoveToAndClick(aeCancelBtn2);
									Thread.Sleep(2000);
								}
							}
							#endregion
							
							return;
						}
					}
					Thread.Sleep(3000);
				}

				Thread.Sleep(2000);
				// Save All Units
				string SaveID = "m_btnSave";
				// Find Save Button  element. 
				AutomationElement aeSaveBtn = AUIUtilities.FindElementByID(SaveID, aeEditC);
				if (aeSaveBtn == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Save Button aeSaveBtn not found";
					Console.WriteLine("FindElementByID failed:" + SaveID);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				else
				{
					Console.WriteLine("aeSaveBtn:");
					Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
					Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
				}

				#endregion

				Thread.Sleep(5000);

				#region // Validate Results
				// find Error screen
				string ErrorScreenID = "ErrorScreen";
				AutomationElement aeErrorScreen = AUIUtilities.FindElementByID(ErrorScreenID, root);
				if (aeErrorScreen != null)
				{
					string ErrorTextID = "m_LblCaption";
					Thread.Sleep(2000);
					// Find  error text  element. 
					AutomationElement aeErrorText = AUIUtilities.FindElementByID(ErrorTextID, aeErrorScreen);
					if (aeErrorText == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "error text aeErrorText not found";
						Console.WriteLine("FindElementByID failed:" + ErrorTextID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						sErrorMessage = aeErrorText.Current.Name;
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindElementByID failed:" + SaveID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						
						// close Error Screen
						string CloseBtnID = "m_BtnClose";
						Thread.Sleep(2000);
						// Find  close btn element. 
						AutomationElement aeCloseBtn = AUIUtilities.FindElementByID(CloseBtnID, aeErrorScreen);
						if (aeCloseBtn == null)
						{
							result = ConstCommon.TEST_FAIL;
							sErrorMessage = "Close button aeCloseBtn not found";
							Console.WriteLine("FindElementByID failed:" + CloseBtnID);
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
							return;
						}
						Console.WriteLine("aeCloseBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCloseBtn));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCloseBtn));
					}

					Thread.Sleep(2000);

				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = " No  Error screen window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}


				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region LocationConfirmManualChange
		public static void LocationConfirmManualChange(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;
		   
			string locationID = "AT48.96.07";
			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);
				#region Find and open carrier overview screen and click New button
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, CARRIERS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Carriers Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIERS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the CarriersOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = REELS_OVERVIEW_TITLE + " Window not found";
					Console.WriteLine("FindElementByID failed:" + REELS_OVERVIEW_TITLE);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find New... Button  //wwwwwwwwwwwwwwwwwww
				// Find tool bar 
				AutomationElement aeToolBar = AUIUtilities.FindElementByType(ControlType.ToolBar, aeOverview);
				if (aeToolBar == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "ToolBar " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("ToolBar " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				// Find Edit textbox 
				AutomationElement aeEdit = AUIUtilities.FindElementByType(ControlType.Edit, aeToolBar);
				if (aeEdit == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "EditTextBox " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("EditTextBox " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Double Bottom = aeEdit.Current.BoundingRectangle.Bottom;
				Double Left = aeEdit.Current.BoundingRectangle.Left;
				Double Right = aeEdit.Current.BoundingRectangle.Right;
				Double Width = aeEdit.Current.BoundingRectangle.Width;
				Double Height = aeEdit.Current.BoundingRectangle.Height;
				Double X = aeEdit.Current.BoundingRectangle.X;
				Double Y = aeEdit.Current.BoundingRectangle.Y;
				Double Top = aeEdit.Current.BoundingRectangle.Top;

				double xclick = Right + 80;
				Point NewBtnPoint = new Point(xclick, (Bottom + Top) / 2);

				Input.MoveTo(NewBtnPoint);

				#endregion

				#region // Add New Carriers
				for (int i = 0; i < 1; i++)
				{
					Thread.Sleep(3000);
					Input.MoveToAndClick(NewBtnPoint);
					Thread.Sleep(3000);

					//find new Egemin Shell Window screen
					System.Windows.Automation.Condition c2p = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					// Find the New ReelInput Screen element.
					AutomationElement aeNewCarrierInput = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2p);
					if (aeNewCarrierInput == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "New Reels Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}

					// select Carrier Type
					AutomationElement aeCombo = AUIUtilities.FindElementByID("m_ComboCarrierType", aeNewCarrierInput);
					if (aeCombo == null)
					{
						Console.WriteLine("failed to find CarrierType aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						sErrorMessage = "failed to find Carriertype aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss");
						TestCheck = ConstCommon.TEST_FAIL;
						sEventEnd = true;
						return;
					}

					SelectionPattern selectPattern =
					   aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

					Thread.Sleep(1000);
					
					AutomationElement item
						= AUIUtilities.FindElementByName( "RackType", aeCombo);
					if (item != null)
					{
						Console.WriteLine("RackType item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						Thread.Sleep(2000);

						SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
						itemPattern.Select();
					}
					else
					{
						Console.WriteLine("Finding CarrierType combo box item RackType failed");
						sErrorMessage = "Finding CarrierType combo box item RackType failed";
						TestCheck = ConstCommon.TEST_FAIL;
						sEventEnd = true;
						return;
					}

					Thread.Sleep(3000);
					string CarrierAutoID = "m_TxtCarrierId";
					// "m_TxtCarrierId"   
					// "m_TxtUnitId"  This is Edit Control, should use setFocus + sendKeys
					if (AUIUtilities.FindDocumentAndSendText(CarrierAutoID, aeNewCarrierInput, "R0003", ref sErrorMessage))
						Thread.Sleep(3000);
					else
					{
						MessageBox.Show(sErrorMessage);
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine("FindTextBoxAndChangeValue failed:" + CarrierAutoID);
						sErrorMessage = "FindTextBoxAndChangeValue failed:" + CarrierAutoID;
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				   
					//----------------------------------------------------------------
					//  "m_BtnSearchLocation"
					string BtnSearchLocationID = "m_BtnSearchLocation";
					// Find Save Button  element. 
					AutomationElement aeSearchLocationBtn = AUIUtilities.FindElementByID(BtnSearchLocationID, aeNewCarrierInput);
					if (aeSearchLocationBtn == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Search Location Button aeSearchLocationBtn not found";
						Console.WriteLine("FindElementByID failed:" + BtnSearchLocationID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeSearchLocationBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSearchLocationBtn));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSearchLocationBtn));
					}
					// This is visiable Screen
					//find new Egemin Shell Window screen
					System.Windows.Automation.Condition c2s = new AndCondition(
					  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
					  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
					);

					// Find the New LocationID Screen element. ps not from root should be 
					AutomationElement aeNewLocSearch = aeNewCarrierInput.FindFirst(TreeScope.Element | TreeScope.Descendants, c2s);
					if (aeNewLocSearch == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "New Location Search Window not found";
						Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeNewLocSearch Screen : Now find Grid ");
					}

					Thread.Sleep(3000);


					// input AT48.96.07 in search field
					// Find Search Field
					AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeNewLocSearch);
					if (aeSearchField == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
						Console.WriteLine(sErrorMessage);
						return;
					}
					else
					{
						Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}
					Thread.Sleep(2000);

					Input.MoveTo(aeSearchField);
					Thread.Sleep(2000);

					#region  // Search and click on this Location
					
					Input.MoveToAndClick(aeSearchField);
					Thread.Sleep(1000);
					ValuePattern vp = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
					Thread.Sleep(1000);
					string getValue = vp.Current.Value;
					Thread.Sleep(2000);
					vp.SetValue(locationID);

					AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeNewLocSearch);
					if (aeGrid == null)
					{
						sErrorMessage = "Find LocationsDataGrid failed:" + KC_GRIDDATA_ID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
						Console.WriteLine("LocationsDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

					Thread.Sleep(3000);
					// Construct the Grid Cell Element Name
					string cellname = sLocationIDHeader+" Row " + 0;
					// Get the Element with the Row Col Coordinates
					AutomationElement aeCell = FindCellFromGridAtHeaderRow(aeGrid, "Location id", 0);
	

					if (aeCell == null)
					{
						sErrorMessage = "Find LocationsDataGridView aeCell failed:" + cellname;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}
					else
					{
						Console.WriteLine("cell LocationsDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					}

					Input.MoveToAndClick(aeCell);
					
					#endregion
					
					Thread.Sleep(2000);

					string OKID2 = "m_btnSave";
					// Find Save Button  element. 
					AutomationElement aeOK2 = AUIUtilities.FindElementByID(OKID2, aeNewLocSearch);
					if (aeOK2 == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "OK2 Button aeOK2 not found";
						Console.WriteLine("FindElementByID failed:" + OKID2);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeOK2:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeOK2));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeOK2));
					}

					Thread.Sleep(5000);

					Console.WriteLine("Save............"  );
					Thread.Sleep(5000);

					string SaveID = "m_btnSave";
					// Find Save Button  element. 
					AutomationElement aeSave = AUIUtilities.FindElementByID(SaveID, aeNewCarrierInput);
					if (aeSave == null)
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = "Save Button aeSave not found";
						Console.WriteLine("FindElementByID failed:" + SaveID);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						return;
					}
					else
					{
						Console.WriteLine("aeSave:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSave));
						Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSave));
					}

					Thread.Sleep(8000);
				}
				#endregion
				Thread.Sleep(5000);
				//----------------------------------------------------------------
				#region // Validation results
				AUICommon.ClearDisplayedScreens(root, 2);
				#region Find and open Carriers overview screen and Searching field and input LocationID
				
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, STORAGE_LOCATIONS , ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Location Window
				System.Windows.Automation.Condition c2 = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, STORAGE_LOCATIONS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the StorageOverview element.
				AutomationElement aeStorageLocOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
				if (aeStorageLocOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = STORAGE_LOCATIONS_OVERVIEW_TITLE + " Window not found";
					Console.WriteLine("FindElementByID failed:" + STORAGE_LOCATIONS_OVERVIEW_TITLE);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				 // Find Search Field
				AutomationElement aeLocSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeStorageLocOverview);
				if (aeLocSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeLocSearchField);
				Thread.Sleep(3000);
				Input.MoveToAndClick(aeLocSearchField);
				Thread.Sleep(1000);
				ValuePattern vps = (ValuePattern)aeLocSearchField.GetCurrentPattern(ValuePattern.Pattern);
				Thread.Sleep(1000);
				string getLocValue = vps.Current.Value;
				Thread.Sleep(2000);
				vps.SetValue(locationID);
				Thread.Sleep(1000);

				Thread.Sleep(3000);
				// 
				AutomationElement aeSearchLocationGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeStorageLocOverview);
				if ( aeSearchLocationGrid == null)
				{
					sErrorMessage = "Find LocationDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("CarrierDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Thread.Sleep(3000);

				// find cell LocationID
				string header = sLocationIDHeader;
				string cellValue = locationID;
				int cellRow = 0;
				AutomationElement ae9607Cell = FindCellFromGrid(aeSearchLocationGrid, header, cellValue, ref cellRow);

				if (ae9607Cell == null)
				{
					sErrorMessage = "Find LocationDataGridView aeCell failed:" + cellValue;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("cell LocationDataGridView found: " + cellValue);
					Thread.Sleep(1000);
					Input.MoveToAndDoubleClick( AUIUtilities.GetElementCenterPoint(ae9607Cell)      );
				}

				Thread.Sleep(3000);
				#endregion

				// Check State
				Thread.Sleep(2000);
				//find new Egemin Shell Window screen
				System.Windows.Automation.Condition c2x = new AndCondition(
				  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
				  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
				);

				// Find the  Edit Location Screen element.
				AutomationElement aeEditL2 = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2x);
				if (aeEditL2 == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Edit Locations Window not found";
					Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find State Combo
				string stateId = "m_ComboState";
				AutomationElement aeComboState = AUIUtilities.FindElementByID(stateId, aeEditL2);
				if (aeComboState == null)
				{
					sErrorMessage = "Find State Combo failed:" + stateId;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				bool confirmSelection = false;

				SelectionPattern selectPatternLoc =
				   aeComboState.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

				AutomationElement itemLoc
					= AUIUtilities.FindElementByName("Confirm manual change", aeComboState);
				if (itemLoc != null)
				{
					Console.WriteLine("ConfirmManualChange item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
					Thread.Sleep(2000);

					SelectionItemPattern itemPattern = itemLoc.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;

					confirmSelection = itemPattern.Current.IsSelected;
				}
				else
				{
					Console.WriteLine("Finding item ConfirmManualChange failed");
					sErrorMessage = "Finding item ConfirmManualChange failed";
					TestCheck = ConstCommon.TEST_FAIL;
					ClickCancelButton("m_btnCancel", aeEditL2);
					return;
				}

				ClickCancelButton("m_btnCancel", aeEditL2);
			   
				if (confirmSelection)
				{
					result = ConstCommon.TEST_PASS;
					string ms = " OK " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);

				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					string ms = " OK " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogFail(slogFilePath, testname, Constants.TEST);
				}

				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion

		#region CarrierConfirmManualChange
		public static void CarrierConfirmManualChange(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root, 2);
				#region Find and open Carriers overview screen and Searching field and input C100
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INVENTORY, CARRIERS, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Carriers Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, CARRIERS_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the CarriersOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = REELS_OVERVIEW_TITLE + " Window not found";
					Console.WriteLine("FindElementByID failed:" + REELS_OVERVIEW_TITLE);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				 // Find Search Field
				AutomationElement aeSearchField = AUIUtilities.FindElementByType(ControlType.Edit, aeOverview);
				if (aeSearchField == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField);
				Thread.Sleep(3000);
				Input.MoveToAndClick(aeSearchField);
				Thread.Sleep(1000);
				ValuePattern vps = (ValuePattern)aeSearchField.GetCurrentPattern(ValuePattern.Pattern);
				Thread.Sleep(1000);
				string getValue = vps.Current.Value;
				Thread.Sleep(2000);
				vps.SetValue("C100");
				Thread.Sleep(1000);

				Thread.Sleep(3000);
				//
				AutomationElement aeSearchCarrierGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if ( aeSearchCarrierGrid == null)
				{
					sErrorMessage = "Find CarrierDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("CarrierDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Thread.Sleep(3000);

				// find cell C1001
				string header = sCarrierIDHeader;
				string cellValue = "C1001";
				int cellRow = -1;
				AutomationElement aeP1001Cell = FindCellFromGrid(aeSearchCarrierGrid, header, cellValue, ref cellRow);

				if (aeP1001Cell == null)
				{
					sErrorMessage = "Find CarrierTypeDataGridView aeCell failed:" + cellValue;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("cell CarrierTypeDataGridView found: " + cellValue);
					Thread.Sleep(1000);
					Input.MoveToAndDoubleClick( AUIUtilities.GetElementCenterPoint(aeP1001Cell)      );
				}

				Thread.Sleep(3000);
				#endregion

				#region // Add Units 
				Thread.Sleep(2000);
				//find new Egemin Shell Window screen
				System.Windows.Automation.Condition c2 = new AndCondition(
				  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
				  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
				);

				// Find the  Edit Carrier Screen element.
				AutomationElement aeEditC = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
				if (aeEditC == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Edit Carrier Window not found";
					Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find Unit Add button
				string addButtinId = "m_BtnUnitAdd";
				AutomationElement aeUnitAdd = AUIUtilities.FindElementByID(addButtinId, aeEditC);
				if (aeUnitAdd == null)
				{
					sErrorMessage = "Find Unit Add button failed:" + addButtinId;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				Thread.Sleep(2000);

				Thread.Sleep(2000);
				Input.MoveToAndClick(aeUnitAdd);
				Thread.Sleep(2000);

				// Find the  Unit Screen.
				//find new Egemin Shell Window screen
				System.Windows.Automation.Condition cu = new AndCondition(
				  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
				  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
				);

				AutomationElement aeUnitScreen = aeEditC.FindFirst(TreeScope.Element | TreeScope.Descendants, cu);
				if (aeUnitScreen == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Unit Screen Window not found";
					Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
			   
					
				// Search unit
				// Find Search Field
				AutomationElement aeSearchField1 = AUIUtilities.FindElementByType(ControlType.Edit, aeUnitScreen);
				if (aeSearchField1 == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Failed find " + "aeSearchField " + " at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(sErrorMessage);
					return;
				}
				else
				{
					Console.WriteLine("aeSearchField " + ": found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Input.MoveTo(aeSearchField1);
				string searchValue = "C1000";
				Thread.Sleep(3000);

				Input.MoveToAndClick(aeSearchField1);
				Thread.Sleep(1000);
				ValuePattern vp = (ValuePattern)aeSearchField1.GetCurrentPattern(ValuePattern.Pattern);
				Thread.Sleep(1000);
				vp.SetValue(searchValue);
				Thread.Sleep(1000);

				AutomationElement aeSearchGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeUnitScreen);
				if (aeSearchGrid == null)
				{
					sErrorMessage = "Find UnitSearchGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("UnitSearchGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

				Thread.Sleep(2000);
				Console.WriteLine("search unit: " + searchValue);
				int row = 0;
				AutomationElement aeUnit = FindCellFromGridAtHeaderRow(aeSearchGrid, sCarrierIDHeader, row);

				if (aeUnit == null)
				{
					sErrorMessage = "Find Unit  aeCell failed:" + searchValue;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					#region // Cancel screen
					string CancelId = "m_btnCancel";
					// Find Cancel Button  element. 
					AutomationElement aeCancelBtn = AUIUtilities.FindElementByID(CancelId, aeUnitScreen);
					if (aeCancelBtn == null)
					{
						sErrorMessage = sErrorMessage + "Cancel Button aeCancelBtn not found";
						Console.WriteLine("FindElementByID failed:" + CancelId);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					}
					else
					{
						Console.WriteLine("found aeCancelBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn));
						Thread.Sleep(2000);
						Input.MoveToAndClick(aeCancelBtn);
						Thread.Sleep(2000);
						Input.MoveToAndClick(aeCancelBtn);
						Thread.Sleep(2000);

						AutomationElement aeCancelBtn2 = AUIUtilities.FindElementByID(CancelId, aeEditC);
						if (aeCancelBtn2 == null)
						{
							sErrorMessage = sErrorMessage + "Cancel Button aeCancelBtn2 not found";
							Console.WriteLine("FindElementByID failed 2 :" + CancelId);
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						}
						else
						{
							Console.WriteLine(" found aeCancelBtn2:");
							Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn2));
							Thread.Sleep(2000);
							Input.MoveToAndClick(aeCancelBtn2);
							Thread.Sleep(2000);
						}
					}
					#endregion
					return;
				}
				else
				{
					// Check thsi UnitID equal to searchValue
					ValuePattern vpUnit = aeUnit.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
					string UnitValue = vpUnit.Current.Value;
					if (UnitValue.Equals(searchValue))
					{
						Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeUnit));

						Console.WriteLine("Unit " + searchValue + "  added");
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					}
					else
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = searchValue + " is not found during Searching, but:" + UnitValue;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						#region // Cancel screen
						string CancelId = "m_btnCancel";
						// Find Cancel Button  element. 
						AutomationElement aeCancelBtn = AUIUtilities.FindElementByID(CancelId, aeUnitScreen);
						if (aeCancelBtn == null)
						{
							sErrorMessage = sErrorMessage + "Cancel Button aeCancelBtn not found";
							Console.WriteLine("FindElementByID failed:" + CancelId);
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
						}
						else
						{
							Console.WriteLine(" found aeCancelBtn:");
							Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn));
							Thread.Sleep(2000);
							Input.MoveToAndClick(aeCancelBtn);
							Thread.Sleep(2000);
							Input.MoveToAndClick(aeCancelBtn);
							Thread.Sleep(2000);

							AutomationElement aeCancelBtn2 = AUIUtilities.FindElementByID(CancelId, aeEditC);
							if (aeCancelBtn2 == null)
							{
								sErrorMessage = sErrorMessage + "Cancel Button aeCancelBtn2 not found";
								Console.WriteLine("FindElementByID failed 2 :" + CancelId);
								Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
							}
							else
							{
								Console.WriteLine( "found aeCancelBtn2:");
								Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn2));
								Thread.Sleep(2000);
								Input.MoveToAndClick(aeCancelBtn2);
								Thread.Sleep(2000);
							}
						}
						#endregion
						return;
					}
				}
				Thread.Sleep(3000);
				
				// Save All Units
				string SaveID = "m_btnSave";
				// Find Save Button  element. 
				AutomationElement aeSaveBtn = AUIUtilities.FindElementByID(SaveID, aeEditC);
				if (aeSaveBtn == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Save Button aeSaveBtn not found";
					Console.WriteLine("FindElementByID failed:" + SaveID);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				else
				{
					Console.WriteLine("aeSaveBtn:");
					Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
					Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSaveBtn));
				}

				#endregion

				Thread.Sleep(5000);
				//----------------------------------------------------------------
				#region // Validation results

				//
				AutomationElement aeGrid = AUIUtilities.FindElementByID(KC_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find CarrierDataGrid failed:" + KC_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("CarrierDataGrid found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Thread.Sleep(3000);

				// find cell C1000
				string cellValueP1000 = "C1000";
				int cellRow1 = -1;
				AutomationElement aeP1000Cell = FindCellFromGrid(aeGrid, header, cellValueP1000, ref cellRow1);

				if (aeP1000Cell == null)
				{
					sErrorMessage = "Find CarrierTypeDataGridView aeCell failed:" + cellValueP1000;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("cell CarrierDataGridView found: " + cellValueP1000);
					Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeP1000Cell));
					Thread.Sleep(3000);

				}
				// Check State
				Thread.Sleep(2000);
				//find new Egemin Shell Window screen
				System.Windows.Automation.Condition c2x = new AndCondition(
				  new PropertyCondition(AutomationElement.NameProperty, "Egemin Shell"),
				  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
				);

				// Find the  Edit Carrier Screen element.
				AutomationElement aeEditC2 = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2x);
				if (aeEditC2 == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "Edit Carrier Window not found";
					Console.WriteLine("FindElementByID failed:" + "Egemin Shell");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find State Combo
				string stateId = "m_ComboCarrierState";
				AutomationElement aeComboState = AUIUtilities.FindElementByID(stateId, aeEditC2);
				if (aeComboState == null)
				{
					sErrorMessage = "Find State Combo failed:" + stateId;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				bool confirmSelection = false;

					SelectionPattern selectPattern =
					   aeComboState.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

					AutomationElement item
						= AUIUtilities.FindElementByName("Confirm manual change", aeComboState);
					if (item != null)
					{
						Console.WriteLine("LanguageSettings item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
						Thread.Sleep(2000);

						SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;

						confirmSelection = itemPattern.Current.IsSelected;
					}
					else
					{
						Console.WriteLine("Finding item ConfirmManualChange failed");
						sErrorMessage = "Finding item ConfirmManualChange failed";
						TestCheck = ConstCommon.TEST_FAIL;
						return;
					}

					#region // Cancel Scrren
					string CancelIdx = "m_btnCancel";
					// Find Cancel Button  element. 
					AutomationElement aeCancelBtnx = AUIUtilities.FindElementByID(CancelIdx, aeEditC2);
					if (aeCancelBtnx == null)
					{
						sErrorMessage = sErrorMessage + "Cancel Button aeCancelBtn not found";
						Console.WriteLine("FindElementByID failed:" + CancelIdx);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					}
					else
					{
						Console.WriteLine("found aeCancelBtn:");
						Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtnx));
						Thread.Sleep(2000);
						Input.MoveToAndClick(aeCancelBtnx);
						Thread.Sleep(2000);
					}
					#endregion

					if (confirmSelection)
					{
						result = ConstCommon.TEST_PASS;
						string ms = " OK " + System.DateTime.Now.ToString("HH:mm:ss");
						Console.WriteLine(ms);
						Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);

					}
					else
					{
						result = ConstCommon.TEST_FAIL;
						string ms = " OK " + System.DateTime.Now.ToString("HH:mm:ss");
						Console.WriteLine(ms);
						Epia3Common.WriteTestLogFail(slogFilePath, testname, Constants.TEST);
					
					
					}

				
			   
				

				#endregion
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);

			}
		}
		#endregion
		
		#region SystemOverviewDisplay
		public static void SystemOverviewDisplay(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, SYSTEM, SYSTEM_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
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

				// Find System Overview Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, SYSTEM_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the SystemOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = SYSTEM_OVERVIEW + " Window not found";
					Console.WriteLine("FindElementByID failed:" + SYSTEM_OVERVIEW);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = SYSTEM_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}


		}
		#endregion SystemOverviewDisplay
	   
		#region AgvOverviewDisplay
		public static void AgvOverviewDisplay(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = AGV_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}


		}
		#endregion AgvOverviewDisplay
	   
		#region LocationOverviewDisplay
		public static void LocationOverviewDisplay(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, LOCATION_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Input.MoveToAndClick(aePanelLink);

				Thread.Sleep(10000);

				// Find Location Overview Window
				System.Windows.Automation.Condition c = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, LOCATION_OVERVIEW_TITLE),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the LocationOverview element.
				AutomationElement aeOverview = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
				if (aeOverview == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = LOCATION_OVERVIEW + " Window not found";
					Console.WriteLine("FindElementByID failed:" + "Locations");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = LOCATION_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}


		}
		#endregion LocationOverviewDisplay
	   
		#region TransportOverviewDisplay
		public static void TransportOverviewDisplay(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					//Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					//result = ConstCommon.TEST_FAIL;
					//return;
				}
				//else
				//    Input.MoveToAndClick(aeNode);

				if (aeNode == null)
				{
					Console.WriteLine("Node not exist:" + TRANSPORT_OVERVIEW);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = TRANSPORT_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}


		}
		#endregion TransportOverviewDisplay
	   
		#region AgvOverviewOpenDetail
		public static void AgvOverviewOpenDetail(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find AGV GridView
				AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find AgvDataGrid failed:" + AGV_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, testname + ": " + ex.Message + "---" + ex.StackTrace, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Check AGV text value
				string textID = "m_IdValueLabel";
				AutomationElement aeAgvText = AUIUtilities.FindElementByID(textID, aeDetailScreen);
				if (aeAgvText == null)
				{
					sErrorMessage = "Find AgvTextElement failed:" + textID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				string agvTextValue = aeAgvText.Current.Name;
				if (agvTextValue.Equals(AgvValue))
				{
					result = ConstCommon.TEST_PASS;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = detailScreenName + "Agv Value should be " + AgvValue + ", but  " + agvTextValue;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion
	   
		#region LocationOverviewOpenDetail
		public static void LocationOverviewOpenDetail(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, LOCATION_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				// Find Location GridView
				AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
				if (aeGrid == null)
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine("Find LocationDataGridView failed:" + "LocationDataGridView");
					Epia3Common.WriteTestLogFail(slogFilePath, "Find LocationDataGridView failed:" + "LocationDataGridView", Constants.TEST);
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
						Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find LocationDataGridView cell failed:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "---" + ex.StackTrace, Constants.TEST);
					return;
				}

				// Open Detail screen
				Input.MoveToAndDoubleClick(point);
				Thread.Sleep(2000);

				// wait extra one minute if project is TestProject.zip
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "End id: " + cellValue, Constants.TEST);

					string locScreenName = "Location detail - " + cellValue;
					Console.WriteLine("locScreenName: " + locScreenName);
					Epia3Common.WriteTestLogMsg(slogFilePath, "locScreenName: " + locScreenName, Constants.TEST);

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
						Epia3Common.WriteTestLogFail(slogFilePath, "Find LocationDetailView failed:" + "LocationDetailView", Constants.TEST);
						return;
					}

					// Check Location Value
					string textID = "m_IdValueLabel";
					AutomationElement aeLocText = AUIUtilities.FindElementByID(textID, aeDetailScreen);
					if (aeLocText == null)
					{
						sErrorMessage = "Find locTextElement failed:" + textID;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					string locTextValue = aeLocText.Current.Name;
					if (locTextValue.Equals(cellValue))
					{
						result = ConstCommon.TEST_PASS;
						Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					}
					else
					{
						result = ConstCommon.TEST_FAIL;
						sErrorMessage = locScreenName + "Loc Value should be " + cellValue + ", but  " + locTextValue;
						Console.WriteLine(sErrorMessage);
						Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					}
				}
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion

		#region LocationModeManual
		public static void LocationModeManual(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, LOCATION_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				AutomationElement aeGrid = AUIUtilities.FindElementByID(DATAGRIDVIEW_ID, aeOverview);
				if (aeGrid == null)
				{
					Console.WriteLine("Find LocationDataGridView failed:" + DATAGRIDVIEW_ID);
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView failed:" + DATAGRIDVIEW_ID, Constants.TEST);
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
						Epia3Common.WriteTestLogPass(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find LocationDataGridView cell failed:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "LocDataGridView cell value not found:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "---" + ex.StackTrace, Constants.TEST);
					return;
				}

				string ModeValue = AUICommon.GetDataGridViewCellValueAt(2, "Mode", aeGrid);
				if (ModeValue == null || ModeValue == string.Empty)
				{
					sErrorMessage = "LocDataGridView aeCell Mode Value not found:" + "Mode Row 2";
					Console.WriteLine("LocDataGridView aeCell Mode Value not found:" + "Mode Row 2");
					Epia3Common.WriteTestLogMsg(slogFilePath, "LocDataGridView cell Mode value not found:" + "Mode Row 2", Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				Input.MoveToAndRightClick(point);

				System.Windows.Point ModePoint = new Point(point.X + 20, point.Y + 70);
				Thread.Sleep(2000);
				Input.MoveTo(ModePoint);
				Console.WriteLine("move modepoint");

				Thread.Sleep(2000);
				System.Windows.Point ManPoint = new Point(ModePoint.X + 180, ModePoint.Y + 50);
				Thread.Sleep(1000);
				Input.MoveTo(ManPoint);
				Console.WriteLine("move manual point");
				Thread.Sleep(2000);
				Console.WriteLine("click manual point");
				Input.MoveToAndClick(ManPoint);
				Thread.Sleep(2000);

				// Find Restart Loc Dialog Window
				System.Windows.Automation.Condition c2 = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, "Manual Location"),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the Manual Location Dialog element
				AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
				if (aeDialog == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = " Dialog Window not found";
					Console.WriteLine("FindElementByID failed: Dialog");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(testname + " ---fail --- " + StateValue);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
				}
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion

		#region RestartAgv
		public static void RestartAgv(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find AGV GridView
				AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find AgvDataGrid failed:" + AGV_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellState, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + "State Row 0", Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				*/
				Input.MoveToAndRightClick(point);

				System.Windows.Point RestartPoint = new Point(point.X, point.Y + 55);
				Thread.Sleep(2000);
				Input.MoveTo(RestartPoint);
				Thread.Sleep(2000);
				Input.MoveToAndClick(RestartPoint);
				Thread.Sleep(2000);

				// Find Restart Agv Dialog Window
				System.Windows.Automation.Condition c2 = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, "Restart Agv"),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ARestart Agv Dialog element
				AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
				if (aeDialog == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = " Dialog Window not found";
					Console.WriteLine("FindElementByID failed: Dialog");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Console.WriteLine(testname + " ---pass --- " + StateValue);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(testname + " ---fail --- " + StateValue);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
				}
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				sErrorMessage = sErrorMessage.Trim();
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion

		#region AgvJobOverview
		public static void AgvJobOverview(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, Constants.TEST);
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
				System.Windows.Point JobsPoint = new Point(point.X, point.Y + 195);
				Thread.Sleep(2000);
				Input.MoveTo(JobsPoint);
				Thread.Sleep(2000);
				Input.MoveToAndClick(JobsPoint);
				Thread.Sleep(2000);

				// Find Agv Job Overview Window
				string JobsWindowID = "Job overview - " + AgvValue;
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					Console.WriteLine(testname + " ---pass --- ");
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion

		#region AgvJobOverviewOpenDetail
		public static void AgvJobOverviewOpenDetail(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, Constants.TEST);
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
				System.Windows.Point JobsPoint = new Point(point.X, point.Y + 195);
				Thread.Sleep(2000);
				Input.MoveTo(JobsPoint);
				Thread.Sleep(2000);
				Input.MoveToAndClick(JobsPoint);
				Thread.Sleep(2000);

				// Find Agv Job Overview Window
				string JobsWindowID = "Job overview - " + AgvValue;
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find Job Detail screen        
				AutomationElement aeJobGrid = AUIUtilities.FindElementByID(DATAGRIDVIEW_ID, aeJobsOverview);
				if (aeJobGrid == null)
				{
					Console.WriteLine("Find JobDataGridView failed:" + DATAGRIDVIEW_ID);
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find JobDataGridView failed:" + "JobDataGridView", Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
					Console.WriteLine("JobDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

				string JobCellname = "Id Row 0";
				AutomationElement aeJobCell = AUIUtilities.FindElementByName(JobCellname, aeJobGrid);
				if (aeJobCell == null)
				{
					sErrorMessage = "Find JobDataGridView aeCell failed:" + JobCellname;
					Console.WriteLine("Find JobDataGridView aeCell failed:" + JobCellname);
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find JobDataGridView cell failed:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find aeJobCell value failed:" + cellname, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}


				Console.WriteLine("Job Id value: " + JobValue);
				Epia3Common.WriteTestLogMsg(slogFilePath, "Job Id cell value: " + JobValue, Constants.TEST);

				Thread.Sleep(3000);
				System.Windows.Point JobPoint = AUIUtilities.GetElementCenterPoint(aeJobCell);
				Input.MoveToAndDoubleClick(JobPoint);

				Thread.Sleep(3000);

				//string JobValue = "JOB1";
				string JobScreenName = "Job detail - " + AgvValue + " - " + JobValue;
				Console.WriteLine("JobScreenName: " + JobScreenName);
				Epia3Common.WriteTestLogMsg(slogFilePath, "JobScreenName: " + JobScreenName, Constants.TEST);

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
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				string JobTextValue = aeJobText.Current.Name;
				if (JobTextValue.StartsWith(JobValue))
				{
					Console.WriteLine("--pass--");
					result = ConstCommon.TEST_PASS;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = JobScreenName + "Job Value should be " + JobValue + ", but  " + JobTextValue;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion

		#region AgvModeSemiAutomatic
		public static void AgvModeSemiAutomatic(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "---" + ex.StackTrace, Constants.TEST);
					return;
				}

				string StateValue = AUICommon.GetDataGridViewCellValueAt(0, "Mode", aeGrid);
				if (StateValue == null || StateValue == string.Empty)
				{
					sErrorMessage = "AgvDataGridView aeCell Value not found:" + "Mode Row 0";
					Console.WriteLine("AgvDataGridView aeCell Value not found:" + "Mode Row 0");
					Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + "Mode Row 0", Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				Input.MoveToAndRightClick(point);

				System.Windows.Point RestartPoint = new Point(point.X, point.Y + 60);
				Thread.Sleep(2000);
				Input.MoveTo(RestartPoint);
				Thread.Sleep(2000);

				System.Windows.Point ModePoint = new Point(RestartPoint.X, RestartPoint.Y + 145);
				//System.Windows.Point ModePoint = new Point(RestartPoint.X, RestartPoint.Y + 170);
				Thread.Sleep(2000);
				Input.MoveTo(ModePoint);
				Console.WriteLine("modepoint");
				Input.MoveToAndClick(ModePoint);
				Thread.Sleep(2000);

				System.Windows.Point PointH = new Point(ModePoint.X + 200, ModePoint.Y);
				Thread.Sleep(2000);
				Input.MoveTo(PointH);
				Console.WriteLine(" horizontal point");
				Thread.Sleep(2000);
				Input.MoveTo(PointH);

				Thread.Sleep(2000);
				System.Windows.Point SemiPoint = new Point(PointH.X, PointH.Y + 45);
				Thread.Sleep(2000);
				Input.MoveTo(SemiPoint);
				Console.WriteLine("semipoint");

				Input.MoveToAndClick(SemiPoint);
				Thread.Sleep(2000);

				// Find Restart Agv Dialog Window
				System.Windows.Automation.Condition c2 = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, "SemiAuto Agv"),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ARestart Agv Dialog element
				AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
				if (aeDialog == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = " Dialog Window not found";
					Console.WriteLine("FindElementByID failed: Dialog");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeYes));
				Thread.Sleep(2000);
				StateValue = AUICommon.GetDataGridViewCellValueAt(0, "Mode", aeGrid);

				sStartTime = DateTime.Now;
				TimeSpan mTime = DateTime.Now - sStartTime;
				Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
				while (!StateValue.Equals("SemiAuto") && mTime.Seconds < 30)
				{
					Thread.Sleep(2000);
					mTime = DateTime.Now - sStartTime;
					StateValue = AUICommon.GetDataGridViewCellValueAt(1, "Mode", aeGrid);
					Console.WriteLine("time is (sec) : " + mTime.Seconds + " and mode is " + StateValue);
				}

				if (StateValue.Equals("SemiAuto"))
				{
					result = ConstCommon.TEST_PASS;
					Console.WriteLine(testname + " ---pass --- " + StateValue);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(testname + " ---fail --- " + StateValue);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
				}
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region AgvsAllModeRemoved
		public static void AgvsAllModeRemoved(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, ex.Message + "---" + ex.StackTrace, Constants.TEST);
					return;
				}

				string StateValue = AUICommon.GetDataGridViewCellValueAt(0, "Mode", aeGrid);
				if (StateValue == null || StateValue == string.Empty)
				{
					sErrorMessage = "AgvDataGridView aeCell Value not found:" + "Mode Row 0";
					Console.WriteLine("AgvDataGridView aeCell Value not found:" + "Mode Row 0");
					Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + "Mode Row 0", Constants.TEST);
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

				System.Windows.Point RestartPoint = new Point(point.X, point.Y + 60);
				Thread.Sleep(2000);
				Input.MoveTo(RestartPoint);
				Thread.Sleep(2000);

				System.Windows.Point ModePoint = new Point(RestartPoint.X, RestartPoint.Y + 120);
				Thread.Sleep(2000);
				Input.MoveTo(ModePoint);
				Console.WriteLine("modepoint");

				Thread.Sleep(2000);
				System.Windows.Point SemiPoint = new Point(ModePoint.X + 210, ModePoint.Y + 80);
				Thread.Sleep(2000);
				Input.MoveTo(SemiPoint);
				Console.WriteLine("semipoint");

				Input.MoveToAndClick(SemiPoint);
				Thread.Sleep(2000);

				// Find Restart Agv Dialog Window
				System.Windows.Automation.Condition c2 = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, "Remove Agvs"),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the ARestart Agv Dialog element
				AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
				if (aeDialog == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = " Dialog Window not found";
					Console.WriteLine("FindElementByID failed: Dialog");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(testname + " ---fail --- " + StateValue);
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
						Console.WriteLine(testname + " " + i + " de ---pass --- " + StateValue);
					}
					else
					{
						result = ConstCommon.TEST_FAIL;
						Console.WriteLine(testname + " ---fail --- " + StateValue);
						Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
						return;
					}
				}
				Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region AgvsIdSorting
		public static void AgvsIdSorting(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, AGV_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				AutomationElement aeGrid = AUIUtilities.FindElementByID(AGV_GRIDDATA_ID, aeOverview);
				if (aeGrid == null)
				{
					sErrorMessage = "Find AgvDataGridView failed:" + AGV_GRIDDATA_ID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
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
						Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, Constants.TEST);
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
						Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, Constants.TEST);
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
						Epia3Common.WriteTestLogMsg(slogFilePath, "Find AgvDataGridView cell failed:" + cellname, Constants.TEST);
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
						Epia3Common.WriteTestLogMsg(slogFilePath, "AgvDataGridView cell value not found:" + cellname, Constants.TEST);
						result = ConstCommon.TEST_FAIL;
						return;
					}

					AgvsIdCells[i] = AgvValue;
				}

				Console.WriteLine(" ---Total agvs --- " + sNumAgvs);
				Epia3Common.WriteTestLogMsg(slogFilePath, " ---Total agvs --- " + sNumAgvs, Constants.TEST);

				bool sortResult = true;
				// Check result
				for (int i = 0; i < sNumAgvs - 1; i++)
				{
					if (PreSortAscending)
					{
						//Console.WriteLine(AgvsIdCells[i] + " ---compare --- " + AgvsIdCells[i + 1]);
						Epia3Common.WriteTestLogMsg(slogFilePath, AgvsIdCells[i] + " ---compare --- " + AgvsIdCells[i + 1], Constants.TEST);

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
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine(testname + " ---fail --- ");
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
					return;
				}

				Thread.Sleep(3000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion
		
		#region CreateNewTransport
		public static void CreateNewTransport(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					//Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
							Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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

						Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(MoverItem));
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}
				else
				{
					Console.WriteLine("aeCreate Found:");
					Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCreate));
					Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeCreate));
				}

				Thread.Sleep(4000);

				// Dispose new Transport screen, Find Cancel element. 
				string CancelId = "m_btnCancel";
				AutomationElement aeCancel = AUIUtilities.FindElementByID(CancelId, aeNewT);
				if (aeCancel == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = "New Transport aeCancel not found";
					Console.WriteLine("FindElementByID failed:" + NEW_TRANSPORT);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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


				Thread.Sleep(2000);
				// Check transport created
				//AUICommon.ClearDisplayedScreens(root);
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_PASS;
					string ms = TRANSPORT_OVERVIEW + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
					Console.WriteLine(ms);
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}


		}
		#endregion CreateNewTransport

		#region EditTransport
		public static void EditTransport(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find Transport GridView
				AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
				if (aeGrid == null)
				{
					Console.WriteLine("Find TransportDataGridView failed:" + "TransportDataGridView");
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView failed:" + "TransportDataGridView", Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView cell failed:" + cellname, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("cell TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

				Thread.Sleep(3000);
				System.Windows.Point point = AUIUtilities.GetElementCenterPoint(aeCell);

				// Open Edit Transport Screen
				Input.MoveToAndRightClick(point);
				Thread.Sleep(3000);
				System.Windows.Point CancelPoint = new Point(point.X + 20, point.Y + 95);
				Thread.Sleep(2000);
				Input.MoveTo(CancelPoint);
				Thread.Sleep(2000);
				Input.MoveToAndClick(CancelPoint);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView state cell failed:" + cellname_state, Constants.TEST);
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
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);

				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = " Transport destination is not changed to " + DestID + " , but:" + StateValue;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}


		}
		#endregion EditTransport

		#region CancelTransport
		public static void CancelTransport(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);
				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Console.WriteLine("FindElementByID failed:" + TRANSPORT_OVERVIEW);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				// Find Transport GridView
				AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
				if (aeGrid == null)
				{
					Console.WriteLine("Find TransportDataGridView failed:" + "TransportDataGridView");
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView failed:" + "TransportDataGridView", Constants.TEST);
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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView cell failed:" + cellname, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}
				else
				{
					Console.WriteLine("cell TransportDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
				}

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
					Console.WriteLine("TransportDataGridView aeCell Value not found:" + cellname);
					Epia3Common.WriteTestLogMsg(slogFilePath, "TransportDataGridView cell value not found:" + cellname, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				//string TransportValue = AUICommon.GetDataGridViewCellValueAt(0, "Id", aeGrid);
				Input.MoveToAndRightClick(point);
				Thread.Sleep(3000);
				System.Windows.Point CancelPoint = new Point(point.X + 20, point.Y + 55);
				Thread.Sleep(2000);
				Input.MoveTo(CancelPoint);
				Thread.Sleep(2000);
				Input.MoveToAndClick(CancelPoint);
				Thread.Sleep(2000);

				// Find Cancel Transport Dialog Window
				System.Windows.Automation.Condition c2 = new AndCondition(
				   new PropertyCondition(AutomationElement.NameProperty, "Cancel Transport"),
				   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
			   );

				// Find the Cancel Dialog element
				AutomationElement aeDialog = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
				if (aeDialog == null)
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = " Dialog Window not found";
					Console.WriteLine("FindElementByID failed: Dialog");
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeYes));
				Thread.Sleep(5000);

				// Check State Value
				Thread.Sleep(3000);
				// Construct the Grid Cell Element Name
				string cellname_state = "State Row 0";
				// Get the Element with the Row Col Coordinates
				AutomationElement aeStateCell = AUIUtilities.FindElementByName(cellname_state, aeGrid);

				if (aeStateCell == null)
				{
					Console.WriteLine("Find TransportDataGridView aeStateCell failed:" + cellname_state);
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView state cell failed:" + cellname_state, Constants.TEST);
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
				if (StateValue.Equals("Finished"))
				{
					result = ConstCommon.TEST_PASS;
					Console.WriteLine(testname + " ---pass --- ");
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);

				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = " Transport state is not Finished, but:" + StateValue;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}

				Thread.Sleep(2000);
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
			}


		}
		#endregion CancelTransport

		#region TransportOverviewOpenDetail
		public static void TransportOverviewOpenDetail(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;

			AutomationElement aePanelLink = null;
			try
			{
				AUICommon.ClearDisplayedScreens(root);

				aePanelLink = AUICommon.FindTreeViewNodeLevelAll(testname, root, INFRASTRUCTURE, TRANSPORT_OVERVIEW, ref sErrorMessage);
				if (aePanelLink == null)
				{
					Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
					return;
				}

				AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
				if (aeGrid == null)
				{
					result = ConstCommon.TEST_FAIL;
					Console.WriteLine("Find TransportDataGridView failed:" + "TransportDataGridView");
					Epia3Common.WriteTestLogFail(slogFilePath, "Find TransportDataGridView failed:" + "TransportDataGridView", Constants.TEST);

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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Find TransportDataGridView cell failed:" + cellname, Constants.TEST);
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
				Epia3Common.WriteTestLogMsg(slogFilePath, "Id cell value: " + cellValue, Constants.TEST);

				string TrnScreenName = "Transport detail - " + cellValue;
				Console.WriteLine("TrnScreenName: " + TrnScreenName);
				Epia3Common.WriteTestLogMsg(slogFilePath, "TrnScreenName: " + TrnScreenName, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, "Find TransportDetailView failed:" + "TransportDetailView", Constants.TEST);
					return;
				}

				// Check Transport ID Value
				string textID = "m_IdValueLabel";
				AutomationElement aeTrnText = AUIUtilities.FindElementByID(textID, aeDetailScreen);
				if (aeTrnText == null)
				{
					sErrorMessage = "Find TrnTextElement failed:" + textID;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogMsg(slogFilePath, sErrorMessage, Constants.TEST);
					result = ConstCommon.TEST_FAIL;
					return;
				}

				string TrnTextValue = aeTrnText.Current.Name;
				if (TrnTextValue.Equals(cellValue))
				{
					result = ConstCommon.TEST_PASS;
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
				}
				else
				{
					result = ConstCommon.TEST_FAIL;
					sErrorMessage = TrnScreenName + "Trn Value should be " + cellValue + ", but  " + TrnTextValue;
					Console.WriteLine(sErrorMessage);
					Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, Constants.TEST);
				}
			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				sErrorMessage = ex.Message + "----: " + ex.StackTrace;
				Console.WriteLine("Fatal error: " + sErrorMessage);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + sErrorMessage, Constants.TEST);
			}
		}
		#endregion

		#region Epia3Close
		public static void Epia3Close(string testname, AutomationElement root, out int result)
		{
			Console.WriteLine("\n=== Test " + testname + " ===");
			Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, Constants.TEST);
			result = ConstCommon.TEST_UNDEFINED;
			string BtnCloseID = "Close";
			try
			{
				AUICommon.ClearDisplayedScreens(root);

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
					Epia3Common.WriteTestLogMsg(slogFilePath, "Epia3 Closed", Constants.TEST);
					Console.WriteLine("\nTest scenario: Pass");
					Epia3Common.WriteTestLogPass(slogFilePath, testname, Constants.TEST);
					result = ConstCommon.TEST_PASS;
				}
				else
				{
					Console.WriteLine("process id :" + pID);
					Epia3Common.WriteTestLogMsg(slogFilePath, "process id :" + pID, Constants.TEST);
					Console.WriteLine("\nTest scenario: *FAIL*");
					result = ConstCommon.TEST_FAIL;
					Epia3Common.WriteTestLogFail(slogFilePath, testname, Constants.TEST);
				}
				Thread.Sleep(3000);

			}
			catch (Exception ex)
			{
				result = ConstCommon.TEST_FAIL;
				Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
				Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, Constants.TEST);
			}
		}
		#endregion Epia3Close
		
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
				Epia3Common.WriteTestLogFail(slogFilePath, "shell start exception: " + sErrorMessage, Constants.TEST);
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
					//Epia3Common.WriteTestLogFail(slogFilePath, "shell start exception: " + sErrorMessage, Constants.TEST);
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
				Epia3Common.WriteTestLogMsg(slogFilePath, "SERVER open window name: " + name, Constants.TEST);
				return;
			}
			else if (name.Equals("ThemeManagerNotification"))
			{
				Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
				return;
			}
			else if (name.Equals("Epia security"))
			{
				Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
				return;
			}
			else if (name.Equals("Egemin e'pia User Interface Shell"))
			{
				Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
				Thread.Sleep(3000);
			}
			else if (name.Equals("Egemin Shell"))
			{
				Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
				Thread.Sleep(3000);
			}
			else if (name.Equals("Open File - Security Warning"))
			{
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
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
				Epia3Common.WriteTestLogMsg(slogFilePath, "SERVER open other window name: " + name, Constants.TEST);
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
				Epia3Common.WriteTestLogFail(slogFilePath, "shell start exception: " + sErrorMessage, Constants.TEST);
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
					Epia3Common.WriteTestLogFail(slogFilePath, "shell start exception: " + sErrorMessage, Constants.TEST);
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
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
				return;
			}
			else if (name.Equals("ThemeManagerNotification"))
			{
				Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
				return;
			}
			else if (name.Equals("Epia security"))
			{
				Console.WriteLine("Do YYYYYYYYYYYY Name is ------------:" + name);
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
				return;
			}
			else if (name.Equals("Egemin e'pia User Interface Shell"))
			{
				Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
				Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
				Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
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
							Epia3Common.WriteTestLogFail(slogFilePath, "start shell failed: " + sErrorMessage, Constants.TEST);
						}
						else
						{
							Console.WriteLine("Error Message not found ------------:");
							Epia3Common.WriteTestLogFail(slogFilePath, "Error Message pane not found: ", Constants.TEST);
						}
					}

					TestCheck = ConstCommon.TEST_FAIL;
				}
				else
				{
					Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
					Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
					Console.WriteLine("Do OOOOOOOOOOOO Name is ------------:" + name);
					Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
					Thread.Sleep(3000);
				}
			}
			else if (name.Equals("Open File - Security Warning"))
			{
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
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
				Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, Constants.TEST);
			}
			sEventEnd = true;
		}
		#endregion
		#endregion
		
		public static AutomationElement FindCellFromGrid(AutomationElement aeGrid,
			string Header, string CellValue, ref int CellRow)
		{
			for (int i = 0; i < 40; i++)
			{
				// Construct the Grid Cell Element Name
				string cellname = Header+" Row " + i;
				// Get the Element with the Row Col Coordinates
				Console.WriteLine("GET cell: " + Header+" Row " + i);
				AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

				if (aeCell == null)
				{
					return null;
				}
				else
				{
					Console.WriteLine("cell found at row: " + i);
				}

				// check cell value
				string GetValue = string.Empty;
				try
				{
					ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
					GetValue = vp.Current.Value;
					Console.WriteLine("Get element.Current Value:" + GetValue);
				}
				catch (System.NullReferenceException)
				{
					GetValue = string.Empty;
					return null;
				}

				if (GetValue.Equals(CellValue))
				{ 
					CellRow = i;
					Thread.Sleep(2000);
					Input.MoveToAndClick(aeCell);
					Thread.Sleep(2000);   
					return aeCell;
				}
			}
			return null;
		}

		public static AutomationElement FindCellFromGridStartedAtRow(AutomationElement aeGrid,
			string header, string cellValue, int startRow, ref int cellRow)
		{
			for (int i = startRow; i < startRow+40; i++)
			{
				// Construct the Grid Cell Element Name
				string cellname = header + " Row " + i;
				// Get the Element with the Row Col Coordinates
				AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);

				if (aeCell == null)
				{
					return null;
				}
				else
				{
					Console.WriteLine("cell found at row: " + i);
				}

				// check cell value
				string GetValue = string.Empty;
				try
				{
					ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
					GetValue = vp.Current.Value;
					Console.WriteLine("Get element.Current Value:" + GetValue);
				}
				catch (System.NullReferenceException)
				{
					GetValue = string.Empty;
					return null;
				}

				if (GetValue.Equals(cellValue))
				{
					cellRow = i;
					Thread.Sleep(2000);
					Thread.Sleep(2000);
					return aeCell;
				}
			}
			return null;
		}

		// Find Cell at specified row
		public static AutomationElement FindCellFromGridAtHeaderRow(AutomationElement aeGrid,
			string Header, int CellRow)
		{
			// Construct the Grid Cell Element Name
			string cellname = Header + " Row " + CellRow;
			// Get the Element with the Row Col Coordinates
			AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
			return aeCell;
		}

		// Click Cancel button
		public static void ClickCancelButton(string CancelID, AutomationElement ThisScreen )
		{
			// Find Cancel Button  element. 
			AutomationElement aeCancelBtn = AUIUtilities.FindElementByID(CancelID, ThisScreen);
			if (aeCancelBtn == null)
			{
				//sErrorMessage = sErrorMessage + "Cancel Button aeCancelBtn not found";
				Console.WriteLine("FindElementByID failed:" + CancelID);
				//Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, Constants.TEST);
			}
			else
			{
				Console.WriteLine("found aeCancelBtn:");
				Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeCancelBtn));
				Thread.Sleep(2000);
				Input.MoveToAndClick(aeCancelBtn);
				Thread.Sleep(2000);
			}
		}

		public static void SendEmail(string resultFile)
		{
			string sendMail = "false";
			string layout = "EPIA3";
			int failedCounter = sTotalFailed;

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

			try
			{
				System.Net.Mail.MailMessage oMsg = new System.Net.Mail.MailMessage();
				System.Net.Mail.Attachment oAttch = new System.Net.Mail.Attachment(resultFile); 

				int sendHour = System.DateTime.Now.Hour;
				oMsg.From = new System.Net.Mail.MailAddress("TeamSystems@Egemin.be");
				// TODO: Replace with recipient e-mail address.				
				if (sendMail.ToLower().ToString().StartsWith("false"))
				{					
					oMsg.To.Add("Jiemin.Shi@egemin.be;");
					oMsg.Subject = "Only ME(GUI-" + layout + ")[" + sBuildNr + "]" + System.DateTime.Now.ToString("ddMMM-HH:mm")
						+ "-[" + System.Environment.MachineName + "]";
				}
				else
				{
					//if (failedCounter > 0)
					//{
					oMsg.To.Add("Jiemin.Shi@Egemin.be,Wim.VanBetsbrugge@Egemin.be,Dirk.Declercq@Egemin.be");
					oMsg.Subject = "Test Result (" + layout + ")[" + sBuildNr + "]" + System.DateTime.Now.ToString("ddMMM-HH:mm") + "-[" + System.Environment.MachineName + "]";
					//}
					//else
					//{
					//    oMsg.To = "jiemin.shi@egemin.be;";
					//    oMsg.Subject = "E'pia Nightly Test OK (" + layout + ")[" + buildType + "]" + System.DateTime.Now.ToString("ddMMM-HH:mm")
					//    + "-[" + System.Environment.MachineName + "]";
					//}
				}
				oMsg.IsBodyHtml = true;
				// HTML Body (remove HTML tags for plain text).
				//oMsg.Body = "<HTML><BODY><B>Hello World!</B></BODY></HTML>";
				oMsg.Body = TextStatistics; // +testInputData;
				oMsg.Attachments.Add(oAttch);

				Epia3Common.WriteTestLogMsg(slogFilePath, "--------------------------------", Constants.TEST);
				Epia3Common.WriteTestLogMsg(slogFilePath, "SmtpServer: " + ConstCommon.SMTP_SERVERID, Constants.TEST);
				Epia3Common.WriteTestLogMsg(slogFilePath, "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx send mail ======: " + sendMail, Constants.TEST);
				System.Net.Mail.SmtpClient MailClient = new System.Net.Mail.SmtpClient(ConstCommon.SMTP_SERVERID);

				try
				{
					MailClient.Send(oMsg);
				}
				catch (Exception ex)
				{
					Epia3Common.WriteTestLogMsg(slogFilePath, "The following exception occurred: " + ex.ToString(), Constants.TEST);
					//check the InnerException
					while (ex.InnerException != null)
					{
						Epia3Common.WriteTestLogMsg(slogFilePath, "--------------------------------", Constants.TEST);
						Epia3Common.WriteTestLogMsg(slogFilePath, "The following InnerException reported: " + ex.InnerException.ToString(), Constants.TEST);
						ex = ex.InnerException;
					}
				}
				Epia3Common.WriteTestLogMsg(slogFilePath, "email sent to developers ", Constants.TEST);
				oMsg = null;
				oAttch = null;
			}
			catch (Exception e)
			{
				Epia3Common.WriteTestLogMsg(slogFilePath, "send mail : " + e.Message + "----" + e.StackTrace, Constants.TEST);
				Console.WriteLine("{0} Exception caught.", e);
			}
		}
	}
}
