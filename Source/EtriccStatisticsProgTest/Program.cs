using System;
using System.IO;
using System.Configuration;
using System.Diagnostics;
using System.ServiceProcess;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Xml;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using TestTools;
using Excel = Microsoft.Office.Interop.Excel;

using ICSharpCode.SharpZipLib.Zip;

namespace EtriccStatisticsProgTest
{
    class Program
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
        static string sBuildConfig = string.Empty;
        static string sBuildDef = string.Empty;
        static string sBuildNr = string.Empty;
        static string sTestToolsVersion = string.Empty;
        static string m_SystemDrive = string.Empty;
        static string sCurrentProject = "Test";
        static string sTargetPlatform = string.Empty;
        static string sCurrentPlatform = string.Empty;
        static string sTestResultFolder = string.Empty;
        // Testcase not used =================================
        public static string sConfigurationName = string.Empty;
        static string sErrorMessage;
        static bool sEventEnd;
        static bool sErrorScreen = false;
        static string sExcelVisible = string.Empty;
        static bool sAutoTest = true;
        static string sInstallScriptsDir = string.Empty;
        public static string sLayoutName = string.Empty;
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
        static IBuildServer m_BuildSvc;
        static bool TFSConnected = true;
        // excel 	--------------------------------------------------------
        static Excel.Application xApp;
        static Excel.Workbook xBook;
        static Excel.Workbooks xBooks;
        static Excel.Range xRange;
        //static Excel.Worksheet      xSheet;
        static dynamic xSheet;

        static AutomationEventHandler UIErrorEventHandler = new AutomationEventHandler(OnErrorUIEvent);
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

            Console.WriteLine("sOnlyUITest : " + sOnlyUITest);

            sCurrentProject = System.Configuration.ConfigurationManager.AppSettings.Get("CurrentProject");
            Console.WriteLine("sCurrentProject : " + sCurrentProject);

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

                        Epia3Common.WriteTestLogMsg(slogFilePath, "m_SystemDrive: " + m_SystemDrive, sOnlyUITest);

                        // sInstallScriptsDir = X:\CI\Epia 3\Epia - CI_20100324.1\Debug\Installation\Setup\Debug
                        // sBuildDropFolder = X:\CI\Epia 3\Epia - CI_20100324.1
                        // sBuildNr = Epia - CI_20100324.1
                        // sTestApp = Epia
                        // sBuildDef = CI
                        // sBuildConfig = Debug
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
                                TestTools.MessageBoxEx.Show("Team Foundation services are not available from server\nWill try to reconnect the Server after 10 minutes",
                                 kTime++ + " During E'tricc Statistics UI Testing, please not touch the screen, time:" + DateTime.Now.ToLongTimeString(), 10 * 60000);
                                System.Threading.Thread.Sleep(10 * 60000);
                                conn = false;
                            }
                            catch (Exception ex)
                            {
                                TestTools.MessageBoxEx.Show("TeamFoundation getService Exception:" + ex.Message + " ----- " + ex.StackTrace,
                                     kTime++ + " This is automatic testing, please not touch the screen: exception, time:" + DateTime.Now.ToLongTimeString(), 10 * 60000);
                                System.Threading.Thread.Sleep(10 * 60000);
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
            sTestCaseName[13] = ReportName.ANALYSIS_ProjectActivation;
            sTestCaseName[14] = ReportName.ANALYSIS_TransportLookupBySrcDstGroup;
            sTestCaseName[15] = ReportName.ANALYSIS_TransportLookupBySrcDstLocationOrStation;
            sTestCaseName[16] = ReportName.ANALYSIS_TransportWithJobsAndStatusHistory;
            sTestCaseName[17] = ReportName.ANALYSIS_LoadHistory;
            //====_VEHICLES=========================================//
            sTestCaseName[18] = ReportName.PERFORMANCE_VEHICLES_ModeOverview;
            sTestCaseName[19] = ReportName.PERFORMANCE_VEHICLES_StateOverview;
            sTestCaseName[20] = ReportName.PERFORMANCE_VEHICLES_StatusOverview;
            sTestCaseName[21] = ReportName.PERFORMANCE_VEHICLES_StatusCountDayTrend;
            sTestCaseName[22] = ReportName.PERFORMANCE_VEHICLES_StatusCountTop;
            sTestCaseName[23] = ReportName.PERFORMANCE_VEHICLES_StatusDurationDayTrend;
            sTestCaseName[24] = ReportName.PERFORMANCE_VEHICLES_StatusDurationTop;
            //====TRANSPORTS=========================================//
            sTestCaseName[25] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupHour;
            sTestCaseName[26] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupDay;
            sTestCaseName[27] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupMonth;
            sTestCaseName[28] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationHour;
            sTestCaseName[29] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationDay;
            sTestCaseName[30] = ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationMonth;
            //====JOBS=========================================//
            sTestCaseName[31] = ReportName.PERFORMANCE_JOBS_CountByLocationInGroupDay;
            sTestCaseName[32] = ReportName.PERFORMANCE_JOBS_CountByLocationInGroupMonth;
            sTestCaseName[33] = ReportName.PERFORMANCE_JOBS_CountByLocationDay;
            sTestCaseName[34] = ReportName.PERFORMANCE_JOBS_CountByLocationMonth;
           
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
                    || PCName.ToUpper().Equals("EPIATESTSRV3V1"))
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
                    xSheet.Cells[1, 2] = "ETRICC STATISTICS DEPLOYMENT AND UI TESTS";
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
                    xSheet.Cells[10, 1] = "Test Project:";
                    xSheet.Cells[10, 2] = sCurrentProject;
                }

                // start test----------------------------------------------------------
                int sResult = ConstCommon.TEST_UNDEFINED;
                int aantal = 35;
                if (sDemo)
                    aantal = 2;

                if (sOnlyUITest)
                {
                    sTestType = System.Configuration.ConfigurationManager.AppSettings.Get("TestType");
                    if (sTestType.ToLower().StartsWith("all"))
                        aantal = 35;
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
                    Epia3Common.WriteTestLogMsg(slogFilePath, " has build quality", sOnlyUITest);
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
                            TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Started", "EtriccStatistics+" + sCurrentPlatform);
                            Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Started:" + ConstCommon.ETRICCSTATISTICS, sOnlyUITest);
                        }
                    }
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
                            OpenGraphicalView(ReportName.StatusGraphicalView, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_VEHICLES_ModeOverview:
                            OpenPerformanceVehiclesReports(ReportName.PERFORMANCE_VEHICLES_ModeOverview, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_VEHICLES_StateOverview:
                            OpenPerformanceVehiclesReports(ReportName.PERFORMANCE_VEHICLES_StateOverview, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_VEHICLES_StatusOverview:
                            OpenPerformanceVehiclesReports(ReportName.PERFORMANCE_VEHICLES_StatusOverview, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_VEHICLES_StatusCountDayTrend:
                            OpenPerformanceVehiclesReports(ReportName.PERFORMANCE_VEHICLES_StatusCountDayTrend, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_VEHICLES_StatusCountTop:
                            OpenPerformanceVehiclesReports(ReportName.PERFORMANCE_VEHICLES_StatusCountTop, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_VEHICLES_StatusDurationDayTrend:
                            OpenPerformanceVehiclesReports(ReportName.PERFORMANCE_VEHICLES_StatusDurationDayTrend, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_VEHICLES_StatusDurationTop:
                            OpenPerformanceVehiclesReports(ReportName.PERFORMANCE_VEHICLES_StatusDurationTop, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupHour:
                            OpenPerformanceTransportsReports(ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupHour, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupDay:
                            OpenPerformanceTransportsReports(ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupDay, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupMonth:
                            OpenPerformanceTransportsReports(ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstGroupMonth, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationHour:
                            OpenPerformanceTransportsReports(ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationHour, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationDay:
                            OpenPerformanceTransportsReports(ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationDay, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationMonth:
                            OpenPerformanceTransportsReports(ReportName.PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationMonth, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_JOBS_CountByLocationInGroupDay:
                            OpenPerformanceJobsReports(ReportName.PERFORMANCE_JOBS_CountByLocationInGroupDay, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_JOBS_CountByLocationInGroupMonth:
                            OpenPerformanceJobsReports(ReportName.PERFORMANCE_JOBS_CountByLocationInGroupMonth, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_JOBS_CountByLocationDay:
                            OpenPerformanceJobsReports(ReportName.PERFORMANCE_JOBS_CountByLocationDay, aeForm, out sResult);
                            break;
                        case ReportName.PERFORMANCE_JOBS_CountByLocationMonth:
                            OpenPerformanceJobsReports(ReportName.PERFORMANCE_JOBS_CountByLocationMonth, aeForm, out sResult);
                            break;
                        case ReportName.ANALYSIS_ProjectActivation:
                            OpenAnalysisReports(ReportName.ANALYSIS_ProjectActivation, aeForm, out sResult);
                            break;
                        case ReportName.ANALYSIS_TransportLookupBySrcDstGroup:
                            OpenAnalysisReports(ReportName.ANALYSIS_TransportLookupBySrcDstGroup, aeForm, out sResult);
                            break;
                        case ReportName.ANALYSIS_TransportLookupBySrcDstLocationOrStation:
                            OpenAnalysisReports(ReportName.ANALYSIS_TransportLookupBySrcDstLocationOrStation, aeForm, out sResult);
                            break;
                        case ReportName.ANALYSIS_TransportWithJobsAndStatusHistory:
                            OpenAnalysisReports(ReportName.ANALYSIS_TransportWithJobsAndStatusHistory, aeForm, out sResult);
                            break;
                        case ReportName.ANALYSIS_LoadHistory:
                            OpenAnalysisReports(ReportName.ANALYSIS_LoadHistory, aeForm, out sResult);
                            break;
                        default:
                            break;
                    }

                    // write result to Excel
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
                    xSheet.Cells[Counter + 2 + 9, 1] = "Total tests: ";
                    xSheet.Cells[Counter + 3 + 9, 1] = "Total Passes: ";
                    xSheet.Cells[Counter + 4 + 9, 1] = "Total Failed: ";

                    xSheet.Cells[Counter + 2 + 9, 2] = sTotalCounter;
                    xSheet.Cells[Counter + 3 + 9, 2] = sTotalPassed;
                    xSheet.Cells[Counter + 4 + 9, 2] = sTotalFailed;

                    ulong TPhysicalMem = 0;
                    ulong APhysicalMem = 0;
                    ulong TVirtualMem = 0;
                    ulong AVirtualMem = 0;

                    HelpUtilities.GetMemoryInfo(out TPhysicalMem, out APhysicalMem, out TVirtualMem, out AVirtualMem);
                    // Add Legende
                    xSheet.Cells[Counter + 5 + 9, 2] = "Legende";
                    xRange = xApp.get_Range("B" + (Counter + 5 + 9));
                    xRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                    xSheet.Cells[Counter + 6 + 9, 2] = "Pass";
                    xRange = xApp.get_Range("B" + (Counter + 6 + 9));
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_GREEN;
                    xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                    xSheet.Cells[Counter + 7 + 9, 2] = "Fail";
                    xRange = xApp.get_Range("B" + (Counter + 7 + 9));
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_RED;
                    xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                    xSheet.Cells[Counter + 8 + 9, 2] = "Exception";
                    xRange = xApp.get_Range("B" + (Counter + 8 + 9));
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_PINK;
                    xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                    xSheet.Cells[Counter + 9 + 9, 2] = "Untested";
                    xRange = xApp.get_Range("B" + (Counter + 9 + 9));
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_YELLOW;
                    xRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xRange.Borders.Weight = Excel.XlBorderWeight.xlMedium;

                    xSheet.Cells[Counter + 10 + 10, 2] = "TotalPhysicalMemory:" + TPhysicalMem + " MB";
                    xSheet.Cells[Counter + 11 + 10, 2] = "AvailablePhysicalMemory:" + APhysicalMem + " MB";
                    xSheet.Cells[Counter + 12 + 10, 2] = "TotalVirtualMemory:" + TVirtualMem + " MB";
                    xSheet.Cells[Counter + 13 + 10, 2] = "AvailableVirtualMemory:" + AVirtualMem + " MB";
                }

                if (!sOnlyUITest)
                {
                    if (TFSConnected)
                    {
                        Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                            TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS), sBuildNr);
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
                            if (sTotalFailed == 0)
                                TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS),
                                "GUI Tests Passed", m_BuildSvc, sDemo ? "true" : "false");
                            else
                                TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS),
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
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Passed", "EtriccStatistics+" + sCurrentPlatform);
                                Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Passed:" + ConstCommon.ETRICCSTATISTICS, sOnlyUITest);
                            }
                            else
                            {
                                TestTools.FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Failed", "EtriccStatistics+" + sCurrentPlatform);
                                Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Failed:" + ConstCommon.ETRICCSTATISTICS, sOnlyUITest);
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
                Console.WriteLine("main Fatal error: " + ex.Message + "----: " + ex.StackTrace  +"---" + ex.ToString());
                Thread.Sleep(2000);
                if (sAutoTest)
                {

                    FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Exception -->" + sOutFilename + ".log", "EtriccStatistics+" + sCurrentPlatform);
                    Epia3Common.WriteTestLogMsg(slogFilePath, "GUI Tests Exception -->" + sOutFilename + ".log:" + ConstCommon.ETRICCSTATISTICS, sOnlyUITest);

                    Epia3Common.WriteTestLogFail(slogFilePath, ex.Message + "----: " + ex.StackTrace, sOnlyUITest);

                    TestTools.Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                    TestTools.Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                    TestTools.Utilities.CloseProcess("cmd");
                    TestTools.FileManipulation.UpdateTestWorkingFile(sTestResultFolder, "false");

                    if (TFSConnected)
                    {
                        Uri uri = TestTools.TfsUtilities.GetBuildUriFromBuildNumber(m_BuildSvc,
                            TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS), sBuildNr);
                        string quality = m_BuildSvc.GetMinimalBuildDetails(uri).Quality;

                        if (quality.Equals("GUI Tests Failed"))
                        {
                            TestTools.Epia3Common.WriteTestLogMsg(slogFilePath, sBuildNr + " has failed quality, no update needed :" + quality, sOnlyUITest);
                        }
                        else
                        {
                            TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                                TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS),
                                "GUI Tests Failed", m_BuildSvc, sDemo ? "true" : "false");
                        }

                    }
                }
            }
        }
        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

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
            AutomationEventHandler UIAFindLayoutPanelEventHandler = new AutomationEventHandler(OnUninstallEtriccUIEvent);
            
            try
            {
                string pa = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE");
                string systemBits = "" + ((String.IsNullOrEmpty(pa) || String.Compare(pa, 0, "x86", 0, 3, true) == 0) ? 32 : 64);

                Console.WriteLine("start: Programs and features Panel: systemBits " + systemBits);
                System.Threading.Thread.Sleep(5000);
                string InstallerSource = @"C:\Windows\System32\appwiz.cpl";
                Process Proc = new System.Diagnostics.Process();
                Proc.StartInfo.FileName = InstallerSource;
                Proc.StartInfo.CreateNoWindow = false;
                
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                if (systemBits.StartsWith("32"))
                {
                    Proc.Start();
                    Console.WriteLine("started: Programs and features Panel:" + InstallerSource);
                    Thread.Sleep(5000);
                    if (UninstallEtricc5())
                        TestCheck = ConstCommon.TEST_PASS;
                    else
                        TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {
                   Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                   AutomationElement.RootElement, TreeScope.Descendants, UIAFindLayoutPanelEventHandler);
                   Thread.Sleep(5000);
                   Proc.Start();
                   Console.WriteLine("started: Programs and features Panel:" + InstallerSource);

                    while (sEventEnd == false && mTime.Seconds <= 600)
                    {
                        Thread.Sleep(5000);
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
                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement,
                      UIAFindLayoutPanelEventHandler);


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

                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIALayoutXPosEventHandler);

                System.Threading.Thread.Sleep(15000);
                string InstallerSource = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory() + @"\EtriccStatistics\", "Etricc 5.msi");
                Console.WriteLine("start:" + InstallerSource);
                Process Proc = new System.Diagnostics.Process();
                Proc.StartInfo.FileName = InstallerSource;
                Proc.StartInfo.CreateNoWindow = false;
                Proc.Start();
                Console.WriteLine("started:" + InstallerSource);

                sEventEnd = false;
                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
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
                    Console.WriteLine("\nInstall Etricc 5.: Pass");
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
                        System.Windows.Forms.MessageBox.Show("EtriccStatistics.zip unzip exception: " + ex.Message + "--------------" + ex.StackTrace);

                    }
                }
                Thread.Sleep(5000);


                string batFile = Path.Combine(Directory.GetCurrentDirectory(), "DropDatabaseEtriccStatistics_" + sCurrentProject+".bat");
                string  sqlFile = Path.Combine(Directory.GetCurrentDirectory(), "DropDatabaseEtriccStatistics_" + sCurrentProject+".sql");
                
                //string batLine1 = "C:\Program Files\Microsoft SQL Server\100\Tools\Binn\sqlcmd"  -Stcp:%computername%,1433 -E -i DropDatabaseEtriccStatistics_Demo.sql
                string batLine1 = '"' + @"C:\"+OSVersionInfoClass.ProgramFilesx86FolderName()+@"\Microsoft SQL Server\100\Tools\Binn\sqlcmd" + '"' + "\t" 
                    + "-Stcp:%computername%,1433 -E -i DropDatabaseEtriccStatistics_" + sCurrentProject+".sql";
                string batLine2 = System.Environment.NewLine;
                string batLine3 = "pause";
                if (!File.Exists(batFile))
                {
                    StreamWriter writeBat = File.CreateText(batFile);
                    writeBat.WriteLine(batLine1);
                    writeBat.WriteLine(batLine2);
                    writeBat.WriteLine(batLine3);
                    writeBat.Close();
                }

                //IF EXISTS(SELECT 1 FROM sys.databases WHERE name = 'EtriccStatistics_Demo' )
                //DROP database [EtriccStatistics_Demo]
                string sqlLine1 = "IF EXISTS(SELECT 1 FROM sys.databases WHERE name = 'EtriccStatistics_"+ sCurrentProject+"' )";
                string sqlLine2 = "DROP database [EtriccStatistics_"+ sCurrentProject+"]";
                if (!File.Exists(sqlFile))
                {
                    StreamWriter writeSQL = File.CreateText(sqlFile);
                    writeSQL.WriteLine(sqlLine1);
                    writeSQL.WriteLine(sqlLine2);
                    writeSQL.Close();
                }

                // drop database if exist
                TestTools.Utilities.StartProcessNoWait(System.IO.Directory.GetCurrentDirectory(),
                   "DropDatabaseEtriccStatistics_"+ sCurrentProject +".bat", string.Empty);
                Thread.Sleep(10000);


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
                // Add Open MyLayoutScreen window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                   AutomationElement.RootElement, TreeScope.Descendants, UIEventHandler);

                Thread.Sleep(10000);
                Console.WriteLine("start :" + ConstCommon.PARSERCONFIGURATOR_EXE);
                string InstallerSource = System.IO.Path.Combine(TestTools.OSVersionInfoClass.ProgramFilesx86()
                    +ConstCommon.PARSERCONFIGURATOR_ROOT, ConstCommon.PARSERCONFIGURATOR_EXE);
                Console.WriteLine("InstallerSource:" + InstallerSource);
                Process Proc = new System.Diagnostics.Process();
                Proc.StartInfo.FileName = InstallerSource;
                Proc.StartInfo.CreateNoWindow = false;
                Proc.Start();
                Console.WriteLine("started:" + ConstCommon.PARSERCONFIGURATOR_EXE);

               
                sEventEnd = false;
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds);
                while (sEventEnd == false && mTime.TotalMinutes <= 3)
                {
                    Thread.Sleep(10000);
                    mTime = DateTime.Now - mStartTime;
                    Console.WriteLine("wait ParserConfiguratorConnectComputer time is :" + mTime.TotalSeconds + "   --------- sEventEnd:" + sEventEnd);

                }
                //Console.WriteLine("time is:" + mTime.TotalMilliseconds / 1000);
                //Epia3Common.WriteTestLogMsg(slogFilePath, "time is:" + mTime.TotalMilliseconds / 1000, sOnlyUITest);


                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                     AutomationElement.RootElement,
                    UIEventHandler);

                Console.WriteLine("Searching ParserConfiguratorConnectComputer windows ..................");
                Thread.Sleep(4000);

                AutomationElement aeWindow = null;
                AutomationElementCollection aeAllWindows = null;
                // find main window
                System.Windows.Automation.Condition cWindows = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.Window);

                int k = 0;
                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                while (aeWindow == null && mTime.TotalSeconds <= 120)
                {
                    Console.WriteLine("aeWindow[k]=");
                    k++;
                    aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        if (aeAllWindows[i].Current.AutomationId.Equals("MainForm"))
                        {
                            aeWindow = aeAllWindows[i];
                            Console.WriteLine("aeWindow["+i+"]=" + aeWindow.Current.Name);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }

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


                if (TestCheck == ConstCommon.TEST_FAIL)
                {
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
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            AutomationEventHandler UIEventHandler = new AutomationEventHandler(OnCreateXSDsUIEvent);

            try
            {
                Console.WriteLine("Searching main windows ..................");
                AutomationElement aeTreeView = null;
                AutomationElement aeComputerNameNode = null;
                AutomationElement aeXSDsNode = null;

                AutomationElement aeWindow = null;
                AutomationElementCollection aeAllWindows = null;
                // find main window
                System.Windows.Automation.Condition cWindows = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.Window);

                int k = 0;
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                while (aeWindow == null && mTime.TotalSeconds <= 120)
                {
                    Console.WriteLine("aeWindow[k]=");
                    k++;
                    aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        if (aeAllWindows[i].Current.AutomationId.Equals("MainForm"))
                        {
                            aeWindow = aeAllWindows[i];
                            Console.WriteLine("aeWindow[" + i + "]=" + aeWindow.Current.Name);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }

                if (aeWindow == null)
                {
                    sErrorMessage = "MainForm not found";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                {
                    Console.WriteLine("before set focus ----------------------------------------------------");
                    Thread.Sleep(2000);
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
                        aeComputerNameNode = null;
                        TreeWalker walker = TreeWalker.ControlViewWalker;
                        AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
                        while (elementNode != null)
                        {
                            Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
                            if (elementNode.Current.Name.ToLower().Equals(PCName.ToLower()))
                            {
                                //Input.MoveTo(elementNode);
                                aeComputerNameNode = elementNode;
                                Console.WriteLine("Computer node name found , it is: " + aeComputerNameNode.Current.Name);
                                TestCheck = ConstCommon.TEST_PASS;
                                break;
                            }
                            Thread.Sleep(3000);
                            elementNode = walker.GetNextSibling(elementNode);
                        }
                        //return aeNodeLink;
                        if (aeComputerNameNode == null)
                        {
                            sErrorMessage = "Computer node not found, ";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                    }
                }

                Thread.Sleep(5000); 
                // find xsd tree item from computernode
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\n=== Find XSDs node ===");
                    System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();

                    aeXSDsNode = TestTools.AUICommon.WalkEnabledElements(aeTreeView, treeNode, "XSDs");
                    if (aeXSDsNode == null)
                    {
                        sErrorMessage = string.Empty;
                        Console.WriteLine("\n=== XSDs node NOT Exist ===");
                        Input.MoveToAndDoubleClick(aeComputerNameNode.GetClickablePoint());
                        Thread.Sleep(9000);
                        aeXSDsNode = TestTools.AUICommon.WalkEnabledElements(aeTreeView, treeNode, "XSDs");
                        //result = ConstCommon.TEST_FAIL;
                        //return;
                    }
                    else
                    {
                        Console.WriteLine("\n=== XSDs node Exist ===");
                        //Input.MoveToAndClick(aeNode);
                    }
                }

                try
                {
                    ExpandCollapsePattern ep = (ExpandCollapsePattern)aeXSDsNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                    ep.Expand();
                    Thread.Sleep(1000);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("XsdNode can not expaned: " + aeXSDsNode.Current.Name);
                }


                TreeWalker walkerProj = TreeWalker.ControlViewWalker;
                AutomationElement aeMyXsdNode = walkerProj.GetFirstChild(aeXSDsNode);
                while (aeMyXsdNode != null)
                {
                    Console.WriteLine("aeMyProjNode name is: " + aeMyXsdNode.Current.Name);
                    StatUtilities.DeleteSelectedXsd(aeWindow, aeMyXsdNode);

                    Thread.Sleep(3000);
                    //aeMyProjNode = walkerProj..GetNextSibling(aeProjectsNode);
                    aeMyXsdNode = walkerProj.GetFirstChild(aeXSDsNode);
                }
               


                // Add Open Wwindow Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIEventHandler);

                Thread.Sleep(5000);

                sEventEnd = false;
                Point XSDsNodePnt = AUIUtilities.GetElementCenterPoint(aeXSDsNode);
                Input.MoveToAndRightClick(XSDsNodePnt);
                Thread.Sleep(2000);

                Point actionsPnt = new Point(XSDsNodePnt.X + 1.2, XSDsNodePnt.Y + 7.2);
                Input.MoveTo(actionsPnt);
                Thread.Sleep(2000);

                Point CreateXSDsPnt = new Point(actionsPnt.X + 155.2, actionsPnt.Y);
                Input.MoveTo(CreateXSDsPnt);
                Thread.Sleep(2000);

                Input.MoveToAndClick(CreateXSDsPnt);
                Thread.Sleep(2000);

                sEventEnd = false;
                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
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

                Console.WriteLine("Searching windows ............wait :40 sec");
                Thread.Sleep(40000);

                Console.WriteLine("Searching windows ..................");
                Thread.Sleep(4000);

                aeWindow = null;
                aeAllWindows = null;
                k = 0;
                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                while (aeWindow == null && mTime.TotalSeconds <= 120)
                {
                    Console.WriteLine("aeWindow[k]="+k);
                    k++;
                    aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        if (aeAllWindows[i].Current.AutomationId.Equals("MainForm"))
                        {
                            aeWindow = aeAllWindows[i];
                            Console.WriteLine("aeWindow[" + i + "]=" + aeWindow.Current.Name);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }

                if (aeWindow == null)
                {
                    sErrorMessage = "MainForm not found";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                    result = ConstCommon.TEST_FAIL;
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

                        k = 0;
                        mStartTime = DateTime.Now;
                        mTime = DateTime.Now - mStartTime;
                        while (aeDocument == null && mTime.TotalSeconds <= 120)
                        {
                            Console.WriteLine("aeDocument[k]="+k);
                            k++;
                            aeAllDocuments = aeTextBoxPanel.FindAll(TreeScope.Children, cDocs);
                            Thread.Sleep(3000);
                            for (int i = 0; i < aeAllDocuments.Count; i++)
                            {
                                Console.WriteLine("aeDocument[" + i + "] name length=" + aeAllDocuments[i].Current.Name.Length);
                                if (aeAllDocuments[i].Current.Name.Length > 20 )
                                {
                                    aeDocument = aeAllDocuments[i];
                                    Console.WriteLine("aeDocument[" + i + "]=" + aeDocument.Current.Name);
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
                            if (aeDocument.Current.Name.IndexOf("successfully") > 0)
                            {
                                TestCheck = ConstCommon.TEST_PASS;
                            
                            }
                            else
                            {
                                sErrorMessage = "=== Create of XSD finished with error(s)=========:" ;
                                Console.WriteLine(sErrorMessage + aeDocument.Current.Name);
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
                Console.WriteLine("Searching main windows ..................");
                AutomationElement aeTreeView = null;
                AutomationElement aeComputerNameNode = null;
                AutomationElement aeProjectsNode = null;
                #region //Search test project and delte this project if it exist
                AutomationElement aeWindow = StatUtilities.GetMainWindow("MainForm");
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
                                sErrorMessage = "aeComputerNameNode can not expaned " + aeComputerNameNode.Current.Name;
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
                            sErrorMessage = "aeProjectsNode can not expaned " + aeProjectsNode.Current.Name;
                            Console.WriteLine(sErrorMessage);
                        }
                    }
                }

                Thread.Sleep(5000);
                // find Test Project tree item from Project Node
                AutomationElement aeTestNode = null;
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
                    if (!StatUtilities.FindAndClickMenuItemOnThisNode(aeWindow, aeProjectsNode, "New project...", ref sErrorMessage)) ;
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
                    aeWindow = StatUtilities.GetMainWindow("MainForm");
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
                #region searching Test project
                AutomationElement aeWindow = StatUtilities.GetMainWindow("MainForm");
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
                                sErrorMessage = "aeComputerNameNode can not expaned " + aeComputerNameNode.Current.Name;
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
                            sErrorMessage = "aeProjectsNode can not expaned " + aeProjectsNode.Current.Name;
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
                    AutomationElement aeActions = StatUtilities.GetMenuItemFromElement(aeWindow, "Actions", 120, ref sErrorMessage);
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
                    AutomationElement aeCreateDB = StatUtilities.GetMenuItemFromElement(aeWindow, "Create database...", 120, ref sErrorMessage);
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

                //Thread.Sleep(60000);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("check output ....");
                    aeWindow = StatUtilities.GetMainWindow("MainForm");
                    if (aeWindow == null)
                    {
                        sErrorMessage = "MainForm not found";
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                    else
                    {
                        // Make sure our window is usable.
                        // WaitForInputIdle will return before the specified time 
                        // if the window is ready.
                        WindowPattern windowPattern = aeWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                        if (false == windowPattern.WaitForInputIdle(60000))
                        {
                            System.Windows.Forms.MessageBox.Show("Object not responding in a timely manner, click OK continue", "CreateDatabase");
                        }
                        Thread.Sleep(2000);

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
                            mStartTime = DateTime.Now;
                            mTime = DateTime.Now - mStartTime;
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
                                if (aeDocument.Current.Name.IndexOf("=== Create of database finished successfully") > 0)
                                {
                                    TestCheck = ConstCommon.TEST_PASS;

                                }
                                else
                                {
                                    sErrorMessage = "=== Create of database finished with error(s)=========:";
                                    Console.WriteLine(sErrorMessage);
                                    TestCheck = ConstCommon.TEST_FAIL;
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
                #region searching Test project
                AutomationElement aeWindow = StatUtilities.GetMainWindow("MainForm");
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
                                sErrorMessage = "aeComputerNameNode can not expaned " + aeComputerNameNode.Current.Name;
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
                            sErrorMessage = "aeProjectsNode can not expaned " + aeProjectsNode.Current.Name;
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
                    AutomationElement aeActions = StatUtilities.GetMenuItemFromElement(aeWindow, "Actions", 120, ref sErrorMessage);
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
                    AutomationElement aeCreateDB = StatUtilities.GetMenuItemFromElement(aeWindow, "Set database...", 120, ref sErrorMessage);
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

                Thread.Sleep(20000);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("check output ....");
                    aeWindow = StatUtilities.GetMainWindow("MainForm");
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
                            mStartTime = DateTime.Now;
                            mTime = DateTime.Now - mStartTime;
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
                                if (aeDocument.Current.Name.IndexOf("=== Set of database finished successfully") > 0)
                                {
                                    TestCheck = ConstCommon.TEST_PASS;

                                }
                                else
                                {
                                    sErrorMessage = "=== Set of database finished with error(s)=========:";
                                    Console.WriteLine(sErrorMessage + aeDocument.Current.Name);
                                    TestCheck = ConstCommon.TEST_FAIL;
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
                #region searching Test project
                AutomationElement aeWindow = StatUtilities.GetMainWindow("MainForm");
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
                                sErrorMessage = "aeComputerNameNode can not expaned " + aeComputerNameNode.Current.Name;
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
                            sErrorMessage = "aeProjectsNode can not expaned " + aeProjectsNode.Current.Name;
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
                    AutomationElement aeActions = StatUtilities.GetMenuItemFromElement(aeWindow, "Actions", 120, ref sErrorMessage);
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
                    AutomationElement aeCreateEtricc = StatUtilities.GetMenuItemFromElement(aeWindow, "Create Etricc layout...", 120, ref sErrorMessage);
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
                    aeWindow = StatUtilities.GetMainWindow("MainForm");
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

                Thread.Sleep(60000);
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region // validate result
                    Console.WriteLine("check output ....");
                    aeWindow = StatUtilities.GetMainWindow("MainForm");
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
                #region searching Test project
                AutomationElement aeWindow = StatUtilities.GetMainWindow("MainForm");
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
                                sErrorMessage = "aeComputerNameNode can not expaned " + aeComputerNameNode.Current.Name;
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
                            sErrorMessage = "aeProjectsNode can not expaned " + aeProjectsNode.Current.Name;
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
                    AutomationElement aeActivate = StatUtilities.GetMenuItemFromElement(aeWindow, "Activate", 120, ref sErrorMessage);
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
                    TestTools.Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_STATISTICS_PARSERCONFIGURATOR);
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
                    TestTools.Utilities.StartProcessWaitForExit(m_SystemDrive + OSVersionInfoClass.ProgramFilesx86FolderName()+
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
        #region OpenGraphicalView
        public static void OpenGraphicalView(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sErrorScreen = false;

            string reportType = "Performance";
            string reportGroup = "Vehicles";
            string reportName = "Status: graphical view";
            string fromDate = "11/25/2010";
            string toDate = "11/27/2010";
            string validateData = "02:01:13";
            
            AutomationElement aeWindow  = null;
            AutomationElement aeTreeView = null;
            AutomationElement aeEtriccNode = null;

            try
            {
                Console.WriteLine(testname+ "clear window : ");
                aeWindow = StatUtilities.ClearMainWindow();
                if (aeWindow == null)
                {
                    sErrorMessage = "Main Window not opend";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    AUICommon.ClearDisplayedScreens(aeWindow);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region// get TreeView
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
                            Thread.Sleep(3000);
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

                #region// Open report screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.FindGraphicalViewlReport(aeTreeView, reportName, ref sErrorMessage))
                    {
                        Console.WriteLine("aeReport is opend  : ");
                    }
                    else
                    {
                        Console.WriteLine("open aeReport is failed: ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                #endregion

                sEventEnd = false;
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);

                while (sEventEnd == false && mTime.TotalSeconds < 30)
                {
                    Thread.Sleep(2000);
                    Console.WriteLine(" time2 is :" + mTime.TotalSeconds);
                    mTime = DateTime.Now - mStartTime;
                }

                Console.WriteLine(" final time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);
                Console.WriteLine(" final time is :" + mTime.TotalMilliseconds + "    errorscreen = " + sErrorScreen);

                if (sErrorScreen)
                    TestCheck = ConstCommon.TEST_FAIL;


                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement,
                   UIErrorEventHandler);


                #region // validate report screen

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.ValidateReportPerformanceVehiclesReport(reportName, fromDate, toDate, validateData, ref sErrorMessage))
                        TestCheck = ConstCommon.TEST_PASS;
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
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
                    string ms = reportName + " window found at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                    Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine(reportName +  "Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OpenPerformanceVehiclesReports
        public static void OpenPerformanceVehiclesReports(string reportName, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + reportName + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, reportName, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sErrorScreen = false;

            AutomationElement aeTreeView = null;
            AutomationElement aeEtriccNode = null;
            string reportType = "Performance";
            string reportGroup = "Vehicles";
            //string reportName = ReportName.PERFORMANCE_VEHICLES_ModeOverview;
            string fromDate = "11/25/2010";
            string toDate = "11/27/2010";
            string validateData = "02:01:13";


            StatUtilities.GetReportPerformanceVehiclesTestData(reportName, ref fromDate, ref toDate, ref validateData );

            Thread.Sleep(5000);
            try
            {
                AutomationElement aeWindow = StatUtilities.GetMainWindow("MainForm");
                if (aeWindow == null)
                {
                    sErrorMessage = "Main Window not opend";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    AUICommon.ClearDisplayedScreens(aeWindow);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region// get TreeView
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

                #region// Open report screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.FindPerformanceFinalReport(aeTreeView, reportType, reportGroup, reportName, ref sErrorMessage))
                    {
                        Console.WriteLine("aeReport is opend  : ");
                    }
                    else
                    {
                        Console.WriteLine("open aeReport is failed: ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }

                    sEventEnd = false;
                    DateTime mStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - mStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);

                    while (sEventEnd == false && mTime.TotalSeconds < 30)
                    {
                        Thread.Sleep(2000);
                        Console.WriteLine(" time2 is :" + mTime.TotalSeconds);
                        mTime = DateTime.Now - mStartTime;
                    }

                    Console.WriteLine(" final time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);
                    Console.WriteLine(" final time is :" + mTime.TotalMilliseconds + "    errorscreen = " + sErrorScreen);

                    if (sErrorScreen)
                        TestCheck = ConstCommon.TEST_FAIL;
                }
                #endregion

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement,
                   UIErrorEventHandler);


                #region // validate report screen
               
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.ValidateReportPerformanceVehiclesReport(reportName, fromDate, toDate, validateData, ref sErrorMessage))
                        TestCheck = ConstCommon.TEST_PASS;
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                #endregion

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, reportName + ":" + sErrorMessage, sOnlyUITest);
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
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine(reportName + " Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OpenPerformanceTransportsReports
        public static void OpenPerformanceTransportsReports(string reportName, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + reportName + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, reportName, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sErrorScreen = false;

            AutomationElement aeTreeView = null;
            AutomationElement aeEtriccNode = null;
            string reportType = "Performance";
            string reportGroup = "Transports";
            //string reportName = ReportName.PERFORMANCE_VEHICLES_ModeOverview;
            string fromDate = "11/25/2010";
            string toDate = "11/27/2010";
            string validateData = "02:01:13";


            StatUtilities.GetReportPerformanceTransportsTestData(reportName, ref fromDate, ref toDate, ref validateData);

            try
            {
                AutomationElement aeWindow = StatUtilities.GetMainWindow("MainForm");
                if (aeWindow == null)
                {
                    sErrorMessage = "Main Window not opend";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    AUICommon.ClearDisplayedScreens(aeWindow);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region// get TreeView
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
                            Thread.Sleep(3000);
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

                #region// Open report screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.FindPerformanceFinalReport(aeTreeView, reportType, reportGroup, reportName, ref sErrorMessage))
                    {
                        Console.WriteLine("aeReport is opend  : ");
                    }
                    else
                    {
                        Console.WriteLine("open aeReport is failed: ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }

                    sEventEnd = false;
                    DateTime mStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - mStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);

                    while (sEventEnd == false && mTime.TotalSeconds < 30)
                    {
                        Thread.Sleep(2000);
                        Console.WriteLine(" time2 is :" + mTime.TotalSeconds);
                        mTime = DateTime.Now - mStartTime;
                    }

                    Console.WriteLine(" final time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);
                    Console.WriteLine(" final time is :" + mTime.TotalMilliseconds + "    errorscreen = " + sErrorScreen);

                    if (sErrorScreen)
                        TestCheck = ConstCommon.TEST_FAIL;
                }
                #endregion

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement,
                   UIErrorEventHandler);


                #region // validate report screen

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.ValidateReportPerformanceVehiclesReport(reportName, fromDate, toDate, validateData, ref sErrorMessage))
                        TestCheck = ConstCommon.TEST_PASS;
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                #endregion

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, reportName + ":" + sErrorMessage, sOnlyUITest);
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
                sErrorMessage = string.Empty;
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine(reportName + " Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OpenPerformanceJobsReports
        public static void OpenPerformanceJobsReports(string reportName, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + reportName + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, reportName, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;

            AutomationElement aeTreeView = null;
            AutomationElement aeEtriccNode = null;
            string reportType = "Performance";
            string reportGroup = "Jobs";
            //string reportName = ReportName.PERFORMANCE_VEHICLES_ModeOverview;
            string fromDate = "11/25/2010";
            string toDate = "11/27/2010";
            string validateData = "02:01:13";


            StatUtilities.GetReportPerformanceTransportsTestData(reportName, ref fromDate, ref toDate, ref validateData);

            Thread.Sleep(5000);
            try
            {
                AutomationElement aeWindow = StatUtilities.GetMainWindow("MainForm");
                if (aeWindow == null)
                {
                    sErrorMessage = "Main Window not opend";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    AUICommon.ClearDisplayedScreens(aeWindow);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region// get TreeView
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
                            Thread.Sleep(3000);
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

                #region// Open report screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.FindPerformanceFinalReport(aeTreeView, reportType, reportGroup, reportName, ref sErrorMessage))
                    {
                        Console.WriteLine("aeReport is opend  : ");
                    }
                    else
                    {
                        Console.WriteLine("open aeReport is failed: ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }

                    sEventEnd = false;
                    DateTime mStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - mStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);

                    while (sEventEnd == false && mTime.TotalSeconds < 30)
                    {
                        Thread.Sleep(2000);
                        Console.WriteLine(" time2 is :" + mTime.TotalSeconds);
                        mTime = DateTime.Now - mStartTime;
                    }

                    Console.WriteLine(" final time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);
                    Console.WriteLine(" final time is :" + mTime.TotalMilliseconds + "    errorscreen = " + sErrorScreen);

                    if (sErrorScreen)
                        TestCheck = ConstCommon.TEST_FAIL;

                }
                #endregion

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement,
                   UIErrorEventHandler);


                #region // validate report screen

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    bool retry = true;
                    while (retry == true)
                    {
                        try
                        {
                            if (StatUtilities.ValidateReportPerformanceVehiclesReport(reportName, fromDate, toDate, validateData, ref sErrorMessage))
                            {
                                TestCheck = ConstCommon.TEST_PASS;
                                retry = false;
                            }
                            else
                            {
                                Console.WriteLine(sErrorMessage);
                                TestCheck = ConstCommon.TEST_FAIL;
                                retry = false;
                            }
                        }
                        catch (System.Windows.Automation.ElementNotAvailableException ex)
                        {
                            retry = true;
                            Epia3Common.WriteTestLogMsg(slogFilePath, reportName + ":" + ex.ToString(), sOnlyUITest);
                            
                        }
                    }
                }
                #endregion

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, reportName + ":" + sErrorMessage, sOnlyUITest);
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
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine(reportName + " Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
            }
        }
        #endregion
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region OpenAnalysisReports
        public static void OpenAnalysisReports(string reportName, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + reportName + " ===");
            Epia3Common.WriteTestLogTitle(slogFilePath, reportName, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            TestCheck = ConstCommon.TEST_PASS;
            sErrorScreen = false;

            AutomationElement aeTreeView = null;
            AutomationElement aeEtriccNode = null;
            string reportType = "Analysis";
            //string reportName = ReportName.PERFORMANCE_VEHICLES_ModeOverview;
            string fromDate = "11/25/2010";
            string toDate = "11/27/2010";
            string validateData = "02:01:13";

            StatUtilities.GetReportAnalysisTestData(reportName, sCurrentProject, ref fromDate, ref toDate, ref validateData);

            Thread.Sleep(5000);
            try
            {
                AutomationElement aeWindow = StatUtilities.GetMainWindow("MainForm");
                if (aeWindow == null)
                {
                    sErrorMessage = "Main Window not opend";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                }
                else
                    AUICommon.ClearDisplayedScreens(aeWindow);

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    #region// get TreeView
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
                            Thread.Sleep(3000);
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

                #region// Open report screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.FindAnalysisFinalReport(aeTreeView, reportType, reportName, ref sErrorMessage))
                    {
                        Console.WriteLine("aeReport is opend  : ");
                    }
                    else
                    {
                        Console.WriteLine("open aeReport is failed: ");
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }

                    sEventEnd = false;
                    DateTime mStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - mStartTime;
                    Console.WriteLine(" time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);

                    while (sEventEnd == false && mTime.TotalSeconds < 30)
                    {
                        Thread.Sleep(2000);
                        Console.WriteLine(" time2 is :" + mTime.TotalSeconds);
                        mTime = DateTime.Now - mStartTime;
                    }

                    Console.WriteLine(" final time is :" + mTime.TotalMilliseconds + "    testcheck = " + TestCheck);
                    Console.WriteLine(" final time is :" + mTime.TotalMilliseconds + "    errorscreen = " + sErrorScreen);

                    if (sErrorScreen)
                        TestCheck = ConstCommon.TEST_FAIL;

                }
                #endregion

                Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement,
                   UIErrorEventHandler);

                #region // validate report screen
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (StatUtilities.ValidateReportPerformanceVehiclesReport(reportName, fromDate, toDate, validateData, ref sErrorMessage))
                        TestCheck = ConstCommon.TEST_PASS;
                    else
                    {
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                    }
                }
                //else
                //    System.Windows.Forms.MessageBox.Show("Error:" + sErrorMessage);
                #endregion

                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    result = ConstCommon.TEST_FAIL;
                    Console.WriteLine(sErrorMessage);
                    Epia3Common.WriteTestLogFail(slogFilePath, reportName + ":" + sErrorMessage, sOnlyUITest);
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    result = ConstCommon.TEST_PASS;
                    sErrorMessage = string.Empty;
                    string ms = reportName + " window validation OK at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(ms);
                    Epia3Common.WriteTestLogPass(slogFilePath, reportName, sOnlyUITest);
                }
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_FAIL;
                Console.WriteLine(reportName + " Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                Thread.Sleep(3000);
            }
        }
        #endregion
        #endregion Test Cases ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        static public bool UninstallEtricc5()
        {
            bool status = false;
            AutomationElement element = null;

            System.Windows.Automation.Condition cWindows = new PropertyCondition(
                      AutomationElement.ControlTypeProperty, ControlType.Window);

            AutomationElementCollection aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
            Thread.Sleep(3000);
            for (int i = 0; i < aeAllWindows.Count; i++)
            {
                Console.WriteLine("aeWindow[" + i + "]=" + aeAllWindows[i].Current.Name);
                if (aeAllWindows[i].Current.Name.IndexOf("Program") >= 0)
                {
                    element = aeAllWindows[i];
                    break;
                }
            }

            if (element == null)    // Programs and Features panel not found:
            {
                Console.WriteLine("Programs and Features panel not found: ");
                Thread.Sleep(3000);
                status = false;
                return status;
            }

            AutomationElement rootElement = AutomationElement.RootElement;
            //string uninstallWindowName = "Programs and Features";
            string uninstallWindowName = element.Current.Name;
            string sgrid = "Folder View";

            AutomationElement aeEtriccCore = null;
            string sYesButtonName = "Yes";
            string sCloseButtonName = "Close";

            #region // Uninstall Etricc

            // (1) Programs and Features main window
            Console.WriteLine("Programs and Features Main Form found ... Welcom Main window");
            Console.WriteLine("Searching programs item element...");
            AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(element, sgrid);
            if (aeGridView != null)
                Console.WriteLine("Gridview found...");
            
            System.Threading.Thread.Sleep(2000);

            // Set a property condition that will be used to find the control.
            System.Windows.Automation.Condition c = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);

            AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children, c);
                
            System.Threading.Thread.Sleep(5000);
            
            Console.WriteLine("Programs count ..." + aeProgram.Count);
            for (int i = 0; i < aeProgram.Count; i++)
            {

                Console.WriteLine("programs name: " + aeProgram[i].Current.Name);
                if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                    && aeProgram[i].Current.Name.IndexOf("Statistics") < 0
                    && aeProgram[i].Current.Name.IndexOf("Shell") < 0
                    && aeProgram[i].Current.Name.IndexOf("Playback") < 0
                    && aeProgram[i].Current.Name.IndexOf("HostTest") < 0)
                    aeEtriccCore = aeProgram[i];
            }

            if (aeEtriccCore == null)    // Etricc Core not in Programs list
            {
                Console.WriteLine("No Etricc Core name: ");
                AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(element, sCloseButtonName);
                if (btnClose != null)
                {   // (2) Components
                    //UnInstalled = true;
                    AUIUtilities.ClickElement(btnClose);
                    status = true;
                    return status;
                }
            }
            else
            {
                Console.WriteLine("Etricc Core name: "+aeEtriccCore.Current.Name);
                string x = aeEtriccCore.Current.Name;
                AutomationElement dialogElement = null;

                InvokePattern pattern = aeEtriccCore.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                pattern.Invoke();
                System.Threading.Thread.Sleep(20000);

                DateTime startTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - startTime;
                while (dialogElement == null && mTime.TotalSeconds < 300)
                {
                    System.Threading.Thread.Sleep(8000);
                    aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        Console.WriteLine("aeWindow[" + i + "].name=" + aeAllWindows[i].Current.Name);
                        Console.WriteLine("aeWindow[" + i + "].automationId=" + aeAllWindows[i].Current.AutomationId);
                        if (aeAllWindows[i].Current.Name.StartsWith("E'tricc") && aeAllWindows[i].Current.AutomationId.Length < 20)
                        {
                            dialogElement = aeAllWindows[i];
                            break;
                        }
                    }
                    mTime = DateTime.Now - startTime;
                }


                if (dialogElement != null)
                {
                    AutomationElement aeTitleBar =
                            AUIUtilities.FindElementByID("TitleBar", dialogElement);

                    Point pt1 = new Point(
                        (aeTitleBar.Current.BoundingRectangle.Left + aeTitleBar.Current.BoundingRectangle.Right) / 2,
                        (aeTitleBar.Current.BoundingRectangle.Bottom + aeTitleBar.Current.BoundingRectangle.Top) / 2);

                    Point newPt1 = new Point(pt1.X - 500, pt1.Y - 400);
                    Input.MoveTo(pt1);

                    System.Threading.Thread.Sleep(1000);
                    Input.SendMouseInput(pt1.X, pt1.Y, 0, SendMouseInputFlags.LeftDown | SendMouseInputFlags.Absolute);

                    System.Threading.Thread.Sleep(1000);
                    Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.Move | SendMouseInputFlags.Absolute);
                    //Input.MoveTo(newPt1);

                    System.Threading.Thread.Sleep(1000);
                    Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);

                    Console.WriteLine("dlalog element moved =");
                    System.Threading.Thread.Sleep(10000);
                    Console.WriteLine("----- Windows dialog window still open ...");

                    startTime = DateTime.Now;
                    mTime = DateTime.Now - startTime;
                  
                    AutomationElement aeRegisterdialog = null;
                    AutomationElement aeYesButton = null;
                    aeRegisterdialog = null;
                    while (aeRegisterdialog == null && mTime.TotalSeconds < 120)
                    {
                        Console.WriteLine("----- WAIT UNTIL Windows RegistryKeysDialog open ...");
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
                    }
                    
                    // wait until application uninstalled
                    startTime = DateTime.Now;
                    mTime = DateTime.Now - startTime;
                    bool hasApplication = IsApplicationInstalled("EtriccCore");
                    while (hasApplication == true && mTime.TotalMilliseconds < 120000)
                    {
                        System.Threading.Thread.Sleep(8000);
                        mTime = DateTime.Now - startTime;
                        if (mTime.TotalMilliseconds > 120000)
                        {
                            System.Windows.Forms.MessageBox.Show("Uninstall EtriccCore run timeout " + mTime.TotalMilliseconds);
                            break;
                        }
                        hasApplication = IsApplicationInstalled("EtriccCore");
                    }

                    System.Threading.Thread.Sleep(2000);
                    AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(element, sCloseButtonName);
                    if (btnClose != null)
                    {
                        AUIUtilities.ClickElement(btnClose);
                        status = true;
                    }
                }
                
            }
            #endregion
           
            return status;
        }
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
                            TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS), sBuildNr);
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
                            TestTools.TfsUtilities.UpdateBuildQualityStatus(logger, uri,
                            TestTools.TfsUtilities.GetProjectName(TestTools.ConstCommon.ETRICCSTATISTICS),
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

                            FileManipulation.UpdateStatusInTestInfoFile(sTestResultFolder, "GUI Tests Failed" + testout, "EtriccStatistics+" + sCurrentPlatform);
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
        static string uninstallWindowName = "Programs and Features";
        static string sgrid = "Folder View";
        #region OnUninstallEtriccUIEvent
        public static void OnUninstallEtriccUIEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnUninstallEtriccUIEvent");
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
                
                AutomationElement aeEtriccCore = null;
                string sYesButtonName = "Yes";
                string sCloseButtonName = "Close";
                
                #region // Uninstall Etricc
               
                        // (1) Programs and Features main window
                        Console.WriteLine("Programs and Features Main Form found ... Welcom Main window");
                        Console.WriteLine("Searching programs item element...");
                        AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(element, sgrid);
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
                            if ( aeProgram[i].Current.Name.StartsWith("E'tricc")
                                && aeProgram[i].Current.Name.IndexOf("Statistics") < 0
                                && aeProgram[i].Current.Name.IndexOf("Shell") < 0
                                && aeProgram[i].Current.Name.IndexOf("Playback") < 0
                                && aeProgram[i].Current.Name.IndexOf("HostTest") < 0 )
                                aeEtriccCore = aeProgram[i];
                        }

                        if (aeEtriccCore == null)    // Etricc Core not in Programs list
                        {
                            Console.WriteLine("No Etricc Core name: ");
                            AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(element, sCloseButtonName);
                            if (btnClose != null)
                            {   // (2) Components
                                //UnInstalled = true;
                                AUIUtilities.ClickElement(btnClose);
                            }
                        }
                        else
                        {
                            Console.WriteLine("Etricc Core name: " + aeEtriccCore.Current.Name);
                            string x = aeEtriccCore.Current.Name;

                            InvokePattern pattern = aeEtriccCore.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                            pattern.Invoke();
                            System.Threading.Thread.Sleep(2000);
                       
                            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName);
                            AutomationElement dialogElement = element.FindFirst(TreeScope.Children, condition);
                            if (dialogElement != null)
                            {
                                    AutomationElement aeTitleBar =
                                            AUIUtilities.FindElementByID("TitleBar", dialogElement);
                           
                                Point pt1 = new Point(
                                    (aeTitleBar.Current.BoundingRectangle.Left +aeTitleBar.Current.BoundingRectangle.Right) / 2,
                                    (aeTitleBar.Current.BoundingRectangle.Bottom + aeTitleBar.Current.BoundingRectangle.Top) / 2);

                                Point newPt1 = new Point(pt1.X+100, pt1.Y+100);
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

                                        if  (installer2Element == null )
                                            Console.WriteLine("uninstaller dialog closed.");    

                                    }

                                    // wait until application uninstalled
                                    startTime = DateTime.Now;
                                    mTime = DateTime.Now - startTime;
                                    bool hasApplication = IsApplicationInstalled("EtriccCore");
                                    while ( hasApplication == true && mTime.TotalMilliseconds < 120000)
                                    {                  
                                        System.Threading.Thread.Sleep(8000);
                                        mTime = DateTime.Now - startTime;
                                        if (mTime.TotalMilliseconds > 120000)
                                        {
                                            System.Windows.Forms.MessageBox.Show("Uninstall EtriccCore run timeout " + mTime.TotalMilliseconds);
                                            break;
                                        }
                                        hasApplication = IsApplicationInstalled("EtriccCore");
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
                    sErrorMessage = "aeComboBoxComputerNameEdit not found";
                    TestCheck = ConstCommon.TEST_FAIL;
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
        #region OnCreateXSDsUIEvent
        public static void OnCreateXSDsUIEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnCreateXSDsUIEvent");
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
            string str = string.Format("OnCreateXSDsUIEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
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
            else if (name.StartsWith("Remote file browser"))
            {
                Console.WriteLine("Remote file browser is opend -------------- : " + System.DateTime.Now);
                string treeViewId = "m_TreeView";
                AutomationElement aeTreeView = null;
                AutomationElement aeComputerNameNode = null;
                DateTime sTime = DateTime.Now;
                AUIUtilities.WaitUntilElementByIDFound(element, ref aeTreeView, treeViewId, sTime, 60);
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
                    aeComputerNameNode = null;
                    TreeWalker walker = TreeWalker.ControlViewWalker;
                    AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
                    while (elementNode != null)
                    {
                        Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
                        if (elementNode.Current.Name.ToLower().Equals(PCName.ToLower()))
                        {
                            //Input.MoveTo(elementNode);
                            aeComputerNameNode = elementNode;
                            Console.WriteLine("Computer node name found , it is: " + aeComputerNameNode.Current.Name);
                            TestCheck = ConstCommon.TEST_PASS;
                            break;
                        }
                        Thread.Sleep(3000);
                        elementNode = walker.GetNextSibling(elementNode);
                    }
                    //return aeNodeLink;
                    if (aeComputerNameNode == null)
                    {
                        sErrorMessage = "Computer node not found, ";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    else
                        TestCheck = ConstCommon.TEST_PASS;
                }

                // find C Disk
                AutomationElement aeCdiskNode = null;
                // find xsd tree item from computernode
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\n=== Find C: (Local Disk ) node ===");
                    Thread.Sleep(3000);
                    System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();

                    aeCdiskNode = TestTools.AUICommon.WalkEnabledElements(aeComputerNameNode, treeNode, "C: (Local Disk )");
                    if (aeCdiskNode == null)
                    {
                        sErrorMessage = string.Empty;
                        Console.WriteLine("\n=== C: (Local Disk ) node NOT Exist ===");
                        Input.MoveToAndDoubleClick(aeComputerNameNode.GetClickablePoint());
                        Thread.Sleep(9000);
                        aeCdiskNode = TestTools.AUICommon.WalkEnabledElements(aeTreeView, treeNode, "C: (Local Disk )");
                    }
                    else
                    {
                        Console.WriteLine("\n=== C: (Local Disk ) node Exist ===");
                        try
                        {
                            ExpandCollapsePattern ep = (ExpandCollapsePattern)aeCdiskNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                            ep.Expand();
                            Thread.Sleep(9000);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("XsdNode can not expaned: " + aeCdiskNode.Current.Name);
                        }
                    }

                    if (aeCdiskNode == null)
                    {
                        sErrorMessage = "aeCdiskNode node not found, ";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    else
                        TestCheck = ConstCommon.TEST_PASS;
                }

                // find Program files
                AutomationElement aeFrogramFilesNode = null;
                string programFilesFolderName = TestTools.OSVersionInfoClass.ProgramFilesx86FolderName();
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\n=== Find "+programFilesFolderName+" node ===");
                    Thread.Sleep(3000);
                    System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();

                    aeFrogramFilesNode = TestTools.AUICommon.WalkEnabledElements(aeCdiskNode, treeNode, programFilesFolderName);
                    if (aeFrogramFilesNode == null)
                    {
                        sErrorMessage = string.Empty;
                        Console.WriteLine("\n=== "+programFilesFolderName+" node NOT Exist ===");
                        Input.MoveToAndDoubleClick(aeCdiskNode.GetClickablePoint());
                        Thread.Sleep(9000);
                        aeFrogramFilesNode = TestTools.AUICommon.WalkEnabledElements(aeCdiskNode, treeNode, programFilesFolderName);
                    }
                    else
                    {
                        try
                        {
                            Console.WriteLine("\n=== "+programFilesFolderName+" node Exist ===");
                            ExpandCollapsePattern ep = (ExpandCollapsePattern)aeFrogramFilesNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                            ep.Expand();
                            Thread.Sleep(9000);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("XsdNode can not expaned: " + aeFrogramFilesNode.Current.Name);
                        }
                        //Input.MoveToAndClick(aeNode);
                    }

                    if (aeFrogramFilesNode == null)
                    {
                        sErrorMessage = "aeFrogramFilesNode node not found, ";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    else
                    {
                        //ScrollPattern scrollPattern = GetScrollPattern(element);
                        ScrollPattern scrollPattern = (ScrollPattern)aeTreeView.GetCurrentPattern(ScrollPattern.Pattern);
                        if (scrollPattern.Current.VerticallyScrollable)
                        {
                            while (aeFrogramFilesNode.Current.IsOffscreen)
                            {
                                scrollPattern.ScrollVertical(ScrollAmount.LargeIncrement);
                            }
                        }
                        
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                }

                // find Egemin
                AutomationElement aeEgeminNode = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\n=== Find " + "Egemin" + " node ===");
                    Thread.Sleep(3000);
                    System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
                    aeEgeminNode = TestTools.AUICommon.WalkEnabledElements(aeFrogramFilesNode, treeNode, "Egemin");
                    if (aeEgeminNode == null)
                    {
                        sErrorMessage = string.Empty;
                        Console.WriteLine("\n=== " + "Egemin" + " node NOT Exist ===");
                        Input.MoveToAndDoubleClick(aeFrogramFilesNode.GetClickablePoint());
                        Thread.Sleep(9000);
                        aeEgeminNode = TestTools.AUICommon.WalkEnabledElements(aeFrogramFilesNode, treeNode, "Egemin");
                    }
                    else
                    {
                        try
                        {
                            Console.WriteLine("\n=== " + "Egemin" + " node Exist ===");
                            ExpandCollapsePattern ep = (ExpandCollapsePattern)aeEgeminNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                            ep.Expand();
                            Thread.Sleep(9000);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("XsdNode can not expaned: " + aeEgeminNode.Current.Name);
                        }

                    }

                    if (aeEgeminNode == null)
                    {
                        sErrorMessage = "aeEgeminNode node not found, ";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    else
                    {
                        ScrollPattern scrollPattern = (ScrollPattern)aeTreeView.GetCurrentPattern(ScrollPattern.Pattern);
                        if (scrollPattern.Current.VerticallyScrollable)
                        {
                            while (aeEgeminNode.Current.IsOffscreen)
                            {
                                scrollPattern.ScrollVertical(ScrollAmount.LargeIncrement);
                            }
                        }
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                }

                // find Etricc Server
                AutomationElement aeEtriccServerNode = null;
                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    Console.WriteLine("\n=== Find " + "Etricc Server" + " node ===");
                    Thread.Sleep(3000);
                    System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
                    aeEtriccServerNode = TestTools.AUICommon.WalkEnabledElements(aeEgeminNode, treeNode, "Etricc Server");
                    if (aeEtriccServerNode == null)
                    {
                        sErrorMessage = string.Empty;
                        Console.WriteLine("\n=== " + "Etricc Server" + " node NOT Exist ===");
                        Input.MoveToAndDoubleClick(aeEgeminNode.GetClickablePoint());
                        Thread.Sleep(9000);
                        aeEtriccServerNode = TestTools.AUICommon.WalkEnabledElements(aeEgeminNode, treeNode, "Etricc Server");
                    }
                    else
                    {
                        Console.WriteLine("\n=== " + "Etricc Server" + " node Exist ===");
                        //Input.MoveToAndClick(aeNode);
                    }

                    if (aeEtriccServerNode == null)
                    {
                        sErrorMessage = "aeEtriccServerNode node not found, ";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                        Input.MoveToAndClick(aeEtriccServerNode);
                    }
                }
                Thread.Sleep(3000);

                string listViewId = "m_ListView";
                AutomationElement aeListView = null;
                AutomationElement aeWCSdll = null;
                sTime = DateTime.Now;
                AUIUtilities.WaitUntilElementByIDFound(element, ref aeListView, listViewId, sTime, 60);
                if (aeListView == null)
                {
                    sErrorMessage = "aeListView not found name";
                    Console.WriteLine(sErrorMessage);
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                    return;
                }
                else
                {
                    Console.WriteLine("List view found  .........");
                    Thread.Sleep(5000);
                    // Set a property condition that will be used to find the control.
                    System.Windows.Automation.Condition c = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.DataItem);

                    AutomationElementCollection aeAllItems = aeListView.FindAll(TreeScope.Children, c);

                    Console.WriteLine("All items count ..." + aeAllItems.Count);
                    for (int i = 0; i < aeAllItems.Count; i++)
                    {
                        Console.WriteLine("item name: " + aeAllItems[i].Current.Name);
                        if (aeAllItems[i].Current.Name.StartsWith("Egemin.EPIA.WCS.dll"))
                            aeWCSdll = aeAllItems[i];
                    }

                    Thread.Sleep(3000);
                    if (aeWCSdll == null)
                    {
                        sErrorMessage = " aeWCSdll not found, ";
                        Console.WriteLine(sErrorMessage);
                        TestCheck = ConstCommon.TEST_FAIL;
                        sEventEnd = true;
                        return;
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                        Console.WriteLine("aeWCSdll found  .........");

                        ScrollPattern scrollPattern = (ScrollPattern)aeListView.GetCurrentPattern(ScrollPattern.Pattern);
                        if (scrollPattern.Current.VerticallyScrollable)
                        {
                            while (aeWCSdll.Current.IsOffscreen)
                            {
                                scrollPattern.ScrollVertical(ScrollAmount.LargeIncrement);
                            }
                        }

                        //Input.MoveTo(aeWCSdll);
                        Thread.Sleep(5000);
                        bool select = false; //Utilities.SelectItemFromList("nl", aeCombo);
                        if (TestCheck == ConstCommon.TEST_PASS)
                        {
                            SelectionPattern selectPattern =
                               aeListView.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                            AutomationElement item
                                = AUIUtilities.FindElementByName("Egemin.EPIA.WCS.dll", aeListView);
                            if (item != null)
                            {
                                Console.WriteLine("Egemin.EPIA.WCS.dll item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                Thread.Sleep(2000);

                                SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemPattern.Select();
                                select = true;
                            }
                            else
                            {
                                Console.WriteLine("Finding Egemin.EPIA.WCS.dll item nl failed");
                                sErrorMessage = "Finding Egemin.EPIA.WCS.dll item nl failed";
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
                    }
                }

                // Check selected dll
                string FilenameEditBoxId = "m_TextBoxFileName";
                // check get value Egemin.EPIA.WCS.dll exist

                // check select button is enable

                Thread.Sleep(15000);

                AutomationElement aeSelectButton = null;
                string BtnConnectId = "m_BtnSelect";
                Thread.Sleep(3000);
                DateTime mAppTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mAppTime;
                while (aeSelectButton == null && mTime.Minutes < 5)
                {
                    Console.WriteLine("Find Application aeSelectButton : " + System.DateTime.Now);
                    aeSelectButton = AUIUtilities.FindElementByID(BtnConnectId, element);
                    Console.WriteLine("Application aeSelectButton : " + System.DateTime.Now);
                    mTime = DateTime.Now - mAppTime;
                    Thread.Sleep(20000);
                    Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
                }

                if (aeSelectButton == null)
                {
                    sErrorMessage = "aeConnectButton not found";
                    TestCheck = ConstCommon.TEST_FAIL;
                    sEventEnd = true;
                    return;
                }
                else
                {
                    Thread.Sleep(500);
                    TestTools.AUIUtilities.ClickElement(aeSelectButton);
                    Thread.Sleep(2000);
                    sEventEnd = true;
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
        #region Excel ------------------------------------------------------------------------------------------------
        public static void WriteResult(int result, int counter, string name,
            Excel.Worksheet sheet, string errorMSG)
        {
            string time = System.DateTime.Now.ToString("HH:mm:ss");
            xSheet.Cells[counter + 2 + 9, 1] = time;
            xSheet.Cells[counter + 2 + 9, 2] = name;
            xSheet.Cells[counter + 2 + 9, 3] = errorMSG;

            xRange = sheet.get_Range("A" + (Counter + 2 + 9), "A" + (Counter + 2 + 9));
            xRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            xRange.Columns.AutoFit();

            xRange = sheet.get_Range("C" + (Counter + 2 + 9), "C" + (Counter + 2 + 9));
            xRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            xRange.Columns.AutoFit();
           
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
                outFilename = currentPlatform + "Debug-" + outFilename + "-" + PCName;
            else
                outFilename = currentPlatform + "Release-" + outFilename + "-" + PCName;

            if (args[10].ToLower().StartsWith("false"))
                outFilename = "Manual-" + currentPlatform + outFilename;

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
