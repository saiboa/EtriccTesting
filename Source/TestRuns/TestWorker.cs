using System;
using System.Collections;
using System.Collections.Specialized;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Egemin.EPIA;
using Egemin.EPIA.Core.Definitions;
using Egemin.EPIA.WCS.Resources;
using Egemin.EPIA.WCS.Storage;
using Egemin.EPIA.WCS.Transportation;
using Microsoft.Office.Interop.Excel;
using Application = Egemin.EPIA.Core.Application;
using Project = Egemin.EPIA.WCS.Core.Project;

namespace TestRuns
{
    /// <summary>
    /// Summary description for TestWorker.
    /// </summary>
    public class TestWorker : Repeater
    {
        #region Constants

        private const string TOOLNAME = "TestScenarios of Eurobaltic AutoRuns";
        private const string VERSION = "1.0";
        private const string PC_TEAMTESTETRICC5 = "TEAMTESTETRICC5";

        // Test related files
        public static string TEST_CENTER_LOG_FILE = "TestEurobalticLog.txt";
        public static string TEST_RESULT_FILE_SURFIX = "-Test-Eurobaltic";

        #endregion// —— Enums/Constants ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        #region Fields

        // project params
        private static string[] xlsHeader = new string[10];
        private static string[,] xlsBody = new string[100,4];
        // testinput =================================================
        private static Agv[] sTestAgvs;
        //static Agv      sTestAgv = null;
        //static Agv      sTestAgv2 = null;

        private static string sJobType = string.Empty;

        private static string sSourceID = string.Empty;
        private static string sSource2ID = string.Empty;
        private static string sSource3ID = string.Empty;

        private static string sDestinationID = string.Empty;
        private static string sDestination2ID = string.Empty;
        private static string sDestination3ID = string.Empty;

        private static string sLocationID = string.Empty;
        private static string sLocation2ID = string.Empty;
        private static string sLocation3ID = string.Empty;

        private static string sStationID = string.Empty;
        private static string sScheduleID = string.Empty;


        private static string sGroupID = string.Empty;
        private static string sGroup2ID = string.Empty;

        private static string sTransType = string.Empty;

        private static int sWaitTime = 2;

        //static Computer mc;
        //==========================================================
        // test params
        private static Logger mLogger;

        private static int sProjectID = TestConstants.PROJECT_EUROBALTIC;
        private static int sTotalTestCounter;
        private static int sTotalPassCounter;
        private static int sTotalFailCounter;
        private static int sTotalExceptionCounter;
        private static int sTotalUntestedCounter;
        private static int sCounter;
        private static int sTestResult = TestConstants.TEST_UNDEFINED;
        private static string sTSName = string.Empty;
        private static DateTime sTestStartTime = DateTime.Now;
        private static DateTime sTestEndTime = DateTime.Now;
        //static	DateTime	sJobStartTime			= DateTime.Now;
        private static bool sTestMonitorUsed;

        private static string sRunStatus;
        private static string sRunStatus_prev;

        private static string sWaitID = string.Empty;
        private static string sBattID = string.Empty;
        private static string sAgvCurrentLSID = string.Empty;

        private static string sJobID = string.Empty;
        private static string sJob2ID = string.Empty;

        private static int sTestID;
        private static int sTestID_prev;

        private static string sTextTestData = string.Empty;

        private static string sXLSPath = string.Empty;

        // test status

        private static Agv sChargingAgv;
        private static Agv sWaitAgv;
        private static Agv sQueueAgv;
        private static Hashtable sPrjLayout = new Hashtable();
        private static Hashtable sAgvsInitialID = new Hashtable();
        private static Hashtable sAgvsDefaultDropID = new Hashtable();
        private static Hashtable sTestInputParams = new Hashtable();
        private static string PCName = Environment.MachineName;
        private static int sTestCaseNameSend;
        private int mCleanUPStatus = TestConstants.CLEANUP_NOT_STARTED;
        private int mJobStatus = TestConstants.JOB_NOT_STARTED;

        private string mMsg;
        private int mRunStatus = TestConstants.CONTINUE;
        private Agv mTestAgv;
        private string mTestAgvBattID = string.Empty;
        private string mTestAgvParkID = string.Empty;
        private int mTestStatus = TestConstants.TEST_NOT_STARTED;
        private TimeSpan mTime;
        private Application m_Application;
        private Job m_Job;
        private Job m_Job2;
        private Job m_Job3;

        private StringCollection m_Logging = new StringCollection();
        private Project m_Project;
        private Transport m_Transport;
        private Transport m_Transport2;
        private Transport m_Transport3;
        private TransportManager m_TransportManager;
        private Transport.STATE m_TransportState;

        private WeekPlan m_Wp = new WeekPlan();
        private WeekPlan m_Wp2 = new WeekPlan();
        private SqlConnection myConnection;
        private string sScreenTestTitle = string.Empty;
        private string[] sTestCaseName = new string[100];
        private TestConstants.TESTINFO testinfo;
        private Microsoft.Office.Interop.Excel.Application xApp;
        private Workbook xBook;
        private Workbooks xBooks;
        private Range xRange;
        private Worksheet xSheet;

        [DllImport("mpr.dll")]
        public static extern int WNetAddConnection2A(
            [MarshalAs(UnmanagedType.LPArray)] NETRESOURCEA[] lpNetResource,
            [MarshalAs(UnmanagedType.LPStr)] string lpPassword,
            [MarshalAs(UnmanagedType.LPStr)] string UserName,
            int dwFlags);

        [DllImport("mpr.dll")]
        public static extern int WNetCancelConnection2A(string sharename, int dwFlags, int fForce);

        /// <summary>
        /// network related struct
        /// </summary>
        public struct NETRESOURCEA
        {
            public int dwDisplayType;
            public int dwScope;
            public int dwType;
            public int dwUsage;
            [MarshalAs(UnmanagedType.LPStr)] public string lpComment;
            [MarshalAs(UnmanagedType.LPStr)] public string lpLocalName;
            [MarshalAs(UnmanagedType.LPStr)] public string lpProvider;
            [MarshalAs(UnmanagedType.LPStr)] public string lpRemoteName;

            public override String ToString()
            {
                String str = "LocalName: " + lpLocalName + " RemoteName: " + lpRemoteName
                             + " Comment: " + lpComment + " lpProvider: " + lpProvider;
                return (str);
            }
        }

        #endregion // —— Fields ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••?

        protected override void OnActivate()
        {
            base.OnActivate();

            string logPath = Path.Combine(TestConstants.TEST_LOG_PATH, TEST_CENTER_LOG_FILE);
            mLogger = new Logger(logPath);

            if (!Directory.Exists(TestConstants.TEST_LOG_PATH))
                Directory.CreateDirectory(TestConstants.TEST_LOG_PATH);

            // if log file exist, empty log file
            if (!File.Exists(logPath))
            {
                // create an empty file
                StreamWriter writer = File.CreateText(logPath);
                writer.Close();
            }
            else // empty log file first
            {
                StreamWriter writer = File.CreateText(logPath);
                writer.Close();
            }

            mLogger.LogMessageToFile("TestWorker onActive  ");

            string testDir = @"C:\EtriccTests";
            TestUtility.GetTestInformation(testDir, ref testinfo);
            mLogger.LogMessageToFile("test information:  " + testinfo.ToString());

            sTestCaseNameSend = 0;
            sTestCaseName[0] = "All";
            sTestCaseName[1] = "TS220001JobParkSemiAutomatic";
            sTestCaseName[2] = "TS220002JobBattSemiAutomatic";
            sTestCaseName[3] = "TS220003JobWaitSemiAutomatic";
            sTestCaseName[4] = "TS220005JobPickSemiAutomatic";
            sTestCaseName[5] = "TS220006JobDropSemiAutomatic";
            sTestCaseName[6] = "TS220034JobFlushing";
            sTestCaseName[7] = "TS220016JobCanceling";
            sTestCaseName[8] = "TJobCancelCurrent";
            sTestCaseName[9] = "TS220019JobExhausted";
            sTestCaseName[10] = "TS220020JobAborting";
            sTestCaseName[11] = "TS220023JobSuspending";
            sTestCaseName[12] = "TS220027JobReleasing";
            sTestCaseName[13] = "TS220024JobSuspendCurrent";
            sTestCaseName[14] = "TS220028JobReleaseCurrent";
            sTestCaseName[15] = "TS220025JobSuspendAll";
            sTestCaseName[16] = "TS220029JobReleaseAll";
            sTestCaseName[17] = "TS221006JobPickViaStation";
            sTestCaseName[18] = "TS221007JobParkViaStation";
            sTestCaseName[19] = "TS221008JobBattViaStation";
            sTestCaseName[20] = "TS221009JobWaitViaStation";
            sTestCaseName[21] = "TS221010JobDropViaStation";
            sTestCaseName[22] = "TS300005TransOrderPick";
            sTestCaseName[23] = "TS300006TransOrderDrop";
            sTestCaseName[24] = "TS300008TransOrderMove";
            sTestCaseName[25] = "TS300003TransOrderWait";
            sTestCaseName[26] = "TS300072-1-TransOrderExcp1";
            sTestCaseName[27] = "TS300043TransOrderMode";
            sTestCaseName[28] = "TS300071TransOrderState";
            //sTestCaseName[29] = "TS300011TransOrderEdit";
            sTestCaseName[29] = "TS300034TransOrderFlush";
            sTestCaseName[30] = "TS300016TransOrderCancel";
            sTestCaseName[31] = "TS300023TransOrderSuspend";
            sTestCaseName[32] = "TS300027TransOrderRelease";
            sTestCaseName[33] = "TS300031TransOrderFinish";
            sTestCaseName[34] = "TS300025TransOrderSuspendAll";
            sTestCaseName[35] = "TS300029TransOrderReleaseAll";
            sTestCaseName[36] = "TS300026TransOrderSuspendAllPending";
            sTestCaseName[37] = "TS300030TransOrderReleaseAllPending";
            sTestCaseName[38] = "TS350057MutexAuto";
            sTestCaseName[39] = "TS300016TransportSourceVia";
            sTestCaseName[40] = "TS300017TransportDestinationVia";
            sTestCaseName[41] = "TS241099WeekPlanBatteryCharge";
            sTestCaseName[42] = "TS241047WeekPlanBatteryChargeDisable";
            sTestCaseName[43] = "TS241048WeekPlanBatteryChargeDisableAll";
            sTestCaseName[44] = "TS241049WeekPlanBatteryChargeDelete";
            sTestCaseName[45] = "TS242099WeekPlanCalibration";
            sTestCaseName[46] = "TS242047WeekPlanCalibrationDisable";
            sTestCaseName[47] = "TS242048WeekPlanCalibrationDisableAll";
            sTestCaseName[48] = "TS242049WeekPlanCalibrationDelete";
            sTestCaseName[49] = "TS200083AgvStop";
            sTestCaseName[50] = "TS200071AgvState";
            sTestCaseName[51] = "TS200037AgvRetire";
            sTestCaseName[52] = "TS200053AgvDeploy";
            sTestCaseName[53] = "TS200023AgvSuspend";
            sTestCaseName[54] = "TS200027AgvRelease";
            sTestCaseName[55] = "TS200012AgvModeRemoved";
            sTestCaseName[56] = "TS200014AgvModeRemovedAll";
            sTestCaseName[57] = "TS200047AgvModeDisable";
            sTestCaseName[58] = "TS200048AgvModeDisableAll";
            sTestCaseName[59] = "TS200058AgvModeSemiAutomatic";
            sTestCaseName[60] = "TS200061AgvModeSemiAutomaticAll";
            sTestCaseName[61] = "TS309306TransPickDeactiveRestart";
            //sTestCaseName[63] = "TS309307TransDropDeactiveRestart";
            sTestCaseName[62] = "TS420047LocationDisable";
            sTestCaseName[63] = "TS420045LocationManual";
            sTestCaseName[64] = "TS460047StationDisable";
            sTestCaseName[65] = "TS400034LoadFlushAndDiscard";
            sTestCaseName[66] = "TS400076LoadDiscard";
            sTestCaseName[67] = "TS305513TransOrderDelay";
            sTestCaseName[68] = "TS305514TransOrderDivert";
            sTestCaseName[69] = "TS304455LocationClosestHighest";
            sTestCaseName[70] = "TS304457GroupHighestPriority";
            sTestCaseName[71] = "TS304456LoadClosestHighest";
            sTestCaseName[72] = "TS305501OrderAssignmentClosest";
            sTestCaseName[73] = "TS305502TransOrderClosest";
            sTestCaseName[74] = "TS305503OrderAssignmentClosestHighest";
            sTestCaseName[75] = "TS305504TransOrderClosestHighest";
            //sTestCaseName[78] = "TS300055TransOrderPriority";
            sTestCaseName[76] = "TS305505TransOrderOldest";
            sTestCaseName[77] = "TS305507SchedulesDeadlockRulesVia";
            sTestCaseName[78] = "TS305508ScheduleBattRulesQueueSimLow";
            sTestCaseName[79] = "TS330056RoutingDynamic";
            //sTestCaseName[82] = "TS300063TransOrderDoublePlay";
            //sTestCaseName[83] = "TS300064DoublePlayTransReleased";
            sTestCaseName[80] = "TS300080TransPickFromGroup";
            sTestCaseName[81] = "TS300081TransDropToGroup";
            sTestCaseName[82] = "TS200080AgvModeSemiToAuto";
            //sTestCaseName[87] = "TS830001DBSQLSERVERStopStart";
            //sTestCaseName[59] = "TS241001WeekPlanBattChargeActiveCheck";
            //sTestCaseName[60] = "TS241002WeekPlanCalibrationActiveCheck";
            //sTestCaseName[68] = "TS305506ScheduleBattRulesQueueWP";
            //sTestCaseName[73] = "TS242099WeekPlanCalibrationMultipleTrigger";

            for (int i = 0; i <= 81; i++)
            {
                xlsBody[i, 0] = "time_" + i;
                xlsBody[i, 1] = "name_" + i;
                xlsBody[i, 2] = "result_" + i;
                xlsBody[i, 3] = "data_" + i;
            }

            int ret = Disconnect("X:");
            if (ret == 0)
            {
                mLogger.LogMessageToFile("Disconnect MAP DRIVE OK:");
            }
            else
                mLogger.LogMessageToFile("Disconnect  DriveMap failed with error code:" + ret);

            ret = OpenDriveMap(@"\\Teamsystem\Team Systems Builds", "X:");
            if (ret == 0)
            {
                mLogger.LogMessageToFile("Create MAP DRIVE OK:");
            }
            else
                mLogger.LogMessageToFile("OpenDriveMap failed with error code:" + ret);

            sScreenTestTitle = sTestCaseName[0];
            for (int i = 1; i < sTestCaseName.Length; i++)
                sScreenTestTitle = sScreenTestTitle + "," + " " + i + "  " + sTestCaseName[i];

            m_Application = Context as Application;
            mLogger.LogMessageToFile("TestWorker  m_Application = " + m_Application.ID);

            m_Project = m_Application.Context as Project;
            mLogger.LogMessageToFile("TestWorker  m_Project = " + m_Project.ID);

            m_TransportManager = m_Project.TransportManager;

            if (m_Project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
            {
                TEST_CENTER_LOG_FILE = "TestEurobalticLog.txt";


                if (testinfo.autoTestMode_15)
                    TEST_RESULT_FILE_SURFIX = "-Test-Eurobaltic";
                else
                    TEST_RESULT_FILE_SURFIX = "-Manual-Eurobaltic";

                sProjectID = TestConstants.PROJECT_EUROBALTIC;
            }
            else if (m_Project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
            {
                TEST_CENTER_LOG_FILE = "TestOpstellingLog.txt";
                TEST_RESULT_FILE_SURFIX = "-Test-Opstelling";
                sProjectID = TestConstants.PROJECT_TESTOPSTELLING;
            }
            else if (m_Project.ID.ToString().ToUpper().StartsWith("DEMO"))
                sProjectID = TestConstants.PROJECT_DEMO;
            else
                throw new Exception("Project not Allowed");

            sPrjLayout.Clear();
            sPrjLayout.Add(TestConstants.PROJECT_EUROBALTIC.ToString(), "EuroBaltic");
            sPrjLayout.Add(TestConstants.PROJECT_TESTOPSTELLING.ToString(), "TestProject");
            sPrjLayout.Add(TestConstants.PROJECT_DEMO.ToString(), "Demo");

            sAgvsInitialID = TestData.GetAgvssAgvsInitialID(m_Project);
            sAgvsDefaultDropID = TestData.GetAgvsDefaultDropID(m_Project);
            sTestAgvs = TestData.GetTestAgvs(m_Project);
            sWaitTime = TestData.GetWaitTime(m_Project);

            // status reset
            mTestStatus = TestConstants.TEST_NOT_STARTED;
            mCleanUPStatus = TestConstants.CLEANUP_NOT_STARTED;
            //mJobStatus		= TestConstants.JOB_NOT_STARTED;
            sTestResult = TestConstants.TEST_UNDEFINED;

            sRunStatus = "StartRunning";
            sRunStatus_prev = "StartRunning";

            sRunStatus_prev = sRunStatus;
            sTestID_prev = sTestID;

            if (sTestID == 0)
                sCounter = 1;
            else
                sCounter = sTestID;

            // restart test agvs
            for (int i = 0; i < sTestAgvs.Length; i++)
            {
                sTestAgvs[i].Restart();
            }
            Thread.Sleep(5000);

            string root = m_Project.BinDir.ToString();

            // Excel file Header 
            // 
            //if (PCName.ToUpper().StartsWith(PC_TEAMTESTETRICC5))
            //{
            xlsHeader[0] = "Test Machine: " + Environment.MachineName;
            xlsHeader[1] = "OS : " + testinfo.oS_12;
            xlsHeader[2] = "OS version: " + Environment.OSVersion;
            xlsHeader[3] = "E'pia version: " + testinfo.epiaDeployPath_2;
            xlsHeader[4] = "Build type:: " + testinfo.buildType_3;
            xlsHeader[5] = "Build Path: " + testinfo.buildInstallScriptDir_7;
            xlsHeader[6] = "Layout: " + testinfo.projectFile_6;
            xlsHeader[7] = "TestTools version: " + testinfo.testToolsVersion_5;
            //}
            /*else
			{
				xApp			= new Excel.Application();
				xBooks          = xApp.Workbooks;
				//xBook			= xBooks.Add(XlWBATemplate.xlWBATWorksheet);
				xBook           = xBooks.Add( Type.Missing );
				xSheet          = (Excel.Worksheet)xBook.Worksheets[1];
				xApp.Visible = testinfo.excelShow_11;            
				xApp.Interactive = true;
				TestUtility.AddTestInfoToExcel(ref xSheet, testinfo);
			}*/
            //if (info != null && info.Length > 0 )
            //{
            mLogger.LogMessageToFile("BinDir  " + m_Project.BinDir);
            mLogger.LogMessageToFile("TestFile Directory  " + testinfo.testDirectory_13);
            mLogger.LogMessageToFile("*****************************************************************");
            mLogger.LogMessageToFile("------    " + TOOLNAME + " - " + VERSION + " -------   start up  -----*");
            mLogger.LogMessageToFile("*****************************************************************");
            mLogger.LogMessageToFile("******		Test Machine:\t" + Environment.MachineName + "\t *");
            mLogger.LogMessageToFile("******		OS:\t" + Environment.OSVersion + "\t  *");
            mLogger.LogMessageToFile("******		Epia Version:\t\t" + testinfo.epiaDeployPath_2 + "\t  *");
            mLogger.LogMessageToFile("*****************************************************************");

            mLogger.LogMessageToFile("------    Test : " + sRunStatus);
            mLogger.LogMessageToFile("------    Test selected: " + sTestID);
            //}

            sTestMonitorUsed = TestUtility.IsTestMonitorRunning();
            if (sTestMonitorUsed)
            {
                m_Project.Facilities["Tests"].Parameters["LogMessage"].ValueAsString = "Project Activated";
                //m_Project.Facilities["Tests"].Parameters["TestTitleArray"].ValueAsString = "-- test --1234567890qbcdefghijklmnopqrstuvwxyz ";
                m_Project.Facilities["Tests"].Parameters["TestTitleArray"].ValueAsString = "Time:" +
                                                                                           DateTime.Now.ToLocalTime().
                                                                                               ToString() + " -- " +
                                                                                           sScreenTestTitle;

                mLogger.LogMessageToFile("------    Test Monitor is running ");
            }

            sTextTestData = string.Empty;
            sTotalTestCounter = 0;
            sTotalPassCounter = 0;
            sTotalFailCounter = 0;
            sTotalExceptionCounter = 0;
            sTotalUntestedCounter = 0;

            m_Project.Facilities["Tests"].Parameters["RunStatus"].ValueAsString = "StartRunning";
        }

        protected override void OnMinute()
        {
            base.OnMinute();
            //m_Project.Facilities["Tests"].Parameters["TestTitleArray"].ValueAsString = "Time:" + System.DateTime.Now.ToLocalTime().ToString() + " -- " + sScreenTestTitle;
        }

        protected override void OnSecond()
        {
            base.OnSecond();
            try
            {
                if (sTestMonitorUsed)
                {
                    sRunStatus = m_Project.Facilities["Tests"].Parameters["RunStatus"].ValueAsString;
                    sTestID = m_Project.Facilities["Tests"].Parameters["TestID"].ValueAsInt;
                    // log if StartRun changed
                    if (!sRunStatus.Equals(sRunStatus_prev))
                    {
                        sRunStatus_prev = sRunStatus;
                        mLogger.LogMessageToFile("------    Test : " + sRunStatus);
                        TestUtility.RemoteLogMessage("------    Test : " + sRunStatus, sTestMonitorUsed, m_Project);
                    }

                    // log if testID changed
                    if (sTestID != sTestID_prev)
                    {
                        sTestID_prev = sTestID;
                        mTestStatus = TestConstants.TEST_NOT_STARTED;
                        mLogger.LogMessageToFile("------    Test selected: " + sTestID);
                        TestUtility.RemoteLogMessage("------    Test selected: " + sTestID, sTestMonitorUsed, m_Project);
                        // if All tests, start with first test
                        if (sTestID == 0)
                            sCounter = 1;
                        else
                            sCounter = sTestID; // test selected  testID
                    }

                    if (sRunStatus.Equals("StopRunning"))
                    {
                        //m_Project.Facilities["Tests"].Parameters["TestTitleArray"].ValueAsString = "time: "+System.DateTime.Now.ToLocalTime().ToString();
                        //m_Project.Facilities["Tests"].Parameters["TestTitleArray"].ValueAsString = System.DateTime.Now.ToLocalTime().ToString() + sScreenTestTitle;
                        return;
                    }

                    if (sTestCaseNameSend < 10)
                    {
                        //string title = m_Project.Facilities["Tests"].Parameters["TestTitleArray"].ValueAsString;
                        //TestUtility.RemoteLogMessage("------    titel start : " + title.Substring(0, 30), sTestMonitorUsed, m_Project);
                        //if (title.StartsWith("Time") == false)
                        //{
                        m_Project.Facilities["Tests"].Parameters["TestTitleArray"].ValueAsString = "Time:" +
                                                                                                   DateTime.Now.
                                                                                                       ToLocalTime().
                                                                                                       ToString() +
                                                                                                   " -- " +
                                                                                                   sScreenTestTitle;
                        sTestCaseNameSend++;
                        //}
                        //else
                        //    sTestCaseNameSend = true;
                    }


                    m_Project.Facilities["Tests"].Parameters["CurrentTestID"].ValueAsInt = sCounter;
                }

                if (sRunStatus.Equals("StartRunning"))
                {
                    switch (sCounter)
                    {
                        case 1:
                            sTSName = "TS220001JobParkSemiAutomatic";
                            TS220001JobParkSemiAutomatic(sTSName);
                            break;
                        case 2:
                            sTSName = "TS220002JobBattSemiAutomatic";
                            TS220002JobBattSemiAutomatic(sTSName);
                            break;
                        case 3:
                            sTSName = "TS220003JobWaitSemiAutomatic";
                            TS220003JobWaitSemiAutomatic(sTSName);
                            break;
                        case 4:
                            sTSName = "TS220005JobPickSemiAutomatic";
                            TS220005JobPickSemiAutomatic(sTSName);
                            break;
                        case 5:
                            sTSName = "TS220006JobDropSemiAutomatic";
                            TS220006JobDropSemiAutomatic(sTSName);
                            break;
                        case 6:
                            sTSName = "TS220034JobFlushing";
                            TestScenario_TS220034JobFlushing(sTSName);
                            break;
                        case 7:
                            sTSName = "TS220016JobCanceling";
                            TestScenario_TS220016JobCanceling(sTSName);
                            break;
                        case 8:
                            sTSName = "TS220017JobCancelCurrent";
                            TestScenario_TS220017JobCancelCurrent(sTSName);
                            break;
                        case 9:
                            sTSName = "TS220019JobExhausted";
                            TestScenario_TS220019JobExhausted(sTSName);
                            break;
                        case 10:
                            sTSName = "TS220020JobAborting";
                            TestScenario_TS220020JobAborting(sTSName);
                            break;
                        case 11:
                            sTSName = "TS220023JobSuspending";
                            TestScenario_TS220023JobSuspending(sTSName);
                            break;
                        case 12:
                            sTSName = "TS220027JobReleasing";
                            TestScenario_TS220027JobReleasing(sTSName);
                            break;
                        case 13:
                            sTSName = "TS220024JobSuspendCurrent";
                            TestScenario_TS220024JobSuspendCurrent(sTSName);
                            break;
                        case 14:
                            sTSName = "TS220028JobReleaseCurrent";
                            TestScenario_TTS220028JobReleaseCurrent(sTSName);
                            break;
                        case 15:
                            sTSName = "TS220025JobSuspendAll";
                            TestScenario_TS220025JobSuspendAll(sTSName);
                            break;
                        case 16:
                            sTSName = "TS220029JobReleaseAll";
                            TestScenario_TS220029JobReleaseAll(sTSName);
                            break;
                        case 17:
                            sTSName = "TS221006JobPickViaStation";
                            TestScenario_TS221006JobPickViaStation(sTSName);
                            break;
                        case 18:
                            sTSName = "TS221007JobParkViaStation";
                            TestScenario_TS221007JobParkViaStation(sTSName);
                            break;
                        case 19:
                            sTSName = "TS221008JobBattViaStation";
                            TestScenario_TS221008JobBattViaStation(sTSName);
                            break;
                        case 20:
                            sTSName = "TS221009JobWaitViaStation";
                            TestScenario_TS221009JobWaitViaStation(sTSName);
                            break;
                        case 21:
                            sTSName = "TS221010JobDropViaStation";
                            TestScenario_TS221010JobDropViaStation(sTSName);
                            break;
                        case 22:
                            sTSName = "TS300005TransOrderPick";
                            TS300005TransOrderPick(sTSName);
                            break;
                        case 23:
                            sTSName = "TS300006TransOrderDrop";
                            TS300006TransOrderDrop(sTSName);
                            break;
                        case 24:
                            sTSName = "TS300008TransOrderMove";
                            TS300008TransOrderMove(sTSName);
                            break;
                        case 25:
                            sTSName = "TS300003TransOrderWait";
                            TS300003TransOrderWait(sTSName);
                            break;
                        case 26:
                            sTSName = "TS300072-1-TransOrderExcp1";
                            TestScenario_TS300072_1_TransOrderExcp1(sTSName);
                            break;
                        case 27:
                            sTSName = "TS300043TransOrderMode";
                            TestScenario_TS300043TransOrderMode(sTSName);
                            break;
                        case 28:
                            sTSName = "TS300071TransOrderState";
                            TestScenario_TS300071TransOrderState(sTSName);
                            break;
                            //case 29:
                            //    sTSName = "TS300011TransOrderEdit";
                            //    TestScenario_TS300011TransOrderEdit(sTSName);
                            //    break;
                        case 29:
                            sTSName = "TS300034TransOrderFlush";
                            TestScenario_TS300034TransOrderFlush(sTSName);
                            break;
                        case 30:
                            //sTestData = false;
                            sTSName = "TS300016TransOrderCancel";
                            TestScenario_TS300016TransOrderCancel(sTSName);
                            break;
                        case 31:
                            //sTestData = false;
                            sTSName = "TS300023TransOrderSuspend";
                            TestScenario_TS300023TransOrderSuspend(sTSName);
                            break;
                        case 32:
                            //sTestData = false;
                            sTSName = "TS300027TransOrderRelease";
                            TestScenario_TS300027TransOrderRelease(sTSName);
                            break;
                        case 33:
                            //sTestData = false;
                            ///sSourceID = "0070-01-01-01-01";
                            //sTransType = "PICK";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS300031TransOrderFinish";
                            TestScenario_TS300031TransOrderFinish(sTSName);
                            break;
                        case 34:
                            //sTestData = true;
                            //sSourceID = "0030-01-01";
                            //sSource2ID = "0070-01-01-01-01";
                            //sSource3ID = "0450-02-01";
                            //sDestinationID = "0420-01-01";
                            //sDestination2ID = "0360-01-01";
                            //sDestination3ID = "0360-01-02";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;
                            //mTestAgv2 = m_Agv3;
                            //mTestAgv3 = m_Agv7;
                            sTSName = "TS300025TransOrderSuspendAll";
                            TestScenario_TS300025TransOrderSuspendAll(sTSName);
                            break;
                        case 35:
                            //sTestData = false;
                            sTSName = "TS300029TransOrderReleaseAll";
                            TestScenario_TS300029TransOrderReleaseAll(sTSName);
                            break;
                        case 36:
                            //sTestData = false;
                            //sSourceID = "0030-01-01";
                            //sSource2ID = "0070-01-01-01-01";
                            //sSource3ID = "0070-01-01-02-01";
                            //sDestinationID = "0450-02-05";
                            //sDestination2ID = "0360-01-01";
                            //sDestination3ID = "0360-01-03";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS300026TransOrderSuspendAllPending";
                            TestScenario_TS300026TransOrderSuspendAllPending(sTSName);
                            break;
                        case 37: // pre-condition should run after TS300026TransOrderSuspendAllPending
                            //sTestData = false;
                            //mTestAgv = m_Agv11;
                            sTSName = "TS300030TransOrderReleaseAllPending";
                            TestScenario_TS300030TransOrderReleaseAllPending(sTSName);
                            break;
                        case 38:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //sStationID = "X0500_046";
                            //sTransType = "PICK";
                            //                 csc error string sTrans2Type = "DROP";
                            //mTestAgv = m_Agv11;
                            //mTestAgv2 = m_Agv3;
                            sTSName = "TS350057MutexAuto";
                            TestScenario_TS350057MutexAuto(sTSName);
                            break;
                        case 39:
                            //sTestData = false;
                            //sSourceID = "0040-01-01";
                            //sTransType = "PICK";
                            //mTestAgv = m_Agv11;
                            //sStationID = "X0040_013";
                            sTSName = "TS300016TransportSourceVia";
                            TestScenario_TS300016TransportSourceVia(sTSName);
                            break;
                        case 40:
                            //sTestData = false;
                            //sSourceID = "0040-01-01";
                            //sDestinationID = "0040-01-05";
                            //sTransType = "PICK";
                            // csc error               sTrans2Type = "DROP";
                            //mTestAgv = m_Agv11;
                            //sStationID = "X0040_013";
                            sTSName = "TS300017TransportDestinationVia";
                            TestScenario_TS300017TransportDestinationVia(sTSName);
                            break;
                        case 41:
                            //sTestData = false;
                            //mTestAgv = m_Agv3;
                            sTSName = "TS241099WeekPlanBatteryCharge";
                            TestScenario_TS241099WeekPlanBatteryCharge(sTSName);
                            break;
                        case 42:
                            //sTestData = false;
                            //mTestAgv = m_Agv3;
                            sTSName = "TS241047WeekPlanBatteryChargeDisable";
                            TestScenario_TS241047WeekPlanBatteryChargeDisable(sTSName);
                            break;
                        case 43:
                            //sTestData = false;
                            //mTestAgv = m_Agv3;
                            sTSName = "TS241048WeekPlanBatteryChargeDisableAll";
                            TestScenario_TS241048WeekPlanBatteryChargeDisableAll(sTSName);
                            break;
                        case 44:
                            //sTestData = false;
                            //mTestAgv = m_Agv3;
                            sTSName = "TS241049WeekPlanBatteryChargeDelete";
                            TestScenario_TS241049WeekPlanBatteryChargeDelete(sTSName);
                            break;
                        case 45:
                            //sTestData = false;
                            //mTestAgv = m_Agv3;
                            sTSName = "TS242099WeekPlanCalibration";
                            TestScenario_TS242099WeekPlanCalibration(sTSName);
                            break;
                        case 46:
                            //sTestData = false;
                            //mTestAgv = m_Agv3;
                            sTSName = "TS242047WeekPlanCalibrationDisable";
                            TestScenario_TS242047WeekPlanCalibrationDisable(sTSName);
                            break;
                        case 47:
                            //sTestData = false;
                            //mTestAgv = m_Agv3;
                            sTSName = "TS242048WeekPlanCalibrationDisableAll";
                            TestScenario_TS242048WeekPlanCalibrationDisableAll(sTSName);
                            break;
                        case 48:
                            //sTestData = false;
                            //mTestAgv = m_Agv3;
                            sTSName = "TS242049WeekPlanCalibrationDelete";
                            TestScenario_TS242049WeekPlanCalibrationDelete(sTSName);
                            break;
                        case 49:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS200083AgvStop";
                            TestScenario_TS200083AgvStop(sTSName);
                            break;
                        case 50:
                            //sTestData = false;
                            //sLocationID = "PARK_BAT";
                            //sJobType = "BATT";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS200071AgvState";
                            TestScenario_TS200071AgvState(sTSName);
                            break;
                        case 51:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv3;
                            //mTestAgvParkID = AGV3_PARKID;
                            sTSName = "TS200037AgvRetire";
                            TestScenario_TS200037AgvRetire(sTSName);
                            break;
                        case 52:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv3;
                            //mTestAgvParkID = AGV3_PARKID;
                            sTSName = "TS200053AgvDeploy";
                            TestScenario_TS200053AgvDeploy(sTSName);
                            break;
                        case 53:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv3;
                            //mTestAgvParkID = AGV3_PARKID;
                            sTSName = "TS200023AgvSuspend";
                            TestScenario_TS200023AgvSuspend(sTSName);
                            break;
                        case 54:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv3;
                            //mTestAgvParkID = AGV3_PARKID;
                            sTSName = "TS200027AgvRelease";
                            TestScenario_TS200027AgvRelease(sTSName);
                            break;
                        case 55:
                            //sTestData = false;
                            //sLocationID = AGV3_PARKID; ;
                            //sJobType = "PARK";
                            //mTestAgv = m_Agv11;
                            //Agv test2Agv = m_Agv3;
                            //mTestAgvParkID = AGV3_PARKID;
                            sTSName = "TS200012AgvModeRemoved";
                            TestScenario_TS200012AgvModeRemoved(sTSName);
                            break;
                        case 56:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = null;
                            //mTestAgvParkID = AGV3_PARKID;
                            sTSName = "TS200014AgvModeRemovedAll";
                            TestScenario_TS200014AgvModeRemovedAll(sTSName);
                            break;
                        case 57:
                            //sTestData = false;
                            //sJobType = "BATT";
                            //mTestAgv = m_Agv3;
                            sTSName = "TS200047AgvModeDisable";
                            TestScenario_TS200047AgvModeDisable(sTSName);
                            break;
                        case 58:
                            //sTestData = false;
                            //sJobType = "BATT";
                            //mTestAgv = m_Agv3;
                            sTSName = "TS200048AgvModeDisableAll";
                            TestScenario_TS200048AgvModeDisableAll(sTSName);
                            break;
                        case 59:
                            //sTestData = false;
                            //sJobType = "PICK";
                            //sSourceID = "0070-01-01-01-01";
                            //mTestAgv = m_Agv3;
                            sTSName = "TS200058AgvModeSemiAutomatic";
                            TestScenario_TS200058AgvModeSemiAutomatic(sTSName);
                            break;
                        case 60:
                            //sTestData = false;
                            //sJobType = "PICK";
                            //sSourceID = "0070-01-01-01-01";
                            //mTestAgv = m_Agv3;
                            sTSName = "TS200061AgvModeSemiAutomaticAll";
                            TestScenario_TS200061AgvModeSemiAutomaticAll(sTSName);
                            break;
                        case 61:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sTransType = "PICK";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS309306TransPickDeactiveRestart";
                            TestScenario_TS309306TransPickDeactiveRestart(sTSName);
                            break;
                            //case 61:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //sTransType = "DROP";
                            //mTestAgv = m_Agv11;
                            //  sTSName = "TS309307TransDropDeactiveRestart";
                            //  TestScenario_TS309306TransDropDeactiveRestart(sTSName);
                            //  break;
                        case 62:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sTransType = "PICK";
                            //mTestAgv = m_Agv3;
                            //mTestAgvParkID = AGV3_PARKID;
                            sTSName = "TS420047LocationDisable";
                            TestScenario_TS420047LocationDisable(sTSName);
                            break;
                        case 63:
                            //sTestData = false;
                            //sLocationID = "PARK_BAT";
                            //sJobType = "BATT";
                            //mTestAgv = m_Agv3;
                            sTSName = "TS420045LocationManual";
                            TestScenario_TS420045LocationManual(sTSName);
                            break;
                        case 64:
                            //sTestData = false;
                            //sLocationID = AGV_BATTID;
                            //sJobType = "BATT";
                            //sStationID = "X0060_093";
                            //mTestAgv = m_Agv3;
                            //mTestAgvParkID = AGV3_PARKID;
                            sTSName = "TS460047StationDisable";
                            TestScenario_TS460047StationDisable(sTSName);
                            break;
                        case 65:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //string sLoadID = "LoadA";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS400034LoadFlushAndDiscard";
                            TestScenario_TS400034LoadFlushAndDiscard(sTSName);
                            break;
                        case 66:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0070-01-01-03-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS400076LoadDiscard";
                            TestScenario_TS400076LoadDiscard(sTSName);
                            break;
                        case 67:
                            //sTestData = false;
                            //sSourceID = "0040-01-05";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS305513TransOrderDelay";
                            TestScenario_TS305513TransOrderDelay(sTSName);
                            break;
                        case 68:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS305514TransOrderDivert";
                            TestScenario_TS305514TransOrderDivert(sTSName);
                            break;
                        case 69:
                            //sTestData = false;
                            //sSourceID       = "0060-03-01";
                            //sSource2ID      = "0070-01-01-01-01";
                            //sSource3ID      = "0040-01-01";
                            //sDestinationID  = "0360-01-01";
                            //sTransType      = "MOVE";
                            //mTestAgv        = m_Agv11;
                            sTSName = "TS304455LocationClosestHighest";
                            TestScenario_TS304455LocationClosestHighest(sTSName);
                            break;
                        case 70:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sGroupID = "0070-01-01";
                            //sSource2ID = "0030-01-01";
                            //sGroup2ID = "AREA30";
                            //sTransType = "PICK";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS304457GroupHighestPriority";
                            TestScenario_TS304457GroupHighestPriority(sTSName);
                            break;
                        case 71:
                            //sTestData = false;
                            //sSourceID = "0060-03-01";
                            //sSource2ID = "0070-01-01-01-01";
                            //sSource3ID = "0040-01-01";
                            //sDestinationID = "0360-01-01";
                            //mTestAgv = m_Agv11;
                            //sTransType = "MOVE";
                            sTSName = "TS304456LoadClosestHighest";
                            TestScenario_TS304456LoadClosestHighest(sTSName);
                            break;
                        case 72:
                            //sTestData = false;
                            //sSourceID = "0240-01-01";
                            //sDestinationID = "0360-01-01";
                            //sSource2ID = "0060-03-01";
                            //sDestination2ID = "0360-01-01";
                            //sSource3ID = "0060-13-01";
                            //sDestination3ID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;  // PARK_40         to sourceID = 106 eqv cost
                            //mTestAgv2 = m_Agv3;  // PARK_060_04     to sourceID = 205 eqv cost
                            //mTestAgv3 = m_Agv7;  // PARK_060_LIFT   to sourceID = 144 eqv cost
                            sTSName = "TS305501OrderAssignmentClosest";
                            TestScenario_TS305501OrderAssignmentClosest(sTSName);
                            break;
                        case 73:
                            //sTestData = false;
                            //sSourceID = "0060-03-01";
                            //sDestinationID = "0360-01-01";
                            //sSource2ID = "0070-01-01-01-01";
                            //sDestination2ID = "0360-01-01";
                            //sSource3ID = "0040-01-01";
                            //sDestination3ID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS305502TransOrderClosest";
                            TestScenario_TS305502TransOrderClosest(sTSName);
                            break;
                        case 74:
                            //sTestData = false;
                            //sSourceID = "0240-01-01";
                            //sDestinationID = "0360-01-01";
                            //sSource2ID = "0060-03-01";
                            //sDestination2ID = "0360-01-01";
                            //sSource3ID = "0060-13-01";
                            //sDestination3ID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;  // PARK_40         to sourceID = 106 eqv cost
                            //mTestAgv2 = m_Agv3;  // PARK_060_04     to sourceID = 205 eqv cost
                            //mTestAgv3 = m_Agv7;  // PARK_060_LIFT   to sourceID = 144 eqv cost
                            sTSName = "TS305503OrderAssignmentClosestHighest";
                            TestScenario_TS305503OrderAssignmentClosestHighest(sTSName);
                            break;
                        case 75:
                            //sTestData = false;
                            //sSourceID = "0060-03-01";
                            //sDestinationID = "0360-01-01";
                            //sSource2ID = "0070-01-01-01-01";
                            //sDestination2ID = "0360-01-01";
                            //sSource3ID = "0040-01-01";
                            //sDestination3ID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS305504TransOrderClosestHighest";
                            TestScenario_TS305504TransOrderClosestHighest(sTSName);
                            break;
                        case 76:
                            //sTestData = false;
                            //sSourceID = "0060-03-01";
                            //sDestinationID = "0360-01-01";
                            //sSource2ID = "0070-01-01-01-01";
                            //sDestination2ID = "0360-01-01";
                            //sSource3ID = "0040-01-01";
                            //sDestination3ID = "0360-01-01";
                            //sTransType = "MOVE";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS305505TransOrderOldest";
                            TestScenario_TS305505TransOrderOldest(sTSName);
                            break;
                        case 77: // deadlock via		
                            //sTestData = false;
                            //sSourceID = "0030-03-01";
                            //sDestinationID = "0360-01-01";
                            //sSource2ID = "0030-01-01";
                            //sStationID = "X0500_018";
                            //sTransType = "PICK";
                            //mTestAgv = m_Agv11;
                            //mTestAgv2 = m_Agv3;
                            sTSName = "TS305507SchedulesDeadlockRulesVia";
                            TestScenario_TS305507SchedulesDeadlockRulesVia(sTSName);
                            break;
                        case 78: // battery low set to true		
                            //sTestData = false;
                            //sSourceID = "0030-03-01";
                            //sDestinationID = "0360-01-01";
                            //sSource2ID = "0030-01-01";
                            //sStationID = "X0500_018";
                            //sTransType = "PICK";
                            sTSName = "TS305508ScheduleBattRulesQueueSimLow";
                            TestScenario_TS305508ScheduleBattRulesQueueSimLow(sTSName);
                            break;
                        case 79:
                            //sTestData = false;
                            //sSourceID = "0070-01-01-01-01";
                            //sDestinationID = "0360-01-01";
                            //sTransType = "MOVE";
                            //sStationID = "X059";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS330056RoutingDynamic";
                            TestScenario_TS330056RoutingDynamic(sTSName);
                            break;
                            //case 82:	// double play	
                            //sTestData = false;
                            //sSourceID = "0240-01-01";
                            //sDestinationID = "0070-02-01-01-01";
                            //sSource2ID = "0070-02-01-03-01";
                            //string groupAID = "AREA1";
                            //      string groupBID = "0070-02-01";
                            //sTransType = "MOVE";
                            //       sTrans2Type = "PICK";
                            //mTestAgv = m_Agv11;
                            //mTestAgv2 = m_Agv3;
                            //mTestAgv3 = m_Agv7;
                            //     sTSName = "TS300063TransOrderDoublePlay";
                            //   TestScenario_TS300063TransOrderDoublePlay(sTSName);
                            //   break;
                            //case 80:	// double play	Transport released
                            //sTestData = false;
                            //sSourceID = "0240-01-01";
                            //sSource2ID = "0070-02-01-02-01";
                            //sDestinationID = "0070-02-01-03-01";
                            //sGroupID = "AREA1";
                            //groupBID = "0070-02-01";
                            //sTransType = "MOVE";
                            // csc error              sTrans2Type = "PICK";
                            //mTestAgv = m_Agv11;
                            //mTestAgv2 = m_Agv3;
                            //mTestAgv3 = m_Agv7;
                            //sTSName = "TS300064DoublePlayTransReleased";
                            //TestScenario_TS300064DoublePlayTransReleased(sTSName);
                            //break;
                        case 80:
                            sTSName = "TS300080TransPickFromGroup";
                            TestScenario_TS300080TransPickFromGroup(sTSName);
                            break;
                        case 81:
                            sTSName = "TS300081TransDropToGroup";
                            TestScenario_TS300081TransDropToGroup(sTSName);
                            break;
                        case 82:
                            sTSName = "TS200080AgvModeSemiToAuto";
                            TestScenario_TS200080AgvModeSemiToAuto(sTSName);
                            break;
                            //case 87:
                            //  sTSName = "TS830001DBSQLSERVERStopStart";
                            //  TestScenario_TS830001DBSQLSERVERStopStart(sTSName);
                            //  break;
                            /*                    case 59:	// test Agv IsAvailable with WP active		
							string args = "PARK_BAT";
							mTestAgv = m_Agv11;
							sTSName = "TS241001WeekPlanBattChargeActiveCheck";
							TestScenario_TS24XX01WeekPlanActiveCheck(sTSName, mTestAgv, "WPBatt", ref mTestStatus, args);
							break;
						case 60:	// test Agv IsAvailable with WP active		
							args = "CX01";
							mTestAgv = m_Agv11;
							sTSName = "TS241002WeekPlanCalibrationActiveCheck";
							TestScenario_TS24XX01WeekPlanActiveCheck(sTSName, mTestAgv, "WPCalib", ref mTestStatus, args);
							break;                
						case 68:
							sSourceID = "PARK_BAT";
							sTransType = "BATT";
							mTestAgv = m_Agv11;  // PARK_40         
							mTestAgv2 = m_Agv3;  // PARK_060_04     
							mTestAgv3 = m_Agv7;  // PARK_060_LIFT   
							sTSName = "TS305506ScheduleBattRulesQueueWP";
							TestScenario_TS305506ScheduleBattRulesQueueWP(sTSName, mTestAgv, mTestAgv2, mTestAgv3);
							break;
						case 73:
							sTSName = "TS242099WeekPlanCalibrationMultipleTrigger";
							TestScenario_TS242099WeekPlanCalibrationMultipleTrigger(sTSName, sTestAgvs);
							break;
						*/
                        case 199:
                            sSourceID = "0070-01-01-01-01";
                            sDestinationID = "0360-01-01";
                            sTransType = "MOVE";
                            sStationID = "X0500_050";
                            //mTestAgv = m_Agv11;
                            sTSName = "TS830002DBBufferingNotEmptyAtStartup";
                            TestScenario_TS830002DBBufferingNotEmptyAtStartup(sTSName, mTestAgv, sTransType, sSourceID,
                                                                              sDestinationID);
                            break;
                        case -1:
                            sRunStatus = "StopRunning";
                            m_Project.Facilities["Tests"].Parameters["RunStatus"].ValueAsString = sRunStatus;

                            if (testinfo.demo_9.ToLower().StartsWith("true"))
                                mLogger.LogMessageToFile(xlsBody[0, 0] + ";" + xlsBody[0, 1] + ";" + xlsBody[0, 2] + ";" +
                                                         xlsBody[0, 3] + ";");
                            else
                            {
                                for (int i = 0; i <= 81; i++)
                                {
                                    mLogger.LogMessageToFile("TestCase --> " + i);
                                    mLogger.LogMessageToFile(xlsBody[i, 0] + ";" + xlsBody[i, 1] + ";" + xlsBody[i, 2] +
                                                             ";" + xlsBody[i, 3] + ";");
                                }
                            }


                            mLogger.LogMessageToFile("------    Test Run Stoped ----");
                            sCounter = -99;
                            break;
                        case -99:
                            break;
                        default:
                            sRunStatus = "StopRunning";
                            m_Project.Facilities["Tests"].Parameters["RunStatus"].ValueAsString = sRunStatus;
                            mLogger.LogMessageToFile("------    Test case undefined: " + sCounter);
                            sCounter = -1;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
            }
        }


        /// <summary>
        /// OnDeactivate
        /// </summary>
        protected override void OnDeactivate()
        {
            try
            {
                m_Project = null;
                xApp = null;
                //m_Application = null;
                base.OnDeactivate();
            }
            catch
            {
            }
        }

        #region  //------ Test Scenarios ------------------------------------------

        //------ Test Scenarios ------------------------------------------
        //------ Test Scenarios ------------------------------------------
        //------ Test Scenarios ------------------------------------------

        // TS220001JobParkSemiAutomatic 
        private void TS220001JobParkSemiAutomatic(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  JobType:" + "PARK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    m_Project.Agvs.Automatic();
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1) Set test Agv mode semi-automatic  
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].SemiAutomatic();
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                // (2) Create Park Job at sLocationID
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    Job job = TestUtility.CreateTestJob("JobA" + sTSName, "PARK", sLocationID, sTSName, sProjectID);
                    TestUtility.RemoteLogMessage(" job carrier id : " + job.CarrierID, sTestMonitorUsed, m_Project);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA" + sTSName, "PARK", sLocationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                // (3) Wait until JobA Finished
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (4) Check test Agv lsid
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mMsg = sTestAgvs[0].ID + " at " + sTestAgvs[0].CurrentLSID;
                    if (sTestAgvs[0].CurrentLSID.ToString().Equals(sLocationID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    if (testinfo.demo_9.ToLower().StartsWith("true"))
                        EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData);
                    else
                        EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TS220002JobBattSemiAutomatic 
        private void TS220002JobBattSemiAutomatic(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    TestUtility.RemoteLogTestRunStartup(sTSName + " testagvid: " + sTestAgvs[0].ID, sTestMonitorUsed,
                                                        m_Project);

                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  JobType:" + "BATT";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    m_Project.Agvs.Automatic();
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1) Set test Agv mode semi-automatic  
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].SemiAutomatic();
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                // (2) Create Batt Job at sLocationID
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    Job job = TestUtility.CreateTestJob("JobA" + sTSName, "BATT", sLocationID, sTSName, sProjectID);
                    TestUtility.RemoteLogMessage(" job carrier id : " + job.CarrierID, sTestMonitorUsed, m_Project);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA" + sTSName, "BATT", sLocationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                // (3) Wait until JobA Finished
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (4) Check test Agv lsid
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mMsg = sTestAgvs[0].ID + " at " + sTestAgvs[0].CurrentLSID;
                    if (sTestAgvs[0].CurrentLSID.ToString().ToUpper().Equals(sLocationID.ToUpper()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    if (testinfo.demo_9.ToLower().StartsWith("true"))
                        EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData);
                    else
                        EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TS220003JobWaitSemiAutomatic
        private void TS220003JobWaitSemiAutomatic(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  JobType:" + "WAIT";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    m_Project.Agvs.Automatic();
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1) Set test Agv mode semi-automatic  
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].SemiAutomatic();
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                // (2) Create Wait Job at sLocationID
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    Job job = TestUtility.CreateTestJob("JobA" + sTSName, "WAIT", sLocationID, sTSName, sProjectID);

                    TestUtility.RemoteLogMessage(" job carrier id : " + job.CarrierID, sTestMonitorUsed, m_Project);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA" + sTSName, "WAIT", sLocationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                // (3) Wait until JobA Finished
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (4) Check test Agv lsid
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mMsg = sTestAgvs[0].ID + " at " + sTestAgvs[0].CurrentLSID;
                    if (sTestAgvs[0].CurrentLSID.ToString().ToUpper().Equals(sLocationID.ToUpper()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    if (testinfo.demo_9.ToLower().StartsWith("true"))
                        EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData);
                    else
                        EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TS220005JobPickSemiAutomatic
        private void TS220005JobPickSemiAutomatic(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  JobType:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mLogger.LogMessageToFile(sTextTestData);
                    m_Project.Agvs.Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1) Set test Agv mode semi-automatic  
                // (2) Check Agv11 loaded FALSE  
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 8000) // wait 8 sec
                    {
                        sTestResult = TestConstants.TEST_UNDEFINED;
                        mMsg = sTestAgvs[0].ID + " is loaded " + sTestAgvs[0].Loaded;
                        sTestAgvs[0].SemiAutomatic();
                        if (sTestAgvs[0].Loaded)
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_STARTING;
                        }
                    }
                }
                // (2) Create PCIK Job at sSourceIDby test AGV
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    Job job = TestUtility.CreateTestJob("JobA" + sTSName, "PICK", sSourceID, sTSName, sProjectID);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA" + sTSName, "PICK", sSourceID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                // (4) Wait until JobA Finished
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //(5) Check test Agv loaded TRUE 
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mMsg = sTestAgvs[0].ID + " is loaded " + sTestAgvs[0].Loaded;
                    if (sTestAgvs[0].Loaded)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                //(6) Check test Agv LSID 
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mMsg = sTestAgvs[0].ID + " at " + sTestAgvs[0].CurrentLSID;
                    if (sTestAgvs[0].CurrentLSID.ToString().ToUpper().Equals(sSourceID.ToUpper()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                //(7) Create JobB Pick at sSource2ID 
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    Job job = TestUtility.CreateTestJob("JobB" + sTSName, "PICK", sSource2ID, sTSName, sProjectID);
                    m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobB" + sTSName, "PICK", sSource2ID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS4;
                }
                //(7) wait 10sec
                //(8) Check JobB state PENDING 
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        mMsg = m_Job2.ID + " state " + m_Job2.State.ToString();
                        if (m_Job2.State == Job.STATE.PENDING)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mMsg = sTestAgvs[0].ID + " at " + sTestAgvs[0].CurrentLSID;
                    if (sTestAgvs[0].CurrentLSID.ToString().ToUpper().Equals(sSourceID.ToUpper()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    if (testinfo.demo_9.ToLower().StartsWith("true"))
                        EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData);
                    else
                        EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TS220006JobDropSemiAutomatic
        private void TS220006JobDropSemiAutomatic(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    m_Project.Agvs.Automatic();
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  JobType:" + "DROP";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1) Set test Agv mode semi-automatic  
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].SemiAutomatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (2) Create JobA Pick at sSourceID  
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // wait 2 sec
                    {
                        Job job = TestUtility.CreateTestJob("JobA" + sTSName, "PICK", sSourceID, sTSName, sProjectID);
                        m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                        mLogger.LogCreatedJob("JobA" + sTSName, "PICK", sSourceID, sTestAgvs[0].ID.ToString());
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                // (3) Wait until JobA Finished  
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (4) Create JobB Drop at sDestinationID  
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    string dropJobID = "JobB" + sTSName;
                    Job job = TestUtility.CreateTestJob(dropJobID, "DROP", sDestinationID, sTSName, sProjectID);
                    m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob(dropJobID, "DROP", sDestinationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS1;
                }
                // (5) Wait until JobB Finished  
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job2, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (6) Check test Agv loaded FALSE  
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mMsg = sTestAgvs[0].ID + " loaded " + sTestAgvs[0].Loaded;
                    if (sTestAgvs[0].Loaded)
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                }
                // (7) Check test Agv LSID
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mMsg = sTestAgvs[0].ID + " at " + sTestAgvs[0].CurrentLSID;
                    if (sTestAgvs[0].CurrentLSID.ToString().ToUpper().Equals(sDestinationID.ToUpper()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (8) Create JobC Drop at sDestination2ID  
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    string dropJob2ID = "JobC" + sTSName;
                    Job job = TestUtility.CreateTestJob(dropJob2ID, "DROP", sDestination2ID, sTSName, sProjectID);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob(dropJob2ID, "DROP", sDestinationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS5;
                }
                // (9) Wait 10 sec
                // (10) Check JobC state PENDING
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        mMsg = m_Job.ID + " state " + m_Job.State.ToString();
                        if (m_Job.State == Job.STATE.PENDING)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS6;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (11) Check test Agv LSID at sDestinationID
                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    mMsg = sTestAgvs[0].ID + " at " + sTestAgvs[0].CurrentLSID;
                    if (sTestAgvs[0].CurrentLSID.ToString().ToUpper().Equals(sDestinationID.ToUpper()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS220034JobFlushing
        private void TestScenario_TS220034JobFlushing(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sLocation2ID = sTestInputParams["sLocation2ID"].ToString();
                    sLocation3ID = sTestInputParams["sLocation3ID"].ToString();

                    sTextTestData = sTextTestData + "  JobA PICK:" + sSourceID;
                    sTextTestData = sTextTestData + "  JobB WAIT:" + sLocationID;
                    sTextTestData = sTextTestData + "  JobC BATT:" + sLocation2ID;
                    sTextTestData = sTextTestData + "  JobD PARK:" + sLocation3ID;
                    sTextTestData = sTextTestData + "  JobE PICK:" + sSourceID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    m_Project.Agvs.SemiAutomatic();

                    mLogger.LogMessageToFile(sTextTestData);
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1) Set test Agv mode semi-automatic  
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].SemiAutomatic();
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (2) Create JobA Pick at sSourceID
                // (3) Create JobB Wait at sLocationID
                // (4) Create JobC Batt at sLocation2ID
                // (5) Create JobD Park at sLocation3ID
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName + " job 1",
                                                        sProjectID); // create multiple jobs
                    m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    job = TestUtility.CreateTestJob("JobB-" + sTSName, "WAIT", sLocationID, sTSName + " job 2",
                                                    sProjectID);
                    m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    job = TestUtility.CreateTestJob("JobC-" + sTSName, "BATT", sLocation2ID, sTSName + " job 3",
                                                    sProjectID);
                    m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    job = TestUtility.CreateTestJob("JobD-" + sTSName, "PARK", sLocation3ID, sTSName + " job 4",
                                                    sProjectID);
                    m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mRunStatus = TestConstants.CONTINUE;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                // (6) Wait until All Jobs Finished
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilAgvAllJobsFinished(ref mRunStatus, sTestAgvs[0], sTestStartTime, 4*sWaitTime,
                                                            ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        sTestStartTime = DateTime.Now;
                    }
                }
                // (7) Create JobE Pick at sSourceID
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 second
                    {
                        Job job = TestUtility.CreateTestJob("JobE-" + sTSName, "PICK", sSourceID, sTSName + " job 5",
                                                            sProjectID);
                        m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                        mTestStatus = TestConstants.TEST_RUNS1;
                        sTestStartTime = DateTime.Now;
                    }
                }
                // (8) Flushing All Jobs
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 second
                    {
                        sTestAgvs[0].Jobs.FlushAll();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                // (9) Check test Agv only has JobE
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 second
                    {
                        if (sTestAgvs[0].Jobs.GetArray().Length == 1) // check agv still has 1 jobs?
                        {
                            var job = (Job) sTestAgvs[0].Jobs.GetArray()[0];
                            mMsg = "ONE job in the job list it is " + job.ID + " with state:" + job.State.ToString();
                            if (job.ID.ToString().ToUpper().StartsWith("JOBE"))
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                mLogger.LogPassLine(mMsg);
                                TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mLogger.LogFailLine(mMsg);
                                TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            }
                        }
                        else if (sTestAgvs[0].Jobs.GetArray().Length == 0) // agv no jobs?
                        {
                            mMsg = "All jobs flushed includepending job ";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            mMsg = "not all FINISHED jobs flushed, there are " + sTestAgvs[0].Jobs.GetArray().Length +
                                   " in jobs list ";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
            }
        }

        // TestScenario_TS220016JobCanceling
        private void TestScenario_TS220016JobCanceling(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.Automatic();
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].SemiAutomatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (2).	Create JobA Pick at sSourceID by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName, sProjectID);
                        // create pick job
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                // (3).	Create JobB Pick at sSource2ID by AGV11 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // wait 2 second
                    {
                        Job job = TestUtility.CreateTestJob("JobB-" + sTSName, "PICK", sSource2ID,
                                                            sTSName + " pending job", sProjectID);
                            // create a pending job
                        m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                // (4).	During execution of JobA, Cancel JobB
                if (mTestStatus == TestConstants.TEST_RUNS) // cancel pending job
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 second
                    {
                        m_Job2.Cancel();
                        mLogger.LogMessageToFile("Cancel job:" + m_Job2.ID);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                // (5).	Wait until JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, 2, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (6).	Check JobB state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mMsg = "Canceled job state is:" + m_Job2.State.ToString();
                    if (m_Job2.State == Job.STATE.FINISHED)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (7).	Check JobB outcome CANCELLED
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mMsg = "Canceled job outcome is:" + m_Job2.Outcome.ToString();
                    if (m_Job2.Outcome == Job.OUTCOME.CANCELLED)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (8).	Check JobB Cancelled True
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mMsg = "Canceled job cancelled is:" + m_Job2.Cancel();
                    if (m_Job2.Cancel())
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS220017JobCancelCurrent
        private void TestScenario_TS220017JobCancelCurrent(string sTSName)
        {
            try
            {
                // clearn up
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.Automatic();
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  Source ID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1).	Set Agv11 mode semi-automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].SemiAutomatic();
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (2).	Create JobA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName + " current job",
                                                        sProjectID); // create batt job
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                // (3).	Create JobB Pick at 0040-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    Job job2 = TestUtility.CreateTestJob("JobB-" + sTSName, "DROP", sDestinationID,
                                                         sTSName + " current job", sProjectID); // create batt job
                    m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job2);
                    mRunStatus = TestConstants.CONTINUE;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                // (4).	During execution of JobA, Cancel JobA
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.BUSY, sTestStartTime, sWaitTime,
                                                  ref mMsg);
                    if (mMsg.StartsWith("OK"))
                    {
                        if (mRunStatus == TestConstants.CHECK_END)
                        {
                            if (mMsg.StartsWith("OK"))
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                mLogger.LogPassLine(mMsg);
                                TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                                m_Job.Cancel();
                                sTestStartTime = DateTime.Now;
                                mTestStatus = TestConstants.TEST_RUNS1;
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mLogger.LogFailLine(mMsg);
                                TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                                mTestStatus = TestConstants.TEST_FINISHED;
                            }
                        }
                    }
                }
                // (5).	Check JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mMsg = "Job " + m_Job.ID + " state is " + m_Job.State.ToString();
                    if (m_Job.State == Job.STATE.FINISHED)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                    else
                    {
                        mTime = DateTime.Now - sTestStartTime;
                        if (mTime.TotalMilliseconds >= sWaitTime*60*1000) // wait swaittime min
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("after waitTime: " + mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (6).	Check JobA outcome CANCELLED
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mMsg = "Canceled job outcome is:" + m_Job.Outcome.ToString();
                    if (m_Job.Outcome == Job.OUTCOME.CANCELLED)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (7).	Check JobA Cancelled True
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mMsg = "Canceled job cancelled is:" + m_Job.Cancel();
                    if (m_Job.Cancel())
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS220019JobExhausted(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED) // clearn up
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.SemiAutomatic();

                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  Source ID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        // (2) Create TransportA
                        //Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, sDestinationID, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        mMsg = string.Empty;
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    // (3) Wait until TransportA state RETRIEVED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2,
                                                        Transport.STATE.RETRIEVED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        var job = (Job) sTestAgvs[0].Jobs.GetArray()[1];
                        job.Cancel(); // cancel job
                        mLogger.LogMessageToFile("cancel job:" + job.ID);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        var job = (Job) sTestAgvs[0].Jobs.GetArray()[1];
                        mMsg = "After 10 sec, " + job.ID + " has status " + job.State.ToString();
                        if (job.State == Job.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    // (5) Wait until TransportA state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS220020JobAborting
        private void TestScenario_TS220020JobAborting(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED) // clearn up
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  Source ID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1).	Set Agv11 mode semi-automatic
                // (2).	Create JobA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].SemiAutomatic();
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName, sProjectID);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA-" + sTSName, "PICK", sSourceID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (3).	Wait until JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, 2, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_STARTING;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (4).	Create JobB Drop at 0360-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    Job job = TestUtility.CreateTestJob("JobB-" + sTSName, "DROP", sDestinationID, sTSName, sProjectID);
                    m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobB-" + sTSName, "DROP", sDestinationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                // (5).	Abort JobB immediately
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        m_Job2.Abort(); // abort job
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                // (6).	Check JobB state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = m_Job2.ID + " state is: " + m_Job2.State.ToString();
                        if (m_Job2.State == Job.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //get agv current lsid
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // Wait 10 sec.
                    {
                        sAgvCurrentLSID = sTestAgvs[0].CurrentLSID.ToString();
                        mMsg = sTestAgvs[0].ID + " current lsid id:: " + sAgvCurrentLSID;
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                }
                // (7).	Check Agv11 is not moving
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    TestUtility.CheckAgvNotMoving(ref mRunStatus, sTestAgvs[0], sAgvCurrentLSID,
                                                  sTestStartTime, 10 /*sec*/, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (8).	Check Agv11 LSID not at sDestinationID
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mMsg = sTestAgvs[0].ID + "is located at " + sTestAgvs[0].CurrentLSID;
                    if (!sTestAgvs[0].CurrentLSID.ToString().Equals(sDestinationID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS5;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (9).	Create JobC Drop at 0360-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    Job job = TestUtility.CreateTestJob("JobC-" + sTSName, "DROP", sDestinationID, sTSName, sProjectID);
                    m_Job3 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobC-" + sTSName, "DROP", sDestinationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS6;
                }
                // (10).Wait until 0360-01-01 is locked by Agv11 --- this is test load will not be dropped at destination
                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    TestUtility.WaitUntilAgvLockLSID(ref mRunStatus, sTestAgvs[0], sDestinationID + ".DROP",
                                                     sTestStartTime, 2, ref mMsg);
                    TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS7;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (11)	Abort JobC
                if (mTestStatus == TestConstants.TEST_RUNS7)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        m_Job3.Abort(); // abort job
                        mLogger.LogMessageToFile("Abort Jib: " + m_Job3.ID);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS8;
                    }
                }
                // (12)	Wait until Agv11 at 0360-01-01
                if (mTestStatus == TestConstants.TEST_RUNS8)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime,
                                                     sDestinationID, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS9;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (13)	Check JobC state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS9)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = m_Job3.ID + " state is : " + m_Job3.State.ToString();
                        if (m_Job3.State == Job.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_STOP;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (14)	Check Agv11 loaded TRUE
                if (mTestStatus == TestConstants.TEST_STOP)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = sTestAgvs[0].ID + " loaded is : " + sTestAgvs[0].Loaded;
                        if (sTestAgvs[0].Loaded)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_AFTER_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                // (15)	Check JobC outcome ABORTED
                if (mTestStatus == TestConstants.TEST_AFTER_RUNS)
                {
                    mMsg = "Canceled job outcome is:" + m_Job3.Outcome.ToString();
                    if (m_Job3.Outcome == Job.OUTCOME.ABORTED)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_AFTER;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (16)	Check JobC IsAborted True
                if (mTestStatus == TestConstants.TEST_AFTER)
                {
                    mMsg = "Canceled job IsAborted is:" + m_Job3.IsAborted();
                    if (m_Job3.IsAborted())
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS220023JobSuspending
        private void TestScenario_TS220023JobSuspending(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED) // clearn up
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.Automatic();
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  Source ID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                // (2)	Create JobA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].SemiAutomatic();
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName, sProjectID);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA-" + sTSName, "PICK", sSourceID, sTestAgvs[0].ID.ToString());
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (3)Create JobB Drop at 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    Job job = TestUtility.CreateTestJob("JobB-" + sTSName, "DROP", sDestinationID, sTSName, sProjectID);
                    m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobB-" + sTSName, "DROP", sDestinationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                // (4)	During JobA, Suspend JobB
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        m_Job2.Suspend();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                // (5) Wait until JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (6)	Check JobB state PENDING
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // Wait 5 sec.
                    {
                        mMsg = m_Job2.ID + " state is : " + m_Job2.State.ToString();
                        mMsg = mMsg + Environment.NewLine + m_Job.ID + " state is : " + m_Job.State.ToString();

                        if (m_Job2.State == Job.STATE.PENDING)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (7)	Check Agv11 LSID
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = sTestAgvs[0].ID + " currentLSID is : " + sTestAgvs[0].CurrentLSID;
                        if (!sTestAgvs[0].CurrentLSID.ToString().Equals(sDestinationID))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (8)	Check JobB suspend TRUE
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = m_Job2.ID + " suspend is : " + m_Job2.Suspended;
                        if (m_Job2.Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS7;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //TODO suspended job
                //9.	Create Suspended JobC DROP at 0360-01-01 by AGV11
                /*if (mTestStatus == TestConstants.TEST_RUNS4)
				{
					Job job = TestUtility.CreateTestJob("JobC-" + sTSName, "DROP", sDestinationID, sTSName, sProjectID);
					job.Suspend();
					m_Job3 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
					mLogger.LogCreatedJob("JobC-" + sTSName, "DROP", sDestinationID, sTestAgvs[0].ID.ToString());
					sTestStartTime = DateTime.Now;
					mTestStatus = TestConstants.TEST_RUNS5;
				}
				//10.	Check JobC Suspended True
				if (mTestStatus == TestConstants.TEST_RUNS5)
				{
					mTime = DateTime.Now - sTestStartTime;
					if (mTime.TotalMilliseconds >= 2000)				// Wait 2 sec.
					{
						mMsg = m_Job3.ID.ToString() + " suspend is : " + m_Job2.Suspended;
						if (m_Job3.Suspended)
						{
							sTestResult = TestConstants.TEST_PASS;
							mLogger.LogPassLine(mMsg);
							TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
							mTestStatus = TestConstants.TEST_RUNS6;
						}
						else
						{
							sTestResult = TestConstants.TEST_FAIL;
							mLogger.LogFailLine(mMsg);
							TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
							mTestStatus = TestConstants.TEST_FINISHED;
						}
						
					}
				}
				//11.	Check Agv11 is not moving
				//get agv current lsid
				if (mTestStatus == TestConstants.TEST_RUNS6)
				{
					mTime = DateTime.Now - sTestStartTime;
					if (mTime.TotalMilliseconds >= 10000)					// Wait 10 sec.
					{
						sAgvCurrentLSID = sTestAgvs[0].CurrentLSID.ToString();
						mMsg = sTestAgvs[0].ID.ToString() + " current lsid id:: " + sAgvCurrentLSID;
						sTestResult = TestConstants.TEST_PASS;
						mLogger.LogPassLine(mMsg);
						TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
						mRunStatus = TestConstants.CONTINUE;
						sTestStartTime = DateTime.Now;
						mTestStatus = TestConstants.TEST_RUNS7;
					}
				}
				 */
                //12.	Release JobB
                if (mTestStatus == TestConstants.TEST_RUNS7)
                {
                    //TestUtility.CheckAgvNotMoving(ref mRunStatus, sTestAgvs[0], sAgvCurrentLSID,
                    //    sTestStartTime, 10, ref mMsg);  // 10 sec
                    //if (mRunStatus == TestConstants.CHECK_END)
                    //{
                    //if (mMsg.StartsWith("OK"))
                    //{
                    m_Job2.Release();
                    //sTestResult = TestConstants.TEST_PASS;
                    //mLogger.LogPassLine(mMsg);
                    mLogger.LogPassLine("Release m_JobB: " + m_Job2.ID);
                    TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS8;
                    //}
                    //else
                    //{
                    //    sTestResult = TestConstants.TEST_FAIL;
                    //    mLogger.LogFailLine(mMsg);
                    //    TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    //    mTestStatus = TestConstants.TEST_FINISHED;
                    //}
                    //}
                }

                //13.	Wait until JobB state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS8)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job2, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }


                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS220027JobReleasing
        private void TestScenario_TS220027JobReleasing(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED) // clearn up
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.Automatic();
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  Source ID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                // (2)	Create JobA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].SemiAutomatic();
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName, sProjectID);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA-" + sTSName, "PICK", sSourceID, sTestAgvs[0].ID.ToString());
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (3)	Create JobB Drop at 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    Job job = TestUtility.CreateTestJob("JobB-" + sTSName, "DROP", sDestinationID, sTSName, sProjectID);
                    m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobB-" + sTSName, "DROP", sDestinationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                // (4)	During JobA, Suspend JobB
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        m_Job2.Suspend();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                // (5) Check JobB state PENDING
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // Wait 5 sec.
                    {
                        mMsg = m_Job2.ID + " state is : " + m_Job2.State.ToString();
                        mMsg = mMsg + Environment.NewLine + m_Job.ID + " state is : " + m_Job.State.ToString();

                        if (m_Job2.State == Job.STATE.PENDING)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (6) Wait until JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (7)	Check Agv11 LSID not at Destination
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = sTestAgvs[0].ID + " currentLSID is : " + sTestAgvs[0].CurrentLSID;
                        if (!sTestAgvs[0].CurrentLSID.ToString().Equals(sDestinationID))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (8)	Check JobB suspend TRUE
                // (9)	Release JobB
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = m_Job2.ID + " suspend is : " + m_Job2.Suspended;
                        if (m_Job2.Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            m_Job2.Release();
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (10)	Wait until JobB state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job2, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (11)	Check JobB suspend FALSE
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = m_Job2.ID + " suspend is : " + m_Job2.Suspended;
                        if (!m_Job2.Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS220024JobSuspendCurrent
        private void TestScenario_TS220024JobSuspendCurrent(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED) // clearn up
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  Source ID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                // (2)	Create JobA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].SemiAutomatic();
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName, sProjectID);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA-" + sTSName, "PICK", sSourceID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (3)	Wait until JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_STARTING;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (4)	Create JobB Drop at 0360-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    Job job = TestUtility.CreateTestJob("JobB-" + sTSName, "DROP", sDestinationID, sTSName, sProjectID);
                    m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobB-" + sTSName, "DROP", sDestinationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                // (5)	Suspend Current jobs for Agv11
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        m_Job2.Suspend(); // suspend job
                        mLogger.LogMessageToFile(m_Job2.ID + " is suspended");
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //get agv current lsid
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // Wait 10 sec.
                    {
                        sAgvCurrentLSID = sTestAgvs[0].CurrentLSID.ToString();
                        mMsg = sTestAgvs[0].ID + " current lsid id:: " + sAgvCurrentLSID;
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                // (6)	Check Agv11 is not moving
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.CheckAgvNotMoving(ref mRunStatus, sTestAgvs[0], sAgvCurrentLSID,
                                                  sTestStartTime, 10 /*sec*/, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (7)	Check JobB Suspended True
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = m_Job2.ID + " suspend is : " + m_Job2.Suspended;
                        if (m_Job2.Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (8)	Release Current jobs for Agv11
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // Wait 3 sec.
                    {
                        mLogger.LogPassLine("job released");
                        TestUtility.RemoteLogPassLine("Releasing " + m_Job2.ID, sTestMonitorUsed, m_Project);
                        m_Job2.Release();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS5;
                    }
                }
                // (9)	Wait until 0360-01-01 is locked by Agv11 and Agv11 at somepoint like X077
                // (10)	Suspend JobB
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    TestUtility.WaitUntilAgvLockLSID(ref mRunStatus, sTestAgvs[0], sDestinationID + ".DROP",
                                                     sTestStartTime, sWaitTime, ref mMsg);
                    TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            m_Job2.Suspend();
                            mLogger.LogPassLine(m_Job2.ID + "  suspended again, but destination locked by this Agv " +
                                                "   and agv will continue going to destination ");
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS6;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (11)	Wait until Agv11 at 0360-01-01
                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[0], sTestStartTime, 2, sDestinationID,
                                                     ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS7;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (12) Check Agv11 loaded is TRUE
                if (mTestStatus == TestConstants.TEST_RUNS7)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = sTestAgvs[0].ID + " loaded is : " + sTestAgvs[0].Loaded;
                        if (sTestAgvs[0].Loaded)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS8;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (7)	Check JobB Suspended True
                if (mTestStatus == TestConstants.TEST_RUNS8)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = m_Job2.ID + " suspend is : " + m_Job2.Suspended;
                        if (m_Job2.Suspended)
                        {
                            m_Job2.Release();
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 20000) // Wait 20 sec.
                    {
                        EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    }
                }
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TTS220028JobReleaseCurrent
        private void TestScenario_TTS220028JobReleaseCurrent(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED) // clearn up
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  Source ID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                // (2)	Create JobA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].SemiAutomatic();
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName, sProjectID);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA-" + sTSName, "PICK", sSourceID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // wait until job status BUSY 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.BUSY,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_STARTING;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (3) Suspend Current Jobs for Agv11
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        m_Job.Suspend(); // suspend job
                        mMsg = m_Job.ID + " is suspended and agv at "
                               + sTestAgvs[0].CurrentLSID;
                        mLogger.LogMessageToFile(mMsg);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                // (4)	Check JobA state 
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 12000) // Wait 12 sec.
                    {
                        mMsg = m_Job.ID + " state " + m_Job.State.ToString();
                        if (m_Job.State == Job.STATE.BUSY)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            //TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (5)	Check Agv11 LSID
                // (6)	Release Current Jobs for Agv11
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 1000) // Wait 1 sec.
                    {
                        string lsid = sTestAgvs[0].CurrentLSID.ToString();
                        mMsg = sTestAgvs[0].ID + " current lsid: " + lsid;
                        if (!lsid.Equals(sSourceID))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            m_Job.Release();
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (7)	Wait until JobA state FINISHED 
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (8)	Check Agv11 LSID
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        string lsid = sTestAgvs[0].CurrentLSID.ToString();
                        mMsg = sTestAgvs[0].ID + " current lsid is : " + lsid;
                        if (lsid.Equals(sSourceID))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (9)	Check JobA Suspended False
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = m_Job.ID + " suspend is : " + m_Job.Suspended;
                        if (!m_Job.Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS220025JobSuspendAll
        private void TestScenario_TS220025JobSuspendAll(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED) // clearn up
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.Automatic();
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sLocation2ID = sTestInputParams["sLocation2ID"].ToString();
                    sLocation3ID = sTestInputParams["sLocation3ID"].ToString();

                    sTextTestData = sTextTestData + "  Source ID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Location2ID:" + sLocation2ID;
                    sTextTestData = sTextTestData + "  Location3ID:" + sLocation3ID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                // (2)	Create JobA Pick at 0070-01-01-01-01 by AGV11
                // (3)	Create JobB Wait at W0070-01-01 by AGV11
                // (4)	Create JobC Batt at PARK_BAT by AGV11 
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].SemiAutomatic();
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName, sProjectID);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA-" + sTSName, "PICK", sSourceID, sTestAgvs[0].ID.ToString());

                    Job job2 = TestUtility.CreateTestJob("JobB-" + sTSName, "WAIT", sLocation2ID, sTSName, sProjectID);
                    m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job2);
                    mLogger.LogCreatedJob("JobB-" + sTSName, "WAIT", sLocation2ID, sTestAgvs[0].ID.ToString());

                    Job job3 = TestUtility.CreateTestJob("JobC-" + sTSName, "BATT", sLocation3ID, sTSName, sProjectID);
                    m_Job3 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job3);
                    mLogger.LogCreatedJob("JobC-" + sTSName, "BATT", sLocation3ID, sTestAgvs[0].ID.ToString());

                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (5)	Suspend All Jobs
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // Wait 5 sec.
                    {
                        sTestAgvs[0].Jobs.Suspend(); // suspend All jobs
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //get agv current lsid
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // Wait 10 sec.
                    {
                        sAgvCurrentLSID = sTestAgvs[0].CurrentLSID.ToString();
                        mMsg = sTestAgvs[0].ID + " current lsid id:: " + sAgvCurrentLSID;
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                // (6)	Check Agv11 is not moving
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.CheckAgvNotMoving(ref mRunStatus, sTestAgvs[0], sAgvCurrentLSID,
                                                  sTestStartTime, 10 /*sec*/, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (7)	Check JobA state BUSY
                // (8)	JobB and JobC state PENDING
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // Wait 5 sec.
                    {
                        mMsg = m_Job.ID + " state is : " + m_Job.State.ToString() + Environment.NewLine
                               + m_Job2.ID + " state is : " + m_Job2.State.ToString() + Environment.NewLine
                               + m_Job3.ID + " state is : " + m_Job3.State.ToString();
                        if (m_Job.State == Job.STATE.BUSY &&
                            m_Job2.State == Job.STATE.PENDING &&
                            m_Job3.State == Job.STATE.PENDING)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            mTime = DateTime.Now - sTestStartTime;
                            if (mTime.TotalMilliseconds >= 2*60000) //2 min
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mLogger.LogFailLine("after 2 min : " + mMsg);
                                TestUtility.RemoteLogFailLine("after 2 min: " + mMsg, sTestMonitorUsed, m_Project);
                                mTestStatus = TestConstants.TEST_FINISHED;
                            }
                        }
                    }
                }
                // (9)	Check JobA, JobB and JobC Suspended True
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // Wait 2 sec.
                    {
                        mMsg = m_Job.ID + " suspend is : " + m_Job.Suspended;
                        if (m_Job.Suspended &&
                            m_Job2.Suspended &&
                            m_Job3.Suspended)
                        {
                            sTestAgvs[0].Jobs.Release();
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mRunStatus = TestConstants.CONTINUE;
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // wait until all state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    TestUtility.WaitUntilAgvAllJobsState(ref mRunStatus, sTestAgvs[0], Job.STATE.FINISHED,
                                                         sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS220029JobReleaseAll
        private void TestScenario_TS220029JobReleaseAll(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED) // clearn up
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.Automatic();
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sLocation2ID = sTestInputParams["sLocation2ID"].ToString();
                    sLocation3ID = sTestInputParams["sLocation3ID"].ToString();

                    sTextTestData = sTextTestData + "  Source ID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Location2ID:" + sLocation2ID;
                    sTextTestData = sTextTestData + "  Location3ID:" + sLocation3ID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //  (1)	Set Agv11 mode semi-automatic
                //  (2)	Create JobA Pick at 0070-01-01-01-01 by AGV11
                //  (3)	Create JobB Wait at W0070-01-01 by AGV11
                //  (4)	Create JobC Batt at PARK_BAT by AGV11 
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].SemiAutomatic();
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName, sProjectID);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA-" + sTSName, "PICK", sSourceID, sTestAgvs[0].ID.ToString());

                    Job job2 = TestUtility.CreateTestJob("JobB-" + sTSName, "WAIT", sLocation2ID, sTSName, sProjectID);
                    m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job2);
                    mLogger.LogCreatedJob("JobB-" + sTSName, "WAIT", sLocation2ID, sTestAgvs[0].ID.ToString());

                    Job job3 = TestUtility.CreateTestJob("JobC-" + sTSName, "BATT", sLocation3ID, sTSName, sProjectID);
                    m_Job3 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job3);
                    mLogger.LogCreatedJob("JobC-" + sTSName, "BATT", sLocation3ID, sTestAgvs[0].ID.ToString());

                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (5)	Suspend All Jobs
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // Wait 5 sec.
                    {
                        sTestAgvs[0].Jobs.Suspend(); // suspend All jobs
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //get agv current lsid
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // Wait 10 sec.
                    {
                        sAgvCurrentLSID = sTestAgvs[0].CurrentLSID.ToString();
                        mMsg = sTestAgvs[0].ID + " current lsid id:: " + sAgvCurrentLSID;
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                // (6)	Check Agv11 is not moving
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.CheckAgvNotMoving(ref mRunStatus, sTestAgvs[0], sAgvCurrentLSID,
                                                  sTestStartTime, 10 /*sec*/, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (7)	Check JobA, state BUSY
                // (8)	Check JobB and JobC state PENDING
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // Wait 5 sec.
                    {
                        mMsg = m_Job.ID + " state is : " + m_Job.State.ToString() + Environment.NewLine
                               + m_Job2.ID + " state is : " + m_Job2.State.ToString() + Environment.NewLine
                               + m_Job3.ID + " state is : " + m_Job3.State.ToString();
                        if (m_Job.State == Job.STATE.BUSY &&
                            m_Job2.State == Job.STATE.PENDING &&
                            m_Job3.State == Job.STATE.PENDING)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            mTime = DateTime.Now - sTestStartTime;
                            if (mTime.TotalMilliseconds >= 2*60000) //2 min
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mLogger.LogFailLine("after 2 min : " + mMsg);
                                TestUtility.RemoteLogFailLine("after 2 min: " + mMsg, sTestMonitorUsed, m_Project);
                                mTestStatus = TestConstants.TEST_FINISHED;
                            }
                        }
                    }
                }
                // (9)	Check All jobs Suspended True
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // Wait 5 sec.
                    {
                        mMsg = m_Job.ID + " suspended is : " + m_Job.Suspended + Environment.NewLine
                               + m_Job2.ID + " suspended is : " + m_Job2.Suspended + Environment.NewLine
                               + m_Job3.ID + " suspended is : " + m_Job3.Suspended;
                        if (m_Job.Suspended &&
                            m_Job2.Suspended &&
                            m_Job3.Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (10)	Release All jobs
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // Wait 3 sec.
                    {
                        sTestAgvs[0].Jobs.Release();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                }
                // (11)	Wait until All jobs state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.WaitUntilAgvAllJobsState(ref mRunStatus, sTestAgvs[0], Job.STATE.FINISHED,
                                                         sTestStartTime, 5, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // 12.	Check All jobs Suspended False
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // Wait 5 sec.
                    {
                        mMsg = m_Job.ID + " suspended is : " + m_Job.Suspended + Environment.NewLine
                               + m_Job2.ID + " suspended is : " + m_Job2.Suspended + Environment.NewLine
                               + m_Job3.ID + " suspended is : " + m_Job3.Suspended;
                        if (!m_Job.Suspended &&
                            !m_Job2.Suspended &&
                            !m_Job3.Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS221006JobPickViaStation
        private void TestScenario_TS221006JobPickViaStation(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  StationID:" + sStationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                // (2)	Set Agv11 speed low (500)
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    sTestAgvs[0].SemiAutomatic();
                    sTestAgvs[0].SimSpeed = 500;
                    TestUtility.RemoteLogTestRunStartup(sTSName, sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (3)	Create JobA Pick at 0040-01-01 by AGV11 via X0040_013
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 8000) // 8 seconde
                    {
                        Job job = TestUtility.CreateTestJob("JobA", "PICK", sSourceID, sStationID, sTSName);
                        m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                        sTestStartTime = DateTime.Now;
                        mLogger.LogCreatedJob(m_Job.ID.ToString(), "PICK", sSourceID, sTestAgvs[0].ID.ToString());
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                // (4)	Wait until Agv11 via station X0040_013
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    string id = sTestAgvs[0].CurrentLSID.ToString();
                    TestUtility.RemoteLogMessage(id, sTestMonitorUsed, m_Project);
                    if (sTestAgvs[0].CurrentLSID.ToString().Equals(sStationID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        sTestAgvs[0].SimSpeed = 10000;
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                    else if (mTime.TotalMilliseconds >= 3*sWaitTime*60*1000) //  3 * waitTime sec
                    {
                        mMsg = "after 300 second,  agv current lsid is" + id + " and via station id is:" + sStationID;
                        sTestResult = TestConstants.TEST_FAIL;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (5)	Check JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED, sTestStartTime, sWaitTime,
                                                  ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[0].SimSpeed = 10000;
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                sTestAgvs[0].SimSpeed = 10000;
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TestScenario_TS221007JobParkViaStation
        private void TestScenario_TS221007JobParkViaStation(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;
                    sTextTestData = sTextTestData + "  StationID:" + sStationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                // (2)	Set Agv11 speed low (500)
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    sTestAgvs[0].SemiAutomatic();
                    sTestAgvs[0].SimSpeed = 500;
                    TestUtility.RemoteLogTestRunStartup(sTSName, sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (3)	Create JobA Park at PARK_250_1 by AGV11 via XSR026_1
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 8000) // 8 seconde
                    {
                        Job job = TestUtility.CreateTestJob("JobA", "PARK", sLocationID, sStationID, sTSName);
                        m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                        sTestStartTime = DateTime.Now;
                        mLogger.LogCreatedJob(m_Job.ID.ToString(), "PARK", sLocationID, sTestAgvs[0].ID.ToString());
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                // (4)	Wait until Agv11 via station XSR026_1
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    string id = sTestAgvs[0].CurrentLSID.ToString();
                    TestUtility.RemoteLogMessage(id, sTestMonitorUsed, m_Project);
                    if (sTestAgvs[0].CurrentLSID.ToString().Equals(sStationID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        sTestAgvs[0].SimSpeed = 10000;
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                    else if (mTime.TotalMilliseconds >= 3*sWaitTime*60*1000) //  3 * waitTime sec
                    {
                        mMsg = "after 300 second,  agv current lsid is" + id + " and via station id is:" + sStationID;
                        sTestResult = TestConstants.TEST_FAIL;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (5)	Check JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED, sTestStartTime, sWaitTime,
                                                  ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[0].SimSpeed = 10000;
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                sTestAgvs[0].SimSpeed = 10000;
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TestScenario_TS221008JobBattViaStation
        private void TestScenario_TS221008JobBattViaStation(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;
                    sTextTestData = sTextTestData + "  StationID:" + sStationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                // (2)	Set Agv11 speed low (500)
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    TestUtility.RemoteLogTestRunStartup(sTSName, sTestMonitorUsed, m_Project);
                    sTestAgvs[0].SimSpeed = 500;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (3)	Create JobA Batt at PARK_BAT by AGV11 via X0040_006
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 8000) // 8 seconde
                    {
                        Job job = TestUtility.CreateTestJob("JobA", "BATT", sLocationID, sStationID, sTSName);
                        m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                        sTestStartTime = DateTime.Now;
                        mLogger.LogCreatedJob(m_Job.ID.ToString(), "BATT", sLocationID, sTestAgvs[0].ID.ToString());
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                // (4)	Wait until Agv11 via station X0040_006
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    string id = sTestAgvs[0].CurrentLSID.ToString();
                    //TestUtility.RemoteLogMessage(id, sTestMonitorUsed, m_Project);
                    if (id.Equals(sStationID))
                    {
                        sTestAgvs[0].SimSpeed = 10000;
                        sTestResult = TestConstants.TEST_PASS;
                        mMsg = "after " + mTime.TotalSeconds + " seconds,  agv current lsid is" + id +
                               " and via station id is:" + sStationID;
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                    else if (mTime.TotalMilliseconds >= 3*sWaitTime*60*1000) //  3 * waitTime sec
                    {
                        mMsg = "after 300 second,  agv current lsid is" + id + " and via station id is:" + sStationID;
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestResult = TestConstants.TEST_FAIL;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (5)	Check JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED, sTestStartTime, sWaitTime,
                                                  ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[0].SimSpeed = 10000;
                    TestUtility.RemoteLogPassLine("Agv speed reset", sTestMonitorUsed, m_Project);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                sTestAgvs[0].SimSpeed = 10000;
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TestScenario_TS221009JobWaitViaStation
        private void TestScenario_TS221009JobWaitViaStation(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;
                    sTextTestData = sTextTestData + "  StationID:" + sStationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                // (2)	Set Agv11 speed low (500)
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    TestUtility.RemoteLogTestRunStartup(sTSName, sTestMonitorUsed, m_Project);
                    sTestAgvs[0].SimSpeed = 500;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (3)	Create JobA Wait at W0070-01-01 by AGV11 via X0040_006
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 8000) // 8 seconde
                    {
                        Job job = TestUtility.CreateTestJob("JobA", "WAIT", sLocationID, sStationID, sTSName);
                        m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                        sTestStartTime = DateTime.Now;
                        mLogger.LogCreatedJob(m_Job.ID.ToString(), "WAIT", sLocationID, sTestAgvs[0].ID.ToString());
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                // (4)	Wait until Agv11 via station X0040_006
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    string id = sTestAgvs[0].CurrentLSID.ToString();
                    mLogger.LogMessageToFile(sTestAgvs[0].ID + " has LSID:" + id);
                    //TestUtility.RemoteLogMessage(id, sTestMonitorUsed, m_Project);
                    if (id.Equals(sStationID))
                    {
                        sTestAgvs[0].SimSpeed = 10000;
                        sTestResult = TestConstants.TEST_PASS;
                        mMsg = "after " + mTime.TotalSeconds + " seconds,  agv current lsid is" + id +
                               " and via station id is:" + sStationID;
                        mLogger.LogMessageToFile(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                    else if (mTime.TotalMilliseconds >= 10*sWaitTime*60*1000) //  3 * waitTime sec
                    {
                        mMsg = "after 20 min,  agv current lsid is" + id + " and via station id is:" + sStationID;
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mLogger.LogMessageToFile(mMsg);
                        sTestResult = TestConstants.TEST_FAIL;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (5)	Check JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED, sTestStartTime, sWaitTime,
                                                  ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[0].SimSpeed = 10000;
                    TestUtility.RemoteLogPassLine("Agv speed reset", sTestMonitorUsed, m_Project);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                sTestAgvs[0].SimSpeed = 10000;
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TestScenario_TS221010JobDropViaStation
        private void TestScenario_TS221010JobDropViaStation(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;

                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  StationID:" + sStationID;

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile(sTextTestData);

                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                // (2)	Create JobA Pick at 0040-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        sTestAgvs[0].SemiAutomatic();
                        // first do a pick job 
                        Job job = TestUtility.CreateTestJob("JobA", "PICK", sSourceID, sTSName, sProjectID);
                        m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;

                        mLogger.LogCreatedJob(m_Job.ID.ToString(), "PICK", sSourceID, sTestAgvs[0].ID.ToString());
                        mTestStatus = TestConstants.TEST_INITED;
                    }
                }

                // (3)	Wait until JobA state FINISHED
                // (4)	Set Agv11 speed low (500)
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED, sTestStartTime, sWaitTime,
                                                  ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            // decrease agv speed
                            sTestAgvs[0].SimSpeed = 500;
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_STARTING;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (5)	Create JobB Drop at 0040-01-05 by AGV11 via X0040_013
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // 2 seconde
                    {
                        Job job = TestUtility.CreateTestJob("JobB", "DROP", sDestinationID, sStationID, sTSName);
                            // create a drop job
                        m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mLogger.LogCreatedJob(m_Job2.ID.ToString(), "DROP", sDestinationID, sTestAgvs[0].ID.ToString());
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                // (6)	Wait until Agv11 via station X0040_013
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime, sStationID,
                                                     ref mMsg);
                    if (mMsg.StartsWith("OK"))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        // increase agv speed
                        sTestAgvs[0].SimSpeed = 10000;
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                    else
                    {
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestResult = TestConstants.TEST_FAIL;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (7)	Wait until JobB state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job2, Job.STATE.FINISHED, sTestStartTime, sWaitTime,
                                                  ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[0].SimSpeed = 10000;
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                sTestAgvs[0].SimSpeed = 10000;
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TS300005TransOrderPick
        private void TS300005TransOrderPick(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    mLogger.LogMessageToFile(sTextTestData);

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (2)	Create TransportA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        // (2) Create TransportA
                        //PICK at 0060-03-01-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, null, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "pick", sSourceID, null,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                // (3)	Check TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (4)	Check Agv11 current LSID
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mMsg = sTestAgvs[0].ID + " current LSID is:" + sTestAgvs[0].CurrentLSID;
                    if (sTestAgvs[0].CurrentLSID.ToString().Equals(sSourceID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (5)	Check Agv11 Loaded
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mMsg = sTestAgvs[0].ID + " is loaded:" + sTestAgvs[0].Loaded;
                    if (sTestAgvs[0].Loaded)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (6)	Check TransportA outcome SUCCESS
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mMsg = m_Transport.ID + " has outcome:" + m_Transport.Outcome.ToString();
                    if (m_Transport.Outcome == Transport.OUTCOME.SUCCESS)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TS300006TransOrderDrop
        private void TS300006TransOrderDrop(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (2)	Create TransportA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        // (2) Create TransportA
                        //PICK at 0060-03-01-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, null, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "pick", sSourceID, null,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                // (3)	Wait until TransportA state FINISHED.
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (4)	Create TransportB Drop at 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.DROP, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, null, sDestinationID, 5, false);
                        transport.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "DROP", sDestinationID, null,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                // (5)	Wait until TransportB state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (6)	Check Agv11 loaded FALSE
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mMsg = sTestAgvs[0].ID + " is loaded:" + sTestAgvs[0].Loaded;
                    if (!sTestAgvs[0].Loaded)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                // (7)	Check TransportB outcome SUCCESS
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mMsg = m_Transport.ID + " has outcome:" + m_Transport.Outcome.ToString();
                    if (m_Transport.Outcome == Transport.OUTCOME.SUCCESS)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TS300008TransOrderMove
        private void TS300008TransOrderMove(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (2)	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, sDestinationID, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                // (3)	Wait until TransportA FINISHED.
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    // (3) Wait until TransportA state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2*sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (4)	Check TransportA Outcome SUCCESS
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mMsg = m_Transport.ID + " has outcome:" + m_Transport.Outcome.ToString();
                    if (m_Transport.Outcome == Transport.OUTCOME.SUCCESS)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TS300003TransOrderWait
        private void TS300003TransOrderWait(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS2;
                }
                // (2)	Create TransportA Wait at W0420-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.WAIT, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, null, sLocationID, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);

                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "WAIT", sLocationID, null,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                // (3)	Check TransportA state VERIFIED
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.VERIFIED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.	Check Wait reason
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        mMsg = "Wait reason: " + m_Transport.WaitTransitionReason.Symbol;
                        if (
                            m_Transport.WaitTransitionReason.Symbol.StartsWith(
                                "TRANSPORT_MANAGER_TRANSPORT_LOAD_NOT_RETRIEVED"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            mMsg = "TransportA wait reason" + m_Transport.WaitTransitionReason.Arguments[0].ToUpper();
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (5)	Cancel TransportA
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        m_Transport.Cancel();
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //6.	Create TransportB Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        // (2) Create TransportA
                        //PICK at 0060-03-01-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, null, 5, false);
                        transport.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "PICK", sSourceID, null,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                }
                // (7)	Wait until TransportB state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (8)	Create TransportC Wait at W0420-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.WAIT, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, null, sLocationID, 5, false);
                        transport.ID = "TransportC" + sTSName;
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "WAIT", sLocationID, null,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS5;
                    }
                }
                // (9)	Wait until TransportC state FINISHED.
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport3, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS6;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (10)	Check Agv11 at wait location 
                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    mMsg = sTestAgvs[0].ID + " has LSID: " + sTestAgvs[0].CurrentLSID;
                    if (sTestAgvs[0].CurrentLSID.ToString().Equals(sLocationID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300072_1_TransOrderExcp1
        private void TestScenario_TS300072_1_TransOrderExcp1(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode automatic
                // (2)	Create TransportA Pick at AREA1 by AGV11 
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].Automatic();
                    m_Transport = TestUtility.CreateTestTransport("PICK", sTestAgvs[0], sSourceID, null, sTSName,
                                                                  ref m_Project, ref m_Transport);
                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "PICK", null, sSourceID,
                                                sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS;
                }

                // if no exception, this test is failed
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconds
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(" No Exception generated");
                        TestUtility.RemoteLogFailLine(" No Exception Generated", sTestMonitorUsed, m_Project);
                        EndTestCase(sTSName, sTestResult, ref sTextTestData);
                        ;
                    }
                }
            }
            catch (Exception ex)
            {
                //3.	Check TransportA not created
                // (4).	Check Exception Message 
                string msg = "TRANSPORT_MANAGER_GROUP_TYPE_NOT_ALLOWED";
                if (ex.Message.StartsWith(msg))
                {
                    sTestResult = TestConstants.TEST_PASS;
                    mLogger.LogPassLine("");
                    TestUtility.RemoteLogPassLine("", sTestMonitorUsed, m_Project);
                }
                else
                {
                    sTestResult = TestConstants.TEST_EXCEPTION;
                    mLogger.LogTestException(ex.Message, ex.StackTrace);
                    TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                }
                EndTestCase(sTSName, sTestResult, ref sTextTestData);
            }
        }

        // TestScenario_TS300043TransOrderMode
        private void TestScenario_TS300043TransOrderMode(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transType:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    //sTextTestData = sTextTestData + "  Agv ParkID:"+sAgvsInitialID[sTestAgv.ID.ToString()].ToString();

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // (1)	Set Agv11 mode semi-automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].SemiAutomatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                // (2)	Create TransportA Move from 0700-01-01-01-01 to 0360-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    var transport = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[0].ID.ToString(), null,
                                                  sSourceID, sDestinationID);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);

                    mLogger.LogCreatedTransport("TransportA" + sTSName, "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                // (3)	Check TransportA state VERIFIED
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.VERIFIED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // (4)	Set Agv11 mode Disable
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    sTestAgvs[0].Disabled();
                    mTestStatus = TestConstants.TEST_RUNS2;
                }
                //5.	Check TransportA state VERIFIED
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.VERIFIED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Set Agv11 mode Remove
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    sTestAgvs[0].Removed();
                    mTestStatus = TestConstants.TEST_RUNS4;
                }
                //7.	Check TransportA state VERIFIED 
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.VERIFIED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //8.	Set Agv11 mode automatic and Restart Agv11
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    sTestAgvs[0].Automatic();
                    sTestAgvs[0].Restart();
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS6;
                }
                // (9)	Check Transport state
                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2*sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300071TransOrderState
        public void TestScenario_TS300071TransOrderState(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    sTextTestData = sTextTestData + "  testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transType:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode semi-automatic
                //2.	Set Agv11 speed very low
                //3.	Create TransportA Move from at 0040-01-01 to 0040-01-05 by AGV11 
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].SemiAutomatic();
                    sTestAgvs[0].SimSpeed = 500;
                    sTestAgvs[0].Suspend();
                    //Move at 0040-01-01 to 0040-01-05 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, sDestinationID, 3*sWaitTime, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //4.	Wait until TransportA state VERIFIED.
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.VERIFIED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            sTestAgvs[0].Release();
                            sTestAgvs[0].Automatic();
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //5.	Wait until TransportA state ASSIGNED.
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2*sWaitTime,
                                                        Transport.STATE.ASSIGNED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Wait until TransportA state RETRIEVING.
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2*sWaitTime,
                                                        Transport.STATE.RETRIEVING, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            sTestAgvs[0].SimSpeed = 10;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //7.	Wait until TransportA state RETRIEVED
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2*sWaitTime,
                                                        Transport.STATE.RETRIEVED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            sTestAgvs[0].SimSpeed = 500;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //8.	Wait until TransportA state STORING.
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2,
                                                        Transport.STATE.STORING, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            sTestAgvs[0].SimSpeed = 10;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //9.	Wait until TransportA state STORED.
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.STORED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            sTestAgvs[0].SimSpeed = 10;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //10.	Wait until TransportA state FINISHED..
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[0].SimSpeed = 10000;
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                sTestAgvs[0].SimSpeed = 10000;
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300011TransOrderEdit
        private void TestScenario_TS300011TransOrderEdit(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transType:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode to automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec, create new transport
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[0].ID.ToString(),
                                                      "Load300011", sSourceID, sDestinationID);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport("TransportA" + sTSName, "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString());
                        TestUtility.RemoteLogMessage("orig transportdestid =" + m_Transport.DestinationID,
                                                     sTestMonitorUsed, m_Project);

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //3.	During Agv11 picking, Edit TransportA destination to 0420-01-01
                if (mTestStatus == TestConstants.TEST_RUNS) // Edit Transport Order
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        var transports
                            = (Transports) m_Transport.Parent;

                        // Check if the transport still exist in the Transports collection
                        if (!transports.Contains(m_Transport.ID))
                        {
                            //MessageBox.Show("Transport '" + m_Transport.ID.ToString() + "' does not exist anymore.", "Edit Transport");
                            throw new Exception("transport id not exist:" + m_Transport.ID);
                        }

                        // Create a new identical Transport object.
                        var editTransport
                            = new Transport();

                        editTransport.SetXml(m_Transport.GetXml());
                        // Set the properties
                        //SetProperties(editTransport);
                        editTransport.DestinationID = sDestination2ID;

                        TestUtility.RemoteLogMessage("edit transport destid =" + editTransport.DestinationID,
                                                     sTestMonitorUsed, m_Project);


                        //				// Verify the transport
                        //				if ( ! m_TransportManager.Verify( editTransport ) )
                        //				{
                        //					MessageBox.Show(
                        //						this,
                        //						"Verification of transport failed. " + System.Environment.NewLine + System.Environment.NewLine + m_TransportManager.LastMessage.GetTextFormatted(),
                        //						"Edit Transport",
                        //						MessageBoxButtons.OK, MessageBoxIcon.Stop
                        //						);
                        //					return;
                        //				}

                        // Edit the transport through the TransportManager
                        if (m_TransportManager.EditTransport(editTransport) == null)
                        {
                            MessageBox.Show(
                                "Edit of transport failed." + Environment.NewLine + Environment.NewLine +
                                m_TransportManager.LastMessage.GetTextFormatted(),
                                "New Transport"
                                );
                            throw new Exception(" Edit the transport through the TransportManager failed :" +
                                                editTransport.ID);
                        }
                        //	Transport transport = new Transport( Transport.COMMAND.MOVE, null, testAgv.ID.ToString(), null, sourceID, dest2ID);
                        //	m_Transport = m_Project.TransportManager.EditTransport(transport);
                        //	mLogger.LogMessageToFile("new destination id is :"+m_Transport.DestinationID.ToString());
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }

                //4.	Wait until TransportA finished
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //5.	Check TransportA destination
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 90000) // wait 90 sec
                    {
                        if (m_Transport.DestinationID.ToString().Equals(sDestination2ID))
                        {
                            mLogger.LogMessageToFile("------    Test " + sTSName + " :   m_Transport destID:" +
                                                     m_Transport.DestinationID);
                            if (m_Transport.State >= Transport.STATE.FINISHED)
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                mLogger.LogPassLine("");
                                TestUtility.RemoteLogPassLine("", sTestMonitorUsed, m_Project);
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mLogger.LogFailLine("");
                                TestUtility.RemoteLogFailLine("" + m_Transport.State, sTestMonitorUsed, m_Project);
                            }
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("");
                            TestUtility.RemoteLogFailLine("" + m_Transport.DestinationID, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message + Environment.NewLine + ex.StackTrace,
                                                   sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300034TransOrderFlush
        private void TestScenario_TS300034TransOrderFlush(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transType:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    // (2) Create TransportA
                    //Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport("TransportA" + sTSName, "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Flushing TransportA during picking  
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    // (3) Wait until TransportA state ASSIGND
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.ASSIGNED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            // Flush
                            m_Project.Transports.FlushAll();

                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.	Check TransportA is exist 
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 seconde
                    {
                        Object[] transportOrders = m_Project.Transports.GetArray();
                        if (transportOrders.Length > 0)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);

                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //5.	Wait until TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Check TransportA is exist
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 seconde
                    {
                        Object[] transportOrders = m_Project.Transports.GetArray();
                        if (transportOrders.Length > 0)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //7.	Flushing TransportA
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 seconde
                    {
                        m_Project.Transports.FlushAll();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                }
                //8.	Check TransportA is not exist
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 seconde
                    {
                        Object[] transportOrders = m_Project.Transports.GetArray();
                        if (transportOrders.Length == 0)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300016TransOrderCancel
        private void TestScenario_TS300016TransOrderCancel(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transType:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 automatic.
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, null, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport("TransportA" + sTSName, "PICK", sSourceID, null,
                                                sTestAgvs[0].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Cancel TransportA during picking 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    // (3) Wait until TransportA state ASSIGND
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.ASSIGNED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            // Cancel
                            m_Transport.Cancel();

                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.	Check TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 4000) // wait 4 sec
                    {
                        mMsg = m_Transport.ID + " has state " + m_Transport.State.ToString();
                        if (m_Transport.State >= Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);

                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //5.	Wait until Agv11 return to park location
                //6.	Check Agv11 state is not ready
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    // (3) Wait until Agv at parklocation state ASSIGND
                    /*TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgv, sTestStartTime, sWaitTime, 
						sAgvsInitialID[sTestAgv.ID.ToString()].ToString(), ref mMsg);
			
					if (mRunStatus == TestConstants.CHECK_END)
					{
						if (mMsg.StartsWith("OK"))
						{
							sTestResult = TestConstants.TEST_PASS;
							mLogger.LogPassLine(mMsg);
							TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
							// Cancel
							m_Transport.Cancel();

							sTestStartTime = DateTime.Now;
							mTestStatus = TestConstants.TEST_RUNS2;
						}
						else
						{
							sTestResult = TestConstants.TEST_FAIL;
							mLogger.LogFailLine(mMsg);
							TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
							mTestStatus = TestConstants.TEST_FINISHED;
						}
					}*/

                    mTestStatus = TestConstants.TEST_RUNS2;
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mMsg = sTestAgvs[0].ID + "has state :" + sTestAgvs[0].State.ToString();
                    /*if (sTestAgv.State == Agv.STATE.NOT_READY)
					{
						sTestResult = TestConstants.TEST_PASS;
						mLogger.LogPassLine(mMsg);
						TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
					   
						sTestStartTime = DateTime.Now;
						
					 }
					else
					{
						sTestResult = TestConstants.TEST_FAIL;
						mLogger.LogFailLine(mMsg);
						TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
						mTestStatus = TestConstants.TEST_FINISHED;
					}*/
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300023TransOrderSuspend
        private void TestScenario_TS300023TransOrderSuspend(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transType:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 4000) //  wait 4 sec
                    {
                        m_Transport = TestUtility.CreateTestTransport("PICK", sTestAgvs[0], sSourceID, null, sTSName,
                                                                      ref m_Project, ref m_Transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "PICK", null, sSourceID,
                                                    sTestAgvs[0].ID.ToString());
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //3.	Suspend TransportA during picking.
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 4000) // wait 4 sec
                    {
                        m_TransportState = m_Transport.State;
                        mLogger.LogPassLine("====before suspend: transport state is :" + m_TransportState.ToString());
                        m_Transport.Suspend();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //4.	Wait 60 sec
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 60000) // wait 60 sec
                    {
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                }
                //5.	Check TransportA state unchanged
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        mLogger.LogPassLine("====after suspend: transport state is :" + m_TransportState.ToString());
                        if (m_Transport.State == m_TransportState)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("state unchanged");
                            TestUtility.RemoteLogPassLine("state unchanged", sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("xxx" + m_Transport.State.ToString());
                            mLogger.LogFailLine("xxx2" + m_TransportState.ToString());
                            TestUtility.RemoteLogFailLine("xxx", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Check TransportA Suspended True
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        if (m_Transport.IsSuspended())
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("");
                            TestUtility.RemoteLogPassLine("", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("");
                            TestUtility.RemoteLogFailLine("", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //7.	Check Agv11state ¡°not ready¡±


                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    m_Transport.Cancel();
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300027TransOrderRelease
        private void TestScenario_TS300027TransOrderRelease(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                {
                    sTextTestData = string.Empty;
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transType:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 4000) //  wait 4 sec
                    {
                        m_Transport = TestUtility.CreateTestTransport("PICK", sTestAgvs[0], sSourceID, null, sTSName,
                                                                      ref m_Project, ref m_Transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "PICK", null, sSourceID,
                                                    sTestAgvs[0].ID.ToString());
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //3.	Suspend TransportA during picking.
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 4000) // wait 4 sec
                    {
                        m_TransportState = m_Transport.State;
                        m_Transport.Suspend();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //4.	Wait until Transport suspended True
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 4000) // wait 4 sec
                    {
                        mMsg = m_Transport.ID + " suspended is:" + m_Transport.Suspended;
                        if (m_Transport.Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //5.	Check TransportA state unchanged
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        if (m_Transport.State == m_TransportState)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("state unchanged");
                            TestUtility.RemoteLogPassLine("state unchanged", sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("");
                            TestUtility.RemoteLogFailLine("", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Release TrabsportA
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    m_Transport.Release();
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS4;
                }
                //7.	Wait until TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(1)" + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //8.	Check TransportA Suspended False
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 4000) // wait 4 sec
                    {
                        mMsg = m_Transport.ID + " suspended is:" + m_Transport.Suspended;
                        if (!m_Transport.Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300031TransOrderFinish
        private void TestScenario_TS300031TransOrderFinish(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transType:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Pick at 0070-01-01-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        m_Transport = TestUtility.CreateTestTransport("PICK", sTestAgvs[0], sSourceID, null, sTSName,
                                                                      ref m_Project, ref m_Transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "PICK", null, sSourceID,
                                                    sTestAgvs[0].ID.ToString());
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //3.	Finish TransportA during picking.
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 4000) // wait 4 sec
                    {
                        m_Transport.Finish();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //4.	Wait 10 sec
                //5.	Check TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        string agvLSID = sTestAgvs[0].CurrentLSID.ToString();
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("");
                            TestUtility.RemoteLogPassLine("", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("" + agvLSID);
                            TestUtility.RemoteLogFailLine("" + agvLSID, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Check TransportA outcome FINISHED-MANUALLY
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mMsg = m_Transport.ID + " has outcome: " + m_Transport.Outcome.ToString();
                    if (m_Transport.Outcome == Transport.OUTCOME.FINISHED_MANUALLY)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300025TransOrderSuspendAll
        private void TestScenario_TS300025TransOrderSuspendAll(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                {
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);
                }

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "testAgv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "testAgv3:" + sTestAgvs[2].ID;
                    sTextTestData = sTextTestData + "transType:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  testAgv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  testAgv3:" + sTestAgvs[2].ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  Destination3ID:" + sDestination3ID;
                    m_Project.Agvs.SemiAutomatic();
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11, Agv3 and Agv7 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    for (int i = 0; i < sTestAgvs.Length; i++)
                    {
                        sTestAgvs[i].Automatic();
                        sTestAgvs[i].SimSpeed = 500;
                    }

                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                ///2.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[0].ID.ToString(), null,
                                                      sSourceID, sDestinationID);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString());
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //3.	Create TransportB Move from 0070-01-02-01-01 to 0360-01-01 by AGV3 
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 1000) // wait 1 sec
                    {
                        var transport2 = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[1].ID.ToString(), null,
                                                       sSource2ID, sDestination2ID);
                        transport2.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport2);
                        mLogger.LogCreatedTransport(transport2.ID.ToString(), "MOVE", sSource2ID, sDestination2ID,
                                                    sTestAgvs[1].ID.ToString());

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //4.	Create TransportC Move from 0070-01-03-01-01 to 0360-01-01 by AGV7 
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 1000) // wait 1 sec
                    {
                        var transport3 = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[2].ID.ToString(), null,
                                                       sSource3ID, sDestination3ID);
                        transport3.ID = "TransportC" + sTSName;
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport3);
                        mLogger.LogCreatedTransport(transport3.ID.ToString(), "MOVE", sSource3ID, sDestination3ID,
                                                    sTestAgvs[2].ID.ToString());

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //5.	Wait until All transports state ASSIGNED
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;

                    bool allAssigned = true;
                    if (m_Transport.State < Transport.STATE.ASSIGNED ||
                        m_Transport2.State < Transport.STATE.ASSIGNED ||
                        m_Transport3.State < Transport.STATE.ASSIGNED)
                    {
                        allAssigned = false;
                        mLogger.LogPassLine("All transports assigned: --> false");
                    }

                    if (allAssigned)
                    {
                        TestUtility.RemoteLogMessage("All assigned done ", sTestMonitorUsed, m_Project);
                        mLogger.LogPassLine("All transports assigned done");
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        if (mTime.TotalMilliseconds >= 60000) // 60 sec
                        {
                            TestUtility.RemoteLogMessage("after 60 sec  not All assigned done ", sTestMonitorUsed,
                                                         m_Project);
                            mLogger.LogFailLine("after 60 sec  not All assigned done ");
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	During picking, Suspend All transports
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        mLogger.LogPassLine("All transports suspending...");
                        m_Project.Transports.Suspend();
                        mLogger.LogPassLine("All transports suspended...");
                        TestUtility.RemoteLogMessage("All suspended done ", sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS5;
                    }
                }
                //7.	Wait 10 sec
                //8.	Check Agv11, Agv3 and Agv7 unloaded
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 15000) // wait 15 seconde
                    {
                        if (!sTestAgvs[0].Loaded && !sTestAgvs[1].Loaded && !sTestAgvs[2].Loaded)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("All testAgvs Unloaded");
                            TestUtility.RemoteLogPassLine("All testAgvs Unloaded", sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("");
                            TestUtility.RemoteLogFailLine("NOT All testAgvs Unloaded", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                //9.	Check All Transports Suspended True
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 seconde
                    {
                        if (m_Transport.IsSuspended() && m_Transport2.IsSuspended() && m_Transport3.IsSuspended())
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("All transport state is suspended");
                            TestUtility.RemoteLogPassLine("All transport state is suspended", sTestMonitorUsed,
                                                          m_Project);
                            m_Project.Transports.Release();
                            mRunStatus = TestConstants.CONTINUE;
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS6;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("Not All transport Suspennded:" + Environment.NewLine
                                                + m_Transport.ID + " is suspended:" + m_Transport.IsSuspended() +
                                                Environment.NewLine
                                                + m_Transport2.ID + " is suspended:" + m_Transport2.IsSuspended() +
                                                Environment.NewLine
                                                + m_Transport3.ID + " is suspended:" + m_Transport3.IsSuspended() +
                                                Environment.NewLine
                                );
                            TestUtility.RemoteLogFailLine("NOT All transport state is suspended", sTestMonitorUsed,
                                                          m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        sTestAgvs[0].SimSpeed = 10000;
                        sTestAgvs[1].SimSpeed = 10000;
                        sTestAgvs[2].SimSpeed = 10000;
                    }
                }

                //10.	Wait until All Transports State FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    TestUtility.WaitUntilAllTransportFinished(ref mRunStatus, m_Project, sTestStartTime, 3*sWaitTime,
                                                              ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300029TransOrderReleaseAll
        private void TestScenario_TS300029TransOrderReleaseAll(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                {
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);
                }

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = string.Empty;
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "testAgv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "testAgv3:" + sTestAgvs[2].ID;
                    sTextTestData = sTextTestData + "transType:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  testAgv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  testAgv3:" + sTestAgvs[2].ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  Destination3ID:" + sDestination3ID;
                    m_Project.Agvs.SemiAutomatic();
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11, Agv3 and Agv7 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    for (int i = 0; i < sTestAgvs.Length; i++)
                        sTestAgvs[i].Automatic();

                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                ///2.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[0].ID.ToString(), null,
                                                      sSourceID, sDestinationID);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString());
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //3.	Create TransportB Move from 0070-01-02-01-01 to 0360-01-01 by AGV3 
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 1000) // wait 1 sec
                    {
                        var transport2 = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[1].ID.ToString(), null,
                                                       sSource2ID, sDestination2ID);
                        transport2.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport2);
                        mLogger.LogCreatedTransport(transport2.ID.ToString(), "MOVE", sSource2ID, sDestination2ID,
                                                    sTestAgvs[1].ID.ToString());

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //4.	Create TransportC Move from 0070-01-03-01-01 to 0360-01-01 by AGV7 
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 1000) // wait 1 sec
                    {
                        var transport3 = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[2].ID.ToString(), null,
                                                       sSource3ID, sDestination3ID);
                        transport3.ID = "TransportC" + sTSName;
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport3);
                        mLogger.LogCreatedTransport(transport3.ID.ToString(), "MOVE", sSource3ID, sDestination3ID,
                                                    sTestAgvs[2].ID.ToString());

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //5.	Wait until All transports state ASSIGNED
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;

                    bool allAssigned = true;
                    if (m_Transport.State < Transport.STATE.ASSIGNED ||
                        m_Transport2.State < Transport.STATE.ASSIGNED ||
                        m_Transport3.State < Transport.STATE.ASSIGNED)
                    {
                        allAssigned = false;
                    }

                    if (allAssigned)
                    {
                        TestUtility.RemoteLogMessage("All assigned done ", sTestMonitorUsed, m_Project);
                        mLogger.LogPassLine("All transports assigned done");
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        if (mTime.TotalMilliseconds >= sWaitTime*60*1000) // waitTime min
                        {
                            TestUtility.RemoteLogMessage("after 60 sec  not All assigned done ", sTestMonitorUsed,
                                                         m_Project);
                            mLogger.LogFailLine("after 60 sec  not All assigned done ");
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	During picking, Suspend All transports
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 15000) // wait 15 sec
                    {
                        TestUtility.RemoteLogMessage("All suspended done ", sTestMonitorUsed, m_Project);
                        mLogger.LogPassLine("aAll suspended done ");
                        m_Project.Transports.Suspend();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS5;
                    }
                }
                //7.	Wait 10 sec
                //8.	Check Agv11, Agv3 and Agv7 unloaded
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 seconde
                    {
                        if (!sTestAgvs[0].Loaded && !sTestAgvs[1].Loaded && !sTestAgvs[2].Loaded)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("All testAgvs Unloaded");
                            TestUtility.RemoteLogPassLine("All testAgvs Unloaded", sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("");
                            TestUtility.RemoteLogFailLine("NOT All testAgvs Unloaded", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                //9.	Check All Transports Suspended True
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 seconde
                    {
                        if (m_Transport.IsSuspended() && m_Transport2.IsSuspended() && m_Transport3.IsSuspended())
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("All transport state is suspended");
                            TestUtility.RemoteLogPassLine("All transport state is suspended", sTestMonitorUsed,
                                                          m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS6;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("NOT All transport state is suspended");
                            TestUtility.RemoteLogFailLine("NOT All transport state is suspended", sTestMonitorUsed,
                                                          m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //10.	Release all transports
                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    m_Project.Transports.Release();
                    mRunStatus = TestConstants.CONTINUE;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS7;
                }
                //11.	Wait until all transports state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS7)
                {
                    TestUtility.WaitUntilAllTransportFinished(ref mRunStatus, m_Project, sTestStartTime, 3*sWaitTime,
                                                              ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            //mMsg = "transports not finished";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
            }
        }

        // TestScenario_TS300026TransOrderSuspendAllPending
        private void TestScenario_TS300026TransOrderSuspendAllPending(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transType:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  Destination3ID:" + sDestination3ID;

                    m_Project.Agvs.SemiAutomatic();

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].Automatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[0].ID.ToString(), null,
                                                      sSourceID, sDestinationID);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString());
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //3.	Create TransportB Move from 0070-01-02-01-01 to 0360-01-01 by AGV11 
                //4.	Create TransportC Move from 0070-01-03-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport2 = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[0].ID.ToString(), null,
                                                       sSource2ID, sDestination2ID);
                        transport2.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport2);
                        mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "MOVE", sSource2ID, sDestination2ID,
                                                    sTestAgvs[0].ID.ToString());

                        var transport3 = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[0].ID.ToString(), null,
                                                       sSource3ID, sDestination3ID);
                        transport3.ID = "TransportC" + sTSName;
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport3);
                        mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "MOVE", sSource3ID, sDestination3ID,
                                                    sTestAgvs[0].ID.ToString());

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //5.	Check TransportB and TransportC state Pending
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    if (m_Transport2.State <= Transport.STATE.PENDING &&
                        m_Transport3.State <= Transport.STATE.PENDING)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine("Transport B and C state Pending");
                        TestUtility.RemoteLogPassLine("Transport B and C state Pending", sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogPassLine("Transport B or C state not Pending");
                        TestUtility.RemoteLogPassLine("Transport B or C state not Pending", sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                //6.	Suspend All pending Transport
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 6000) // wait 6 sec
                    {
                        m_Project.Transports.SuspendPending();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //7.	Wait until TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2*sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            //mMsg = "transports not finished";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                //8.	Check TransportB and TransportC Suspended True
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // wait 2 sec
                    {
                        if (m_Transport2.IsSuspended() &&
                            m_Transport3.IsSuspended())
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("");
                            TestUtility.RemoteLogPassLine("", sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            mMsg = "pending transports not suspended";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS300030TransOrderReleaseAllPending
        private void TestScenario_TS300030TransOrderReleaseAllPending(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transType:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  Destination3ID:" + sDestination3ID;

                    m_Project.Agvs.SemiAutomatic();

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].Automatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[0].ID.ToString(), null,
                                                      sSourceID, sDestinationID);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString());
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //3.	Create TransportB Move from 0070-01-02-01-01 to 0360-01-01 by AGV11 
                //4.	Create TransportC Move from 0070-01-03-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport2 = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[0].ID.ToString(), null,
                                                       sSource2ID, sDestination2ID);
                        transport2.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport2);
                        mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "MOVE", sSource2ID, sDestination2ID,
                                                    sTestAgvs[0].ID.ToString());

                        var transport3 = new Transport(Transport.COMMAND.MOVE, null, sTestAgvs[0].ID.ToString(), null,
                                                       sSource3ID, sDestination3ID);
                        transport3.ID = "TransportC" + sTSName;
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport3);
                        mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "MOVE", sSource3ID, sDestination3ID,
                                                    sTestAgvs[0].ID.ToString());

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //5.	Check TransportB and TransportC state Pending
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    if (m_Transport2.State <= Transport.STATE.PENDING &&
                        m_Transport3.State <= Transport.STATE.PENDING)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine("Transport B and C state Pending");
                        TestUtility.RemoteLogPassLine("Transport B and C state Pending", sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogPassLine("Transport B or C state not Pending");
                        TestUtility.RemoteLogPassLine("Transport B or C state not Pending", sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                //6.	Suspend All pending Transport
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 6000) // wait 6 sec
                    {
                        m_Project.Transports.SuspendPending();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //7.	Wait until TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            mMsg = "transports not finished";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //8.	Check TransportB and TransportC Suspended True 
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // wait 2 sec
                    {
                        if (m_Transport2.IsSuspended() &&
                            m_Transport3.IsSuspended())
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("Transport B and C suspended TRUE");
                            TestUtility.RemoteLogPassLine("Transport B and C suspended TRUE", sTestMonitorUsed,
                                                          m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            mMsg = "pending transports not suspended";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //9.	Release All Pending Transport
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        m_Project.Transports.ReleasePending();
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS5;
                    }
                }
                //10.	Wait until all transports finished
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    TestUtility.WaitUntilAllTransportFinished(ref mRunStatus, m_Project, sTestStartTime, 3*sWaitTime,
                                                              ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS6;
                        }
                        else
                        {
                            mMsg = "transports not finished";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //11.	Check All transports suspended False
                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (!m_Transport.IsSuspended() &&
                            !m_Transport2.IsSuspended() &&
                            !m_Transport3.IsSuspended())
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("All Transports state FINISHED");
                            TestUtility.RemoteLogPassLine("All Transports state FINISHED", sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            mMsg = "not transports not finished";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }


                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS350057MutexAuto
        private void TestScenario_TS350057MutexAuto(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  testAgv2:" + sTestAgvs[1].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Destination:" + sDestinationID;
                    sTextTestData = sTextTestData + "  testAgv2 CurrentLSIDID:" + sStationID;

                    sTestResult = TestConstants.TEST_UNDEFINED;
                    m_Project.Agvs.SemiAutomatic();

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.Set Agv11, Agv3  mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].Automatic();
                    sTestAgvs[1].Automatic();
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.Create TransportA Pick at 0070-01-01-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    var transport = new Transport(Transport.COMMAND.PICK, null, sTestAgvs[0].ID.ToString(), null,
                                                  sSourceID, null);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mRunStatus = TestConstants.CONTINUE;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.Wait until TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.Create TransportB Pick at 0070-01-01-01-01 by AGV3
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    var transport = new Transport(Transport.COMMAND.PICK, null, sTestAgvs[1].ID.ToString(), null,
                                                  sSourceID, null);
                    transport.ID = "TransportB" + sTSName;
                    m_Transport2 = m_Project.TransportManager.NewTransport(transport);

                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS1;
                }
                //5.Wait until Agv3 stop at station X0500_046
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[1], sTestStartTime, sWaitTime, sStationID,
                                                     ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(1)" + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.Create TransportC Drop at 360-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    var transport = new Transport(Transport.COMMAND.DROP, null, sTestAgvs[0].ID.ToString(), null, null,
                                                  sDestinationID);
                    transport.ID = "TransportC" + sTSName;
                    m_Transport3 = m_Project.TransportManager.NewTransport(transport);

                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS3;
                }
                //7.Wait until TransportB state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //8.Wait untilTransportC stateFINISHED
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport3, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // 9,Check Agv3 at location at 0700-01-01-01-01
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mMsg = sTestAgvs[1].ID + " at " + sTestAgvs[1].CurrentLSID;
                    ;
                    if (sTestAgvs[1].CurrentLSID.ToString().Equals(sSourceID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
            }
        }

        // TestScenario_TS300016TransportSourceVia
        public void TestScenario_TS300016TransportSourceVia(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transtype:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  viaStationID:" + sStationID;
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    m_Project.Agvs.SemiAutomatic();
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                //2.	Set testAgv SimSpeed to 500
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    sTestAgvs[0].Automatic();
                    sTestAgvs[0].SimSpeed = 500;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //3.	Create TransportA Pick at 0040-01-01 by AGV11 with priority 5.
                //4.	Select Source Via:  CX03
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2000) // 2 seconde
                    {
                        // Pick at 0040-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, null, 5, false);
                        transport.SourceViaLSID = sStationID;
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "PICK", sSourceID, null,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //5.	Wait until testAgv pass CX03
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime, sStationID,
                                                     ref mMsg);
                    if (mMsg.StartsWith("OK"))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                    else
                    {
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestResult = TestConstants.TEST_FAIL;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 1000) //  1 sec
                    {
                        sTestAgvs[0].SimSpeed = 10000;
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                }
                //6.	Wait until testAgv stop at pick location  0040-01-01
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime, sSourceID,
                                                     ref mMsg);
                    if (mMsg.StartsWith("OK"))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestResult = TestConstants.TEST_FAIL;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[0].SimSpeed = 10000;
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TestScenario_TS300017TransportDestinationVia
        public void TestScenario_TS300017TransportDestinationVia(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transtype:" + "DROP";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  viaStationID:" + sStationID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].Automatic();
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Pick at 0040-01-01 by AGV11 with priority 5.
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    // Pick at 0040-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, null, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "PICK", sSourceID, null,
                                                sTestAgvs[0].ID.ToString(), 5);

                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;

                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Wait until TransportA state FINISHED, 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(1)" + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.	Set testAgv SimSpeed to 500
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    sTestAgvs[0].SimSpeed = 500;
                    mLogger.LogMessageToFile("" + sTestAgvs[0].ID
                                             + " simulation speed is:" + sTestAgvs[0].SimSpeed);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS1;
                }
                //5.	Create TransportB Drop at 0040-01-05 by AGV11 with priority 5
                //6.	Select Destination Via:  CX03
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        var transport = new Transport(Transport.COMMAND.DROP, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, null, sDestinationID, 5, false);
                        transport.DestinationViaLSID = sStationID;
                        transport.ID = "TransportB" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "DROP", null, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //7.	Wait until testAgv pass CX03
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mLogger.LogMessageToFile(sTestAgvs[0].ID + " is AT:"
                                             + sTestAgvs[0].CurrentLSID);

                    mTime = DateTime.Now - sTestStartTime;
                    string id = sTestAgvs[0].CurrentLSID.ToString();
                    //TestUtility.RemoteLogMessage(id, sTestMonitorUsed, m_Project);
                    if (sTestAgvs[0].CurrentLSID.ToString().Equals(sStationID))
                    {
                        mMsg = sTestAgvs[0].ID + "is via " + sStationID;
                        mLogger.LogMessageToFile(mMsg + " and speed again set to 10000");
                        sTestResult = TestConstants.TEST_PASS;
                        TestUtility.RemoteLogPassLine(sTestAgvs[0].ID + "is via " + sStationID, sTestMonitorUsed,
                                                      m_Project);
                        sTestAgvs[0].SimSpeed = 10000;
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_AFTER_RUNS;
                    }
                    else if (mTime.TotalMilliseconds >= 10*sWaitTime*60*1000) //  3*waitTime min
                    {
                        mMsg = sTestAgvs[0].ID + " at : " + id;
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    //else
                    //    mLogger.LogMessageToFile(sTestAgvs[0].ID.ToString() +" is at:"
                    //        +sTestAgvs[0].CurrentLSID.ToString());
                }
                //8.	Wait until drop transport TransportB finished.
                if (mTestStatus == TestConstants.TEST_AFTER_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_INITED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[0].SimSpeed = 10000;
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                sTestAgvs[0].SimSpeed = 10000;
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                //EndTestCaseAndUpdateFiles(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
            }
        }

        // TestScenario_TS241099WeekPlanBatteryCharge
        private void TestScenario_TS241099WeekPlanBatteryCharge(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[1].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[1].Automatic();
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create Agv11 Battery ChargePlan: today and current time + 2 with duration 2 min
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    var wp = new WeekPlan();
                    sTestStartTime = DateTime.Now.AddMinutes(2);
                    wp.ID = sTSName;
                    wp.Day = sTestStartTime.DayOfWeek;
                    wp.StartHour = sTestStartTime.Hour;
                    wp.StartMinute = sTestStartTime.Minute;
                    wp.Duration = 2; // 2 min
                    wp.Arguments.Add(sLocationID);
                    wp.Enable();
                    mLogger.LogMessageToFile(" Weekplan created: batt job will be created at "
                                             + sTestStartTime.ToLongTimeString());

                    sTestAgvs[1].BatteryChargePlan.Deactivate();
                    sTestAgvs[1].BatteryChargePlan.Clear();
                    sTestAgvs[1].BatteryChargePlan.Add(wp, true);
                    sTestAgvs[1].BatteryChargePlan.Activate();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Wait until Batt job started
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    string jobID = string.Empty;
                    Object[] jobs = sTestAgvs[1].Jobs.GetArray();
                    bool jobBatt = false;
                    //mLogger.LogMessageToFile("aantal jobs " + jobs.Length);
                    for (int i = 0; i < jobs.Length; i++)
                    {
                        var jobw = (Job) jobs[i];
                        string jtype = jobw.Type.ToString();
                        mLogger.LogMessageToFile("jobs " + jobw.ID + " with type:" + jtype);
                        if (jobw.Type.ToString().Equals("BATT") &&
                            jobw.LocationID.ToString().Equals(sLocationID))
                        {
                            jobBatt = true;
                            m_Job = jobw;
                        }
                    }

                    if (jobBatt)
                    {
                        mMsg = "batt job created";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                    else
                    {
                        if (mTime.TotalMilliseconds >= 300000) // 300 seconde
                        {
                            mMsg = "After 5 min no batt job created";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.	Wait until Agv11 at charging location
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[1], sTestStartTime, 3*sWaitTime,
                                                     sLocationID, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //5.	Check Agv11 state CHARGING
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        mMsg = sTestAgvs[1].ID + " has state : " + sTestAgvs[1].State.ToString();
                        if (sTestAgvs[1].State == Mover.STATE.CHARGING)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Wait until Agv11 return to Park location
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    if (m_Project.ID.ToString().ToUpper().StartsWith("EURO"))
                        TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[1], sTestStartTime, 3*sWaitTime,
                                                         sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString(), ref mMsg);
                    else
                        TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[1], sTestStartTime, 3*sWaitTime,
                                                         sLocationID, ref mMsg);

                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(m_Project.ID.ToString().ToUpper() + "  " + mMsg);
                            TestUtility.RemoteLogPassLine(m_Project.ID.ToString().ToUpper() + "  " + mMsg,
                                                          sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine((m_Project.ID.ToString().ToUpper() + "  " + mMsg));
                            TestUtility.RemoteLogFailLine(m_Project.ID.ToString().ToUpper() + "  " + mMsg,
                                                          sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS241047WeekPlanBatteryChargeDisable
        private void TestScenario_TS241047WeekPlanBatteryChargeDisable(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[1].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[1].Automatic();
                    ;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create Agv11 Battery Charge Plan: today and current time + 2 with duration  2 min
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    sTestStartTime = DateTime.Now.AddMinutes(2);
                    m_Wp.Deactivate();
                    m_Wp.Clear();
                    m_Wp.ID = sTSName;
                    m_Wp.Day = sTestStartTime.DayOfWeek;
                    m_Wp.StartHour = sTestStartTime.Hour;
                    m_Wp.StartMinute = sTestStartTime.Minute;
                    m_Wp.Duration = 2; // 2 min
                    m_Wp.Arguments.Add(sLocationID);
                    m_Wp.Enable();
                    m_Wp.Activate();

                    sTestAgvs[1].BatteryChargePlan.Deactivate();
                    sTestAgvs[1].BatteryChargePlan.Clear();
                    sTestAgvs[1].BatteryChargePlan.Add(m_Wp, true);
                    sTestAgvs[1].BatteryChargePlan.Activate();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Disable Agv11 Battery Charge Plan
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    m_Wp.Disable();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                //4.	Wait 3 minutes
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 180000) // wait 180 sec
                    {
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //5.	Check Agv11 no Batt job is created
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    string jobID = string.Empty;
                    Object[] jobs = sTestAgvs[1].Jobs.GetArray();
                    bool jobBatt = false;
                    for (int i = 0; i < jobs.Length; i++)
                    {
                        var jobw = (Job) jobs[i];
                        string jtype = jobw.Type.ToString();
                        if (jobw.Type.ToString().Equals("WAIT"))
                        {
                            jobBatt = true;
                            m_Job = jobw;
                        }
                    }

                    if (jobBatt)
                    {
                        mMsg = "batt job created, impossible: battery charge plan is disabled";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        mMsg = "OK, no batt job created.";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //6.	Check Agv11 park location
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    if (sTestAgvs[1].CurrentLSID.ToString().Equals(sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString()))
                    {
                        mMsg = "Agv11 at park location";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        mMsg = "Agv11 not at park location";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS241048WeekPlanBatteryChargeDisableAll
        private void TestScenario_TS241048WeekPlanBatteryChargeDisableAll(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[1].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[1].Automatic();
                    ;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create Agv11 Battery Charge Plan: today and current time + 2 with duration 2 min
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    var wp1 = new WeekPlan();
                    sTestStartTime = DateTime.Now.AddMinutes(2);
                    wp1.ID = sTSName;
                    wp1.Day = sTestStartTime.DayOfWeek;
                    wp1.StartHour = sTestStartTime.Hour;
                    wp1.StartMinute = sTestStartTime.Minute;
                    wp1.Duration = 2; // 2 min
                    wp1.Arguments.Add(sLocationID);
                    wp1.Enable();

                    var wp2 = new WeekPlan();
                    wp2.ID = (sTSName + 2);
                    wp2.Day = sTestStartTime.AddDays(1).DayOfWeek;
                    wp2.StartHour = sTestStartTime.Hour;
                    wp2.StartMinute = sTestStartTime.Minute;
                    wp2.Duration = 2; // 2 min
                    wp2.Arguments.Add(sLocationID);
                    wp2.Enable();

                    var wp3 = new WeekPlan();
                    wp3.ID = (sTSName + 3);
                    wp3.Day = sTestStartTime.AddDays(2).DayOfWeek;
                    wp3.StartHour = sTestStartTime.Hour;
                    wp3.StartMinute = sTestStartTime.Minute;
                    wp3.Duration = 2; // 2 min
                    wp3.Arguments.Add(sLocationID);
                    wp3.Enable();

                    var wp4 = new WeekPlan();
                    wp4.ID = (sTSName + 4);
                    wp4.Day = sTestStartTime.AddDays(3).DayOfWeek;
                    wp4.StartHour = sTestStartTime.Hour;
                    wp4.StartMinute = sTestStartTime.Minute;
                    wp4.Duration = 2; // 2 min
                    wp4.Arguments.Add(sLocationID);
                    wp4.Enable();

                    var wp5 = new WeekPlan();
                    wp5.ID = (sTSName + 5);
                    wp5.Day = sTestStartTime.AddDays(4).DayOfWeek;
                    wp5.StartHour = sTestStartTime.Hour;
                    wp5.StartMinute = sTestStartTime.Minute;
                    wp5.Duration = 2; // 1 min
                    wp5.Arguments.Add(sLocationID);
                    wp5.Enable();

                    var wp6 = new WeekPlan();
                    wp6.ID = (sTSName + 6);
                    wp6.Day = sTestStartTime.AddDays(5).DayOfWeek;
                    wp6.StartHour = sTestStartTime.Hour;
                    wp6.StartMinute = sTestStartTime.Minute;
                    wp6.Duration = 2; // 2 min
                    wp6.Arguments.Add(sLocationID);
                    wp6.Enable();

                    var wp7 = new WeekPlan();
                    wp7.ID = (sTSName + 7);
                    wp7.Day = sTestStartTime.AddDays(6).DayOfWeek;
                    wp7.StartHour = sTestStartTime.Hour;
                    wp7.StartMinute = sTestStartTime.Minute;
                    wp7.Duration = 2; // 2 min
                    wp7.Arguments.Add(sLocationID);
                    wp7.Enable();

                    sTestAgvs[1].BatteryChargePlan.Deactivate();
                    sTestAgvs[1].BatteryChargePlan.Clear();
                    sTestAgvs[1].BatteryChargePlan.Add(wp1, true);
                    sTestAgvs[1].BatteryChargePlan.Add(wp2, true);
                    sTestAgvs[1].BatteryChargePlan.Add(wp3, true);
                    sTestAgvs[1].BatteryChargePlan.Add(wp4, true);
                    sTestAgvs[1].BatteryChargePlan.Add(wp5, true);
                    sTestAgvs[1].BatteryChargePlan.Add(wp6, true);
                    sTestAgvs[1].BatteryChargePlan.Add(wp7, true);
                    sTestAgvs[1].BatteryChargePlan.Activate();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Disable Agv11 Battery Charge Plan
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        sTestAgvs[1].BatteryChargePlan.Disable();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //4.	Wait 3 minutes
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 180000) // wait 180 sec
                    {
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //5.	Check Agv11 no Batt job is created
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    string jobID = string.Empty;
                    Object[] jobs = sTestAgvs[1].Jobs.GetArray();
                    bool jobBatt = false;
                    for (int i = 0; i < jobs.Length; i++)
                    {
                        var jobw = (Job) jobs[i];
                        string jtype = jobw.Type.ToString();
                        if (jobw.Type.ToString().Equals("WAIT"))
                        {
                            jobBatt = true;
                            m_Job = jobw;
                        }
                    }

                    if (jobBatt)
                    {
                        mMsg = "batt job created, impossible: battery charge plan is disabled";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        mMsg = "OK, no batt job created.";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //6.	Check Agv11 park location
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    if (sTestAgvs[1].CurrentLSID.ToString().Equals(sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString()))
                    {
                        mMsg = "Agv11 at park location";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        mMsg = "Agv11 not at park location";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS241049WeekPlanBatteryChargeDelete
        private void TestScenario_TS241049WeekPlanBatteryChargeDelete(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[1].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[1].Automatic();
                    ;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create Agv11 Battery Charge Plan: today and current time + 2 with duration 2 min
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    //Egemin.EPIA.Core.Definitions.WeekPlan wp = new WeekPlan();
                    sTestStartTime = DateTime.Now.AddMinutes(2);
                    m_Wp.Clear();
                    m_Wp.ID = sTSName;
                    m_Wp.Day = sTestStartTime.DayOfWeek;
                    m_Wp.StartHour = sTestStartTime.Hour;
                    m_Wp.StartMinute = sTestStartTime.Minute;
                    m_Wp.Duration = 2; // 2 min
                    m_Wp.Arguments.Add(sLocationID);
                    m_Wp.Enable();

                    sTestAgvs[1].BatteryChargePlan.Deactivate();
                    sTestAgvs[1].BatteryChargePlan.Clear();
                    sTestAgvs[1].BatteryChargePlan.Add(m_Wp, true);
                    sTestAgvs[1].BatteryChargePlan.Activate();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Delete Agv11 Battery Charge Plan
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        sTestAgvs[1].BatteryChargePlan.Deactivate();
                        sTestAgvs[1].BatteryChargePlan.Remove(sTSName);
                        sTestAgvs[1].BatteryChargePlan.Activate();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //4.	Wait 3 minutes
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 180000) // wait 180 sec
                    {
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //5.	Check Agv11 no Batt job created
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    string jobID = string.Empty;
                    Object[] jobs = sTestAgvs[1].Jobs.GetArray();
                    bool jobBatt = false;
                    for (int i = 0; i < jobs.Length; i++)
                    {
                        var jobw = (Job) jobs[i];
                        string jtype = jobw.Type.ToString();
                        if (jobw.Type.ToString().Equals("WAIT"))
                        {
                            jobBatt = true;
                            m_Job = jobw;
                        }
                    }

                    if (jobBatt)
                    {
                        mMsg = "batt job created, impossible: battery charge plan is disabled";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        mMsg = "OK, no batt job created.";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //6.	Check Agv11 location
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    if (sTestAgvs[1].CurrentLSID.ToString().Equals(sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString()))
                    {
                        mMsg = "Agv11 at park location";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        mMsg = "Agv11 not at park location";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                //7.	Check Battery Charge Plan not in list of Battery Charge Plan
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    Object[] plans = sTestAgvs[1].BatteryChargePlan.GetArray();
                    bool inList = false;
                    for (int i = 0; i < plans.Length; i++)
                    {
                        var wp = (WeekPlan) plans[i];
                        if (wp.ID.ToString().StartsWith(sTSName))
                        {
                            inList = true;
                            break;
                        }
                    }

                    if (inList)
                    {
                        mMsg = "plan is still in list, impossible: battery charge plan is deleted";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        mMsg = "OK, no battery charge plan with ID start with:" + sTSName;
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS242099WeekPlanCalibration
        private void TestScenario_TS242099WeekPlanCalibration(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[1].ID;
                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[1].Automatic();
                    ;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create Calibration Plan A: today and current time + 2 with args CX01
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    m_Wp = new WeekPlan();
                    sTestStartTime = DateTime.Now.AddMinutes(1);
                    //m_Wp.Deactivate();
                    m_Wp.Clear();
                    m_Wp.ID = sTSName;
                    m_Wp.Day = sTestStartTime.DayOfWeek;
                    m_Wp.StartHour = sTestStartTime.Hour;
                    m_Wp.StartMinute = sTestStartTime.Minute;
                    m_Wp.Duration = 2; // 2 min
                    m_Wp.Arguments.Add(sLocationID);
                    m_Wp.Enable();
                    //m_Wp.Activate();

                    sTestAgvs[1].CalibrationPlan.Deactivate();
                    sTestAgvs[1].CalibrationPlan.Clear();
                    sTestAgvs[1].CalibrationPlan.Add(m_Wp, true);
                    sTestAgvs[1].CalibrationPlan.Activate();
                    mMsg = "wp = ";
                    object[] w = m_Wp.Arguments.GetArray();
                    for (int i = 0; i < w.Length; i++)
                    {
                        mMsg = mMsg + " has arg: " + w[i];
                    }
                    mMsg = mMsg + "  hour: " + m_Wp.StartHour;
                    mMsg = mMsg + "  min: " + m_Wp.StartMinute;
                    mLogger.LogPassLine(mMsg);
                    TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);


                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Wait until Agv11 state Executing
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    mMsg = sTestAgvs[1].ID + " has state :" + sTestAgvs[1].State.ToString();
                    if (sTestAgvs[1].State == Mover.STATE.EXECUTING)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                    else
                    {
                        if (mTime.TotalMilliseconds >= 180000) // 180 seconde
                        {
                            mMsg = "After 3 min " + mMsg;
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.	Check Agv11 pass CX01
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[1], sTestStartTime, sWaitTime,
                                                     sLocationID, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS242047WeekPlanCalibrationDisable
        private void TestScenario_TS242047WeekPlanCalibrationDisable(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[1].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[1].Automatic();
                    ;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create Calibration Plan A: today and current time + 2 with args CX01
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    sTestStartTime = DateTime.Now.AddMinutes(1);
                    m_Wp.Deactivate();
                    m_Wp.Clear();
                    m_Wp.ID = sTSName;
                    m_Wp.Day = sTestStartTime.DayOfWeek;
                    m_Wp.StartHour = sTestStartTime.Hour;
                    m_Wp.StartMinute = sTestStartTime.Minute;
                    m_Wp.Duration = 2; // 2 min
                    m_Wp.Arguments.Add(sLocationID);
                    m_Wp.Enable();
                    m_Wp.Activate();

                    sTestAgvs[1].CalibrationPlan.Deactivate();
                    sTestAgvs[1].CalibrationPlan.Clear();
                    sTestAgvs[1].CalibrationPlan.Add(m_Wp, true);
                    sTestAgvs[1].CalibrationPlan.Activate();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Disable Agv11 Calibration Plan A 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    m_Wp.Disable();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                //4.	Wait 3 min
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 180000) // wait 180 sec
                    {
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //5.	Check Agv11 no Wait job is created
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    string jobID = string.Empty;
                    Object[] jobs = sTestAgvs[1].Jobs.GetArray();
                    bool jobWait = false;
                    for (int i = 0; i < jobs.Length; i++)
                    {
                        var jobw = (Job) jobs[i];
                        string jtype = jobw.Type.ToString();
                        if (jobw.Type.ToString().Equals("WAIT"))
                        {
                            jobWait = true;
                            m_Job = jobw;
                        }
                    }

                    if (jobWait)
                    {
                        mMsg = "wait job created, impossible: calibration plan is disabled";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        mMsg = "OK, no wait job created.";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //6.	Check Agv11 park location
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    if (sTestAgvs[1].CurrentLSID.ToString().Equals(sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString()))
                    {
                        mMsg = "Agv11 at park location";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        mMsg = "Agv11 not at park location";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS242048WeekPlanCalibrationDisableAll
        private void TestScenario_TS242048WeekPlanCalibrationDisableAll(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[1].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[1].Automatic();
                    ;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create Calibration Plan A: today and current time + 2 with args CX01
                //3.	Create Calibration Plan B: today and current time + 3 with args CX01
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    sTestStartTime = DateTime.Now.AddMinutes(1);
                    m_Wp.Deactivate();
                    m_Wp.Clear();
                    m_Wp.ID = sTSName;
                    m_Wp.Day = sTestStartTime.DayOfWeek;
                    m_Wp.StartHour = sTestStartTime.Hour;
                    m_Wp.StartMinute = sTestStartTime.Minute;
                    m_Wp.Duration = 2; // 2 min
                    m_Wp.Arguments.Add(sLocationID);
                    m_Wp.Enable();
                    m_Wp.Activate();

                    sTestStartTime = DateTime.Now.AddMinutes(2);
                    m_Wp2.Deactivate();
                    m_Wp2.Clear();
                    m_Wp2.ID = sTSName;
                    m_Wp2.Day = sTestStartTime.DayOfWeek;
                    m_Wp2.StartHour = sTestStartTime.Hour;
                    m_Wp2.StartMinute = sTestStartTime.Minute;
                    m_Wp2.Duration = 1; // 1 min
                    m_Wp2.Arguments.Add(sLocationID);
                    m_Wp2.Enable();
                    m_Wp2.Activate();

                    sTestAgvs[1].CalibrationPlan.Deactivate();
                    sTestAgvs[1].CalibrationPlan.Clear();
                    sTestAgvs[1].CalibrationPlan.Add(m_Wp, true);
                    sTestAgvs[1].CalibrationPlan.Add(m_Wp2, true);
                    sTestAgvs[1].CalibrationPlan.Activate();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //4.	Disable Agv11 Calibration Plan All 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    sTestAgvs[1].CalibrationPlan.Disable();
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                //5.	Wait 4 min
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 240000) // wait 240 sec
                    {
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //6.	Check no Agv11 wait job created
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    string jobID = string.Empty;
                    Object[] jobs = sTestAgvs[1].Jobs.GetArray();
                    bool jobWait = false;
                    for (int i = 0; i < jobs.Length; i++)
                    {
                        var jobw = (Job) jobs[i];
                        string jtype = jobw.Type.ToString();
                        if (jobw.Type.ToString().Equals("WAIT"))
                        {
                            jobWait = true;
                            m_Job = jobw;
                        }
                    }

                    if (jobWait)
                    {
                        mMsg = "wait job created, impossible: calibration plan is disabled";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        mMsg = "OK, no wait job created.";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //7.	Check Agv11 park location
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    if (sTestAgvs[1].CurrentLSID.ToString().Equals(sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString()))
                    {
                        mMsg = "Agv11 at park location";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        mMsg = "Agv11 not at park location";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS242049WeekPlanCalibrationDelete
        private void TestScenario_TS242049WeekPlanCalibrationDelete(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[1].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[1].Automatic();
                    ;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create Calibration Plan A: today and current time + 2 with args CX01
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    sTestStartTime = DateTime.Now.AddMinutes(1);
                    m_Wp.Deactivate();
                    m_Wp.Clear();
                    m_Wp.ID = sTSName;
                    m_Wp.Day = sTestStartTime.DayOfWeek;
                    m_Wp.StartHour = sTestStartTime.Hour;
                    m_Wp.StartMinute = sTestStartTime.Minute;
                    m_Wp.Duration = 2; // 2 min
                    m_Wp.Arguments.Add(sLocationID);
                    m_Wp.Enable();
                    m_Wp.Activate();

                    sTestAgvs[1].CalibrationPlan.Deactivate();
                    sTestAgvs[1].CalibrationPlan.Clear();
                    sTestAgvs[1].CalibrationPlan.Add(m_Wp, true);
                    sTestAgvs[1].CalibrationPlan.Activate();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                    mLogger.LogMessageToFile("weekplan count:" + sTestAgvs[1].CalibrationPlan.Count);
                }
                //3.	Delete Agv11 Calibration Plan All 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        mLogger.LogMessageToFile("weekplan count after added:" + sTestAgvs[1].CalibrationPlan.Count);

                        sTestAgvs[1].CalibrationPlan.Deactivate();
                        sTestAgvs[1].CalibrationPlan.Remove(sTSName);
                        sTestAgvs[1].CalibrationPlan.Activate();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //4.	Wait 3 min
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 180000) // wait 180 sec
                    {
                        mLogger.LogMessageToFile("weekplan count after remove:" + sTestAgvs[1].CalibrationPlan.Count);

                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //5.	Check no Agv11 wait job created
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    string jobID = string.Empty;
                    Object[] jobs = sTestAgvs[1].Jobs.GetArray();
                    bool jobWait = false;
                    for (int i = 0; i < jobs.Length; i++)
                    {
                        var jobw = (Job) jobs[i];
                        string jtype = jobw.Type.ToString();
                        if (jobw.Type.ToString().Equals("WAIT"))
                        {
                            mLogger.LogFailLine("wait job:" + jobw.ID + " has state:" + jobw.State.ToString());
                            jobWait = true;
                            m_Job = jobw;
                        }
                    }

                    if (jobWait)
                    {
                        mMsg = "wait job created, impossible: calibration plan is deleted";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        mMsg = "OK, no wait job created.";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }
                //6.	Check Agv11 park location
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    if (sTestAgvs[1].CurrentLSID.ToString().Equals(sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString()))
                    {
                        mMsg = "Agv11 at park location";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        mMsg = "Agv11 not at park location";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                //7.	Check Calibration Plan A not in list of Calibration Plan
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    Object[] plans = sTestAgvs[1].CalibrationPlan.GetArray();
                    bool inList = false;
                    for (int i = 0; i < plans.Length; i++)
                    {
                        var wp = (WeekPlan) plans[i];
                        if (wp.ID.ToString().StartsWith(sTSName))
                        {
                            inList = true;
                            break;
                        }
                    }

                    if (inList)
                    {
                        mMsg = "plan is still in list, impossible: calibration plan is deleted";
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        mMsg = "OK, no calibration plan with ID start with:" + sTSName;
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
            }
        }

        // TestScenario_TS200083AgvStop
        private void TestScenario_TS200083AgvStop(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    sTextTestData = sTextTestData + "  testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  TransType:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, sDestinationID, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //3.	During picking, Stop Agv11
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        sTestAgvs[0].Softstop();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //4.	Check Agv11 is stopped
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        sAgvCurrentLSID = sTestAgvs[0].CurrentLSID.ToString();
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }

                //check agv is not moving
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.CheckAgvNotMoving(ref mRunStatus, sTestAgvs[0], sAgvCurrentLSID,
                                                  sTestStartTime, 10 /*sec*/, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //5.	Check Active Status List contains ¡®Softstop¡¯
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    string st = string.Empty;
                    int MyPos = 1;
                    Int16 iCount = 0;

                    for (int i = 0; i < sTestAgvs[0].ActiveStatusList.GetArray().Length; i++)
                    {
                        st = st + sTestAgvs[0].ActiveStatusList.GetArray()[i] + " ; ";
                    }

                    do
                    {
                        MyPos = st.ToUpper().IndexOf("SOFTSTOP", MyPos + 1);
                        if (MyPos > 0)
                        {
                            iCount += 1;
                        }
                    } while (!(MyPos == -1));

                    mMsg = "ActiveStatusList is " + st + " and found count " + iCount;

                    if (iCount > 0)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS200071AgvState
        private void TestScenario_TS200071AgvState(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    sTextTestData = sTextTestData + "  testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode Remove
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.Removed();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        sTestAgvs[0].Automatic();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //3 Check state "not ready"
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        TestUtility.WaitUntilAgvState(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime,
                                                      Mover.STATE.NOT_READY, ref mMsg);
                        if (mRunStatus == TestConstants.CHECK_END)
                        {
                            if (mMsg.StartsWith("OK"))
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                mLogger.LogPassLine(mMsg);
                                TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                                mTestStatus = TestConstants.TEST_RUNS;
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mLogger.LogFailLine(mMsg);
                                TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                                mTestStatus = TestConstants.TEST_FINISHED;
                            }
                        }
                    }
                }
                //4.	Restart Agv11
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    sTestAgvs[0].Restart();
                    sTestAgvs[1].Restart();
                    sTestAgvs[2].Restart();
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS1;
                }
                //5.	Check Agv11 state: Not ready ? Initialising ? Ready 
                //Check state "INIT"
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilAgvState(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime,
                                                  Mover.STATE.INIT, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            sTestAgvs[0].Release();
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //Check state "READY"
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilAgvState(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime,
                                                  Mover.STATE.READY, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);


                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Create JobA : Batt at PARK_BAT by AGV11
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    Job job = TestUtility.CreateTestJob("JobA" + sTSName, "BATT", sLocationID, sTSName, sProjectID);
                    m_Job = sTestAgvs[0].NewJob(job);
                    mLogger.LogCreatedJob("JobA" + sTSName, "BATT", sLocationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS4;
                }
                //7.	Check Agv11 state: Ready ? Executing ? Charging ? Ready charging 
                //Check state "EXECUTING"
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.WaitUntilAgvState(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime,
                                                  Mover.STATE.EXECUTING, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                //Check state "CHARGING"
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    TestUtility.WaitUntilAgvState(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime,
                                                  Mover.STATE.CHARGING, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS6;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //Check state "READY_CHARGING"
                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    TestUtility.WaitUntilAgvState(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime,
                                                  Mover.STATE.READY_CHARGING, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);

                            for (int i = 0; i < sTestAgvs.Length; i++)
                            {
                                sTestAgvs[i].SemiAutomatic();
                            }

                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS200037AgvRetire
        private void TestScenario_TS200037AgvRetire(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transtype:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationD:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Agv ParkID:" + sAgvsInitialID[sTestAgvs[0].ID.ToString()];

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Set Agv11 Retired
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    sTestAgvs[0].Retire();
                    mLogger.LogMessageToFile(sTestAgvs[0].ID + " is retired");
                    TestUtility.RemoteLogMessage(sTestAgvs[0].ID + " is retired", sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, sDestinationID, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //4.	Check TransportA state VERIFIED
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    // Wait until Transport state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.VERIFIED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //5.	Check TransportA wait reason
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        if (
                            m_Transport.WaitTransitionReason.Symbol.StartsWith(
                                "TRANSPORT_MANAGER_TRANSPORT_MOVER_NOT_AVAILABLE"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            mMsg = "TransportA wait reason" + m_Transport.WaitTransitionReason.Arguments[0].ToUpper();
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Check Agv11 at park location
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        if (m_Transport.State <= Transport.STATE.PENDING &&
                            sTestAgvs[0].CurrentLSID.ToString().Equals(
                                sAgvsInitialID[sTestAgvs[0].ID.ToString()].ToString()))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("");
                            TestUtility.RemoteLogPassLine("", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = "transport state:" + m_Transport.State.ToString();
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //get agv current lsid
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // Wait 3 sec.
                    {
                        sAgvCurrentLSID = sTestAgvs[0].CurrentLSID.ToString();
                        mMsg = sTestAgvs[0].ID + " current lsid id:: " + sAgvCurrentLSID;
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                }
                //7.	Check Agv11 not moving
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.CheckAgvNotMoving(ref mRunStatus, sTestAgvs[0], sAgvCurrentLSID,
                                                  sTestStartTime, 50 /*sec*/, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //8.	Check Agv11 Retired True
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mMsg = sTestAgvs[0].ID + " retired is " + sTestAgvs[0].Retired;
                    if (sTestAgvs[0].Retired)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS200053AgvDeploy
        private void TestScenario_TS200053AgvDeploy(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transtype:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationD:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Agv ParkID:" + sAgvsInitialID[sTestAgvs[0].ID.ToString()];

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                // 1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Set Agv11 Retired
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    sTestAgvs[0].Retire();
                    mLogger.LogMessageToFile(sTestAgvs[0].ID + " is retired");
                    TestUtility.RemoteLogMessage(sTestAgvs[0].ID + " is retired", sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, sDestinationID, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //4.	Check TransportA state VERIFIED
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    // Wait until Transport state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.VERIFIED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //5.	Check TransportA wait reason
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        if (
                            m_Transport.WaitTransitionReason.Symbol.StartsWith(
                                "TRANSPORT_MANAGER_TRANSPORT_MOVER_NOT_AVAILABLE"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            mMsg = "TransportA wait reason" + m_Transport.WaitTransitionReason.Arguments[0].ToUpper();
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Check Agv11 at park location
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        if (m_Transport.State <= Transport.STATE.PENDING &&
                            sTestAgvs[0].CurrentLSID.ToString().Equals(
                                sAgvsInitialID[sTestAgvs[0].ID.ToString()].ToString()))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("");
                            TestUtility.RemoteLogPassLine("", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = "transport state:" + m_Transport.State.ToString();
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //get agv current lsid
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // Wait 3 sec.
                    {
                        sAgvCurrentLSID = sTestAgvs[0].CurrentLSID.ToString();
                        mMsg = sTestAgvs[0].ID + " current lsid id:: " + sAgvCurrentLSID;
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                }
                //7.	Check Agv11 not moving
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.CheckAgvNotMoving(ref mRunStatus, sTestAgvs[0], sAgvCurrentLSID,
                                                  sTestStartTime, 50 /*sec*/, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //8.	Check Agv11 Retired True
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mMsg = sTestAgvs[0].ID + " retired is " + sTestAgvs[0].Retired;
                    if (sTestAgvs[0].Retired)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS6;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                //9.	Set Agv11 Deploy true
                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        sTestAgvs[0].Deploy();
                        mLogger.LogMessageToFile(sTestAgvs[0].ID + " is deployed");
                        TestUtility.RemoteLogMessage(sTestAgvs[0].ID + " is deployed", sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS7;
                    }
                }
                //10.	Check Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_RUNS7)
                {
                    mMsg = sTestAgvs[0].ID + " has mode  " + sTestAgvs[0].Mode.ToString();
                    if (sTestAgvs[0].Mode == Mover.MODE.AUTOMATIC)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS8;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }
                //11.	Wait until TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS8)
                {
                    // Wait until Transport state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS200023AgvSuspend(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transtype:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationD:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Agv ParkID:" + sAgvsInitialID[sTestAgvs[0].ID.ToString()];

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        m_Transport = TestUtility.CreateTestTransport("MOVE", sTestAgvs[0], sSourceID, sDestinationID,
                                                                      sTSName, ref m_Project, ref m_Transport);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //3.	During picking, Suspend Agv11
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        sTestAgvs[0].Suspend();
                        mLogger.LogMessageToFile(sTestAgvs[0].ID + " is suspended");
                        TestUtility.RemoteLogMessage(sTestAgvs[0].ID + " is suspended", sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //4.	Wait 10 sec
                //5.	Check Agv11 not moving
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        if (sTestAgvs[0].IsStopped())
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(sTestAgvs[0].ID + " is stoped");
                            TestUtility.RemoteLogPassLine(sTestAgvs[0].ID + " is stoped", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = "agv state: : " + sTestAgvs[0].State.ToString();
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Check Agv11 Suspended True
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        if (sTestAgvs[0].Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(sTestAgvs[0].ID + " is suspended");
                            TestUtility.RemoteLogPassLine(sTestAgvs[0].ID + " is suspended", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = "agv state: : " + sTestAgvs[0].State.ToString();
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        sTestAgvs[0].Release();
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                if (sTestAgvs[0].IsStopped())
                    sTestAgvs[0].Release();
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS200027AgvRelease(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "transtype:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationD:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Agv ParkID:" + sAgvsInitialID[sTestAgvs[0].ID.ToString()];

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mLogger.LogMessageToFile(sTestAgvs[0].ID + " has mode:" + sTestAgvs[0].Mode.ToString());
                    mLogger.LogMessageToFile(sTestAgvs[0].ID + " has state:" + sTestAgvs[0].State.ToString());
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        m_Transport = TestUtility.CreateTestTransport("MOVE", sTestAgvs[0], sSourceID, sDestinationID,
                                                                      sTSName, ref m_Project, ref m_Transport);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //3.	During picking, Suspend Agv11
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        sTestAgvs[0].Suspend();
                        mLogger.LogMessageToFile(sTestAgvs[0].ID + " is suspended");
                        TestUtility.RemoteLogMessage(sTestAgvs[0].ID + " is suspended", sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }
                //4.	Wait 10 sec
                //5.	Check Agv11 not moving
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        if (sTestAgvs[0].IsStopped())
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(sTestAgvs[0].ID + " is stoped");
                            TestUtility.RemoteLogPassLine(sTestAgvs[0].ID + " is stoped", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = "agv state: : " + sTestAgvs[0].State.ToString();
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Check Agv11 Suspended True
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        if (sTestAgvs[0].Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(sTestAgvs[0].ID + " is suspended");
                            TestUtility.RemoteLogPassLine(sTestAgvs[0].ID + " is suspended", sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = "agv state: : " + sTestAgvs[0].State.ToString();
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //7.	Release Agv11
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    sTestAgvs[0].Release();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS3;
                }
                //8.	Check Agv11 Suspend FALSE
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        if (!sTestAgvs[0].Suspended)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(sTestAgvs[0].ID + " is release");
                            TestUtility.RemoteLogPassLine(sTestAgvs[0].ID + " is release", sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = "agv state: : " + sTestAgvs[0].State.ToString();
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //9.	Wait until TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS200012AgvModeRemoved(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "Removed Agv:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "job type:" + "PARK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);

                    sTextTestData = sTextTestData + "  LocationID:" + sAgvsInitialID[sTestAgvs[1].ID.ToString()];
                    sTextTestData = sTextTestData + "  Removed Agv ParkID:" + sAgvsInitialID[sTestAgvs[1].ID.ToString()];
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode Removed
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[1].Removed(); // test2Agv removed
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 
                //3.	Check TransportA state VERIFIED
                //4.	Check TransportA wait reason
                //5.	Check Agv11 not exist
                // will continue WORKING

                // (2) Create Park Job at sLocationID
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    Job job = TestUtility.CreateTestJob("JobA" + sTSName, "PARK",
                                                        sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString(), sTSName,
                                                        sProjectID);
                    TestUtility.RemoteLogMessage(" job carrier id : " + job.CarrierID, sTestMonitorUsed, m_Project);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA" + sTSName, "PARK",
                                          sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString(),
                                          sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                // (3) Wait until JobA Finished
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);

                            string agvid = "agvid Are:";
                            // check agv
                            for (int i = 0; i < m_Project.Agvs.GetArray().Length; i++)
                            {
                                var agv = (Agv) m_Project.Agvs.GetArray()[i];
                                agvid = agvid + agv.ID + Environment.NewLine;
                            }
                            mLogger.LogFailLine(agvid);
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) //wait  3 sec
                    {
                        if (
                            sTestAgvs[0].CurrentLSID.ToString().Equals(
                                sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString()))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(sTestAgvs[0].ID + " is at " + sAgvsInitialID[sTestAgvs[1].ID.ToString()]);
                            TestUtility.RemoteLogPassLine(
                                sTestAgvs[0].ID + " is at " + sAgvsInitialID[sTestAgvs[1].ID.ToString()],
                                sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS1;
                            ;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = sTestAgvs[0].ID + " is at " + sAgvsInitialID[sTestAgvs[1].ID.ToString()];
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        sTestStartTime = DateTime.Now;
                    }
                }


                // (2) Create Park Job at sLocationID
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    sTestAgvs[0].SemiAutomatic();
                    Job job = TestUtility.CreateTestJob("JobB" + sTSName, "PARK",
                                                        sAgvsInitialID[sTestAgvs[0].ID.ToString()].ToString(), sTSName,
                                                        sProjectID);
                    TestUtility.RemoteLogMessage(" job carrier id : " + job.CarrierID, sTestMonitorUsed, m_Project);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobB" + sTSName, "PARK",
                                          sAgvsInitialID[sTestAgvs[0].ID.ToString()].ToString(),
                                          sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS2;
                }
                // (3) Wait until JobA Finished
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[1].Automatic();
                    sTestAgvs[1].Restart();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_FINISHED1;
                }


                if (mTestStatus == TestConstants.TEST_FINISHED1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 30000) // 30 seconde
                    {
                        EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    }
                }
            }
            catch (Exception ex)
            {
                sTestAgvs[1].Automatic();
                sTestAgvs[1].Restart();
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS200014AgvModeRemovedAll(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "Agv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "     Agv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "     Agv3:" + sTestAgvs[2].ID;
                    sTextTestData = sTextTestData + "   Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "   SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "   DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "   Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "   Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "   Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "   Destination3ID:" + sDestination3ID;

                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    // all agvs removed
                    for (int i = 0; i < m_Project.Agvs.GetArray().Length; i++)
                    {
                        var agv = (Agv) m_Project.Agvs.GetArray()[i];
                        agv.Removed();
                    }
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        //Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, sDestinationID, 5, false);
                        transport.ID = "TransportA";
                        m_Transport = m_Project.TransportManager.NewTransport(transport);

                        var transport2 = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[1].ID.ToString(),
                                                       null, null, sSource2ID, sDestination2ID, 5, false);
                        transport2.ID = "TransportB";
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport2);

                        var transport3 = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[2].ID.ToString(),
                                                       null, null, sSource3ID, sDestination3ID, 5, false);
                        transport3.ID = "TransportC";
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport3);

                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "MOVE", sSource2ID, sDestination2ID,
                                                    sTestAgvs[1].ID.ToString(), 5);
                        mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "MOVE", sSource3ID, sDestination3ID,
                                                    sTestAgvs[2].ID.ToString(), 5);

                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        mMsg = m_Transport.ID + " has state " + m_Transport.State.ToString();
                        if (m_Transport.State == Transport.STATE.VERIFIED &&
                            m_Transport2.State == Transport.STATE.VERIFIED &&
                            m_Transport3.State == Transport.STATE.VERIFIED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }


                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        if (
                            m_Transport.WaitTransitionReason.Symbol.StartsWith(
                                "TRANSPORT_MANAGER_TRANSPORT_MOVER_NOT_AVAILABLE") &&
                            m_Transport2.WaitTransitionReason.Symbol.StartsWith(
                                "TRANSPORT_MANAGER_TRANSPORT_MOVER_NOT_AVAILABLE") &&
                            m_Transport3.WaitTransitionReason.Symbol.StartsWith(
                                "TRANSPORT_MANAGER_TRANSPORT_MOVER_NOT_AVAILABLE"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            mMsg = "TransportA wait reason" + m_Transport.WaitTransitionReason.Arguments[0].ToUpper();
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        // check agvs exist
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    // all agvs mode automatic
                    for (int i = 0; i < m_Project.Agvs.GetArray().Length; i++)
                    {
                        var agv = (Agv) m_Project.Agvs.GetArray()[i];
                        agv.Automatic();
                        agv.Restart();
                        //sTestAgvs[i].Restart();
                    }
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_FINISHED1;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 60000) // 60 seconde
                    {
                        EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    }
                }
            }
            catch (Exception ex)
            {
                m_Project.Agvs.Automatic();
                for (int i = 0; i < sTestAgvs.Length; i++)
                {
                    sTestAgvs[i].Restart();
                }
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS200047AgvModeDisable(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "Agv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "   Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].Disabled(); // testAgv disabled
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }


                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        TestUtility.RemoteLogPassLine(sTestAgvs[0].ID + " mode is: " + sTestAgvs[0].Mode.ToString(),
                                                      sTestMonitorUsed, m_Project);
                        if (sTestAgvs[0].IsDisabled())
                            TestUtility.RemoteLogPassLine(
                                sTestAgvs[0].ID + " is disable: " + sTestAgvs[0].IsDisabled(), sTestMonitorUsed,
                                m_Project);
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    //Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        mMsg = m_Transport.ID + " has state " + m_Transport.State.ToString();
                        if (m_Transport.State == Transport.STATE.VERIFIED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        mMsg = m_Transport.ID + " has wait reason 0:" + m_Transport.WaitTransitionReason.Symbol;
                        if (
                            m_Transport.WaitTransitionReason.Symbol.StartsWith(
                                "TRANSPORT_MANAGER_TRANSPORT_MOVER_NOT_AVAILABLE"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                sTestAgvs[0].Automatic();
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS200048AgvModeDisableAll(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "Agv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "   Agv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "   Agv3:" + sTestAgvs[2].ID;
                    sTextTestData = sTextTestData + "   Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  Destination3ID:" + sDestination3ID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.Disabled(); // ALL Agv disabled	
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        TestUtility.RemoteLogPassLine(sTestAgvs[0].ID + " mode is: " + sTestAgvs[0].Mode.ToString(),
                                                      sTestMonitorUsed, m_Project);
                        if (sTestAgvs[0].IsDisabled())
                            TestUtility.RemoteLogPassLine(sTestAgvs[0].ID + " is disable: " + sTestAgvs[0].IsDisabled(),
                                                          sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    //Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString(), 5);

                    //Move at 0070-01-02-01-01 to 0360-01-01 by AGV3 with priority 5.
                    var transport2 = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[1].ID.ToString(), null,
                                                   null, sSource2ID, sDestination2ID, 5, false);
                    transport2.ID = "TransportB" + sTSName;
                    m_Transport2 = m_Project.TransportManager.NewTransport(transport2);
                    mLogger.LogCreatedTransport(transport2.ID.ToString(), "MOVE", sSource2ID, sDestination2ID,
                                                sTestAgvs[1].ID.ToString(), 5);

                    //Move at 0070-01-31-01-01 to 0360-01-01 by AGV7 with priority 5.
                    var transport3 = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[2].ID.ToString(), null,
                                                   null, sSource3ID, sDestination3ID, 5, false);
                    transport3.ID = "TransportC" + sTSName;
                    m_Transport3 = m_Project.TransportManager.NewTransport(transport3);
                    mLogger.LogCreatedTransport(transport3.ID.ToString(), "MOVE", sSource3ID, sDestination3ID,
                                                sTestAgvs[2].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;

                    mTestStatus = TestConstants.TEST_RUNS;
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        mMsg = m_Transport.ID + " has state " + m_Transport.State.ToString()
                               + " * " + m_Transport2.ID + " has state " + m_Transport2.State.ToString()
                               + " * " + m_Transport3.ID + " has state " + m_Transport3.State.ToString();
                        if (m_Transport.State == Transport.STATE.VERIFIED &&
                            m_Transport2.State == Transport.STATE.VERIFIED &&
                            m_Transport3.State == Transport.STATE.VERIFIED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        mMsg = m_Transport.ID + " has wait reason 0:" + m_Transport.WaitTransitionReason.Symbol
                               + " * " + m_Transport2.ID + " has wait reason 0:" +
                               m_Transport2.WaitTransitionReason.Symbol
                               + " * " + m_Transport3.ID + " has wait reason 0:" +
                               m_Transport3.WaitTransitionReason.Symbol;
                        if (
                            m_Transport.WaitTransitionReason.Symbol.StartsWith(
                                "TRANSPORT_MANAGER_TRANSPORT_MOVER_NOT_AVAILABLE") &&
                            m_Transport2.WaitTransitionReason.Symbol.StartsWith(
                                "TRANSPORT_MANAGER_TRANSPORT_MOVER_NOT_AVAILABLE") &&
                            m_Transport3.WaitTransitionReason.Symbol.StartsWith(
                                "TRANSPORT_MANAGER_TRANSPORT_MOVER_NOT_AVAILABLE"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }


                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                m_Project.Agvs.Automatic();
                for (int i = 0; i < sTestAgvs.Length; i++)
                    sTestAgvs[i].Restart();
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        // TestScenario_TS200058AgvModeSemiAutomatic
        private void TestScenario_TS200058AgvModeSemiAutomatic(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "   Agv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "   Trans Type:" + "MOVE";
                    sTextTestData = sTextTestData + "   Job Type:" + "WAIT";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  sLocationID:" + sLocationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agvs mode All automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].Automatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Move from 0070-01-01-01-01 to 360-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  "LoadA", sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport("TransportA" + sTSName, "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Wait until TransportA state RETRIEVED 
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.RETRIEVED,
                                                        ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.	Set Agvs Agv11 mode semi-automatic
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    sTestAgvs[0].SemiAutomatic();
                    mTestStatus = TestConstants.TEST_RUNS1;
                }
                //5.	Create JobA Wait at W0420-01 by AGV11 
                // W0450-02  replaced W0420-01, its route cost is cheaper
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    Job job = TestUtility.CreateTestJob("JobA" + sTSName, "WAIT", sLocationID, sTSName, sProjectID);
                    TestUtility.RemoteLogMessage(" job carrier id : " + job.CarrierID, sTestMonitorUsed, m_Project);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA" + sTSName, "WAIT", sLocationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS2;
                }
                //6.	Wait 10 sec
                //7.	Set Agvs  mode All automatic 
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // 10 seconde
                    {
                        sTestAgvs[0].Automatic();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                }
                //8.	Wait until JobA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //9.	Wait until TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED,
                                                        ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //10.	Check Agv11 at location 0360-01-01
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mMsg = sTestAgvs[0].ID + "  at location: " + sTestAgvs[0].CurrentLSID;
                    if (sTestAgvs[0].CurrentLSID.ToString().Equals(sDestinationID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS200061AgvModeSemiAutomaticAll(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "   Agv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "   Agv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "   Trans Type:" + "MOVE";
                    sTextTestData = sTextTestData + "   Job Type:" + "WAIT";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sLocation2ID = sTestInputParams["sLocation2ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  sLocationID:" + sLocationID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  sLocation2ID:" + sLocation2ID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agvs mode All automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.Automatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Move from 0070-01-01-01-01 to 360-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  "LoadA", sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport("TransportA" + sTSName, "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Create TransportB Move from 0070-01-02-01-01 to 360-01-02 by AGV3
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[1].ID.ToString(), null,
                                                  "LoadB", sSource2ID, sDestination2ID, 5, false);
                    transport.ID = "TransportB" + sTSName;
                    m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport("TransportB" + sTSName, "MOVE", sSource2ID, sDestination2ID,
                                                sTestAgvs[1].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                //4.	Wait until TransportA and TransportB state ASSIGNED
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.ASSIGNED,
                                                        ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.ASSIGNED,
                                                        ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //5.	Set Agvs mode All semi-automatic
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    m_Project.Agvs.SemiAutomatic();
                    mTestStatus = TestConstants.TEST_RUNS3;
                }
                //6.	Create JobA Wait at W0420-01 by AGV11 
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    Job job = TestUtility.CreateTestJob("JobA" + sTSName, "WAIT", sLocationID, sTSName, sProjectID);
                    TestUtility.RemoteLogMessage(" job carrier id : " + job.CarrierID, sTestMonitorUsed, m_Project);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA" + sTSName, "WAIT", sLocationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS4;
                }
                //7.	Create JobB Wait at W0450-02 by AGV3 
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    Job job = TestUtility.CreateTestJob("JobB" + sTSName, "WAIT", sLocation2ID, sTSName, sProjectID);
                    TestUtility.RemoteLogMessage(" job carrier id : " + job.CarrierID, sTestMonitorUsed, m_Project);
                    m_Job2 = m_Project.Agvs[sTestAgvs[1].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobB" + sTSName, "WAIT", sLocation2ID, sTestAgvs[1].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS5;
                }

                //10.	Wait until JobA  and JobB state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS6;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job2, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS7;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                //8.	Wait 10 sec
                //9.	Set Agvs  mode All automatic 
                if (mTestStatus == TestConstants.TEST_RUNS7)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // 10 seconde
                    {
                        m_Project.Agvs.Automatic();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS8;
                    }
                }

                //11.	Wait until TransportA and TransportB state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS8)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED,
                                                        ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS9;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS9)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED,
                                                        ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //12.	Check Agv11 at location 0360-01-01
                /* if (mTestStatus == TestConstants.TEST_AFTER_RUNS)
				{
					mMsg = sTestAgvs[0].ID.ToString() + "  at location: " + sTestAgvs[0].CurrentLSID.ToString();
					if (sTestAgvs[0].CurrentLSID.ToString().Equals(sDestinationID))
					{
						sTestResult = TestConstants.TEST_PASS;
						mLogger.LogPassLine(mMsg);
						TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
						mTestStatus = TestConstants.TEST_AFTER;
					}
					else
					{
						sTestResult = TestConstants.TEST_FAIL;
						mLogger.LogFailLine(mMsg);
						TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
						mTestStatus = TestConstants.TEST_FINISHED;
					}
					
				}
				//13.	Check Agv3 at location 0360-01-02
				if (mTestStatus == TestConstants.TEST_AFTER)
				{
					mMsg = sTestAgvs[1].ID.ToString() + "  at location: " + sTestAgvs[1].CurrentLSID.ToString();
					if (sTestAgvs[1].CurrentLSID.ToString().Equals(sDestination2ID))
					{
						sTestResult = TestConstants.TEST_PASS;
						mLogger.LogPassLine(mMsg);
						TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
					}
					else
					{
						sTestResult = TestConstants.TEST_FAIL;
						mLogger.LogFailLine(mMsg);
						TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
					}
					mTestStatus = TestConstants.TEST_FINISHED;
				}
				*/
                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                m_Project.Agvs.Automatic();
                for (int i = 0; i < sTestAgvs.Length; i++)
                    sTestAgvs[i].Restart();

                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS309306TransPickDeactiveRestart(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    //Pick at 0070-01-01-01-01 to by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(),
                                                  null, null, sSourceID, null, 5, false);
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    sTestStartTime = DateTime.Now;
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED) //deactivate testAgv
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        sTestAgvs[0].Deactivate();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING) // restart testAgv
                {
                    sTestAgvs[0].Activate();
                    sTestAgvs[0].Restart();
                    mRunStatus = TestConstants.CONTINUE;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS;
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(1)" + mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mMsg = sTestAgvs[0].ID + " at " + sTestAgvs[0].CurrentLSID;
                    ;
                    if (sTestAgvs[0].CurrentLSID.ToString().Equals(sSourceID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS309306TransDropDeactiveRestart(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;

                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "DROP";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    sTestAgvs[0].Automatic();

                    //Pick at 0070-01-01-01-01 to by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(),
                                                  null, null, sSourceID, null, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);

                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(1)" + mMsg, sTestMonitorUsed, m_Project);

                            //Drop at 0360-01-01 to by AGV11 with priority 5.
                            var transport = new Transport(Transport.COMMAND.DROP, null, null, sTestAgvs[0].ID.ToString(),
                                                          null, null, null, sDestinationID, 5, false);
                            transport.ID = "TransportB" + sTSName;
                            m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_INITED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 9000) // 9 seconde
                    {
                        sTestAgvs[0].Deactivate();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }


                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 9000) // 9 seconde
                    {
                        sTestAgvs[0].Activate();
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 9000) // 9 seconde
                    {
                        sTestAgvs[0].Restart();
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(1)" + mMsg, sTestMonitorUsed, m_Project);

                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mMsg = sTestAgvs[0].ID + " at " + sTestAgvs[0].CurrentLSID;
                    ;
                    if (sTestAgvs[0].CurrentLSID.ToString().Equals(sDestinationID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS420047LocationDisable(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Agv ParkID:" + sAgvsInitialID[sTestAgvs[0].ID.ToString()];
                    ;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    m_Project.Locations[sSourceID].Disable();
                    mLogger.LogMessageToFile("location " + sSourceID + " is disabled");
                    TestUtility.RemoteLogMessage("location " + sSourceID + " is disabled", sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        m_Transport = TestUtility.CreateTestTransport("PICK", sTestAgvs[0], sSourceID,
                                                                      null, sTSName, ref m_Project, ref m_Transport);
                        mLogger.LogMessageToFile("Created transport: " + m_Transport.ID);
                        //mJobStatus = TestConstants.JOB_NOT_STARTED;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        if (m_Transport.State <= Transport.STATE.PENDING &&
                            sTestAgvs[0].CurrentLSID.ToString().Equals(
                                sAgvsInitialID[sTestAgvs[0].ID.ToString()].ToString()))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("");
                            TestUtility.RemoteLogPassLine("", sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = "transport state:" + m_Transport.State.ToString();
                            mLogger.LogFailLine(mMsg + Environment.NewLine
                                                + sTestAgvs[0].ID + " current position is:" + sTestAgvs[0].CurrentLSID +
                                                Environment.NewLine
                                                + sTestAgvs[0].ID + " initial position is:" +
                                                sAgvsInitialID[sTestAgvs[0].ID.ToString()] + Environment.NewLine
                                );
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        m_Project.Locations[sSourceID].Automatic();
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                sTestAgvs[0].Automatic();
                m_Project.Locations[sSourceID].Automatic();
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS420045LocationManual(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Job Type:" + "BATT";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    m_Project.Locations[sLocationID].Manual();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    mLogger.LogMessageToFile("location " + sLocationID + " mode is manual");
                    TestUtility.RemoteLogMessage("location " + sLocationID + " mode is manual", sTestMonitorUsed,
                                                 m_Project);
                    sTestStartTime = DateTime.Now;
                    //mJobStatus = TestConstants.JOB_NOT_STARTED;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                        mTestStatus = TestConstants.TEST_STARTING;
                }
                /*
				if (mTestStatus == TestConstants.TEST_STARTING)
				{
					CreateAndWaitUntilJobFinished(ref mJobStatus, sTestAgvs[0], "BATT", sLocationID, sTSName,
						ref sTestResult, ref mMsg);
					if (mJobStatus == TestConstants.JOB_FINISHED)
					{
						if (sTestResult == TestConstants.TEST_PASS)
						{
							mLogger.LogPassLine(mMsg);
							TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
						}
						else
						{
							mLogger.LogFailLine(mMsg);
							TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
						}
						sTestStartTime = DateTime.Now;
						mTestStatus = TestConstants.TEST_FINISHED;
					}
				}
				*/
                // (2) Create Park Job at sLocationID
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    sTestAgvs[0].SemiAutomatic();
                    Job job = TestUtility.CreateTestJob("JobA" + sTSName, "BATT", sLocationID, sTSName, sProjectID);
                    TestUtility.RemoteLogMessage(" job carrier id : " + job.CarrierID, sTestMonitorUsed, m_Project);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA" + sTSName, "BATT", sLocationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                // (3) Wait until JobA Finished
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    m_Project.Locations[sLocationID].Automatic();
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                try
                {
                    m_Project.Locations[sLocationID].Automatic();
                    mLogger.LogTestException(ex.Message, ex.StackTrace);
                    TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                }
                finally
                {
                    EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                }
            }
        }

        private void TestScenario_TS460047StationDisable(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "  Job Type:" + "BATT";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();

                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;
                    sTextTestData = sTextTestData + "  StationD:" + sStationID;
                    sTextTestData = sTextTestData + "  Agv ParkID:" + sAgvsInitialID[sTestAgvs[1].ID.ToString()];

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    m_Project.Stations[sStationID].Disable();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    mLogger.LogMessageToFile("station " + sStationID + " is disabled");
                    TestUtility.RemoteLogMessage("station " + sStationID + " is disabled", sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    //mJobStatus = TestConstants.JOB_NOT_STARTED;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                        mTestStatus = TestConstants.TEST_STARTING;
                }

                /*
				if (mTestStatus == TestConstants.TEST_STARTING)
				{
					CreateAndWaitUntilJobFinished(ref mJobStatus, sTestAgvs[1], "BATT", sLocationID, sTSName,
						ref sTestResult, ref mMsg);
					if (mJobStatus == TestConstants.JOB_FINISHED)
					{
						string testAgvLsid = sTestAgvs[1].CurrentLSID.ToString();
						mMsg = sTestAgvs[0].ID.ToString() + " is  at " + testAgvLsid;
						if (testAgvLsid.Equals(sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString()))
						{
							sTestResult = TestConstants.TEST_PASS;
							mLogger.LogPassLine(mMsg);
							TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
						}
						else
						{
							sTestResult = TestConstants.TEST_FAIL;
							mLogger.LogFailLine(mMsg);
							TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
						}
						m_Project.Stations[sStationID].Enable();
						sTestStartTime = DateTime.Now;
						mTestStatus = TestConstants.TEST_FINISHED;
					}
				}
				*/

                // (2) Create Park Job at sLocationID
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    sTestAgvs[0].SemiAutomatic();
                    Job job = TestUtility.CreateTestJob("JobA" + sTSName, "BATT", sLocationID, sTSName, sProjectID);
                    TestUtility.RemoteLogMessage(" job carrier id : " + job.CarrierID, sTestMonitorUsed, m_Project);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA" + sTSName, "BATT", sLocationID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }
                // (3) Wait until JobA Finished
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.FINISHED,
                                                  sTestStartTime, sWaitTime, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    string testAgvLsid = sTestAgvs[1].CurrentLSID.ToString();
                    mMsg = sTestAgvs[0].ID + " is  at " + testAgvLsid;
                    if (testAgvLsid.Equals(sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    m_Project.Stations[sStationID].Enable();
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    m_Project.Stations[sStationID].Enable();
                    mLogger.LogTestException(ex.Message, ex.StackTrace);
                    TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                }
                finally
                {
                    EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                }
            }
        }

        private void TestScenario_TS400034LoadFlushAndDiscard(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  LoadID:" + "LoadA";

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    //(1)Set testAgv mode automatic
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    mTestStatus = TestConstants.TEST_INITED;
                }


                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    // (2) Create TransportA
                    //Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  "LoadA", sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport("TransportA" + sTSName, "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_STARTING;
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    // (3) Wait until TransportA state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    // (4) Check LoadA state STORED
                    var load = (Load) m_Project.Loads.GetItem("LoadA");
                    mMsg = "(4) " + load.ID + "  has " + load.State.ToString();
                    if (load.State.ToString().ToUpper().Equals(Load.STATE.STORED.ToString().ToUpper()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    // (5)Flush LoadA
                    m_Project.Loads.Flush("LoadA");
                    TestUtility.RemoteLogPassLine("(5) " + "Flush " + "LoadA", sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS2;
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    // (6) Check loadA existence
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 120000) // wait 120 sec
                    {
                        if (m_Project.Loads.GetArray().Length == 0)
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine("LoadA" + " flushed with state store");
                            TestUtility.RemoteLogFailLine("LoadA" + " flushed with store store:",
                                                          sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mMsg = "(6) " + "LoadA" + " not all loads flushed:" + m_Project.Loads.GetArray().Length;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    // (7)Discard LoadA
                    var load = (Load) m_Project.Loads.GetItem("LoadA");
                    load.Discard();
                    TestUtility.RemoteLogPassLine("(7) " + "discard " + "LoadA", sTestMonitorUsed, m_Project);
                    mTestStatus = TestConstants.TEST_RUNS4;
                }

                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    // (8) Check LoadA state DISCARDED
                    var load = (Load) m_Project.Loads.GetItem("LoadA");
                    mMsg = "(8) " + load.ID + " has state: " + load.State.ToString();
                    if (load.State.ToString().ToUpper().Equals(Load.STATE.DISCARDED.ToString().ToUpper()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_RUNS5;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    // (9) Flush LoadA
                    m_Project.Loads.Flush("LoadA");
                    TestUtility.RemoteLogPassLine("(9) " + "Flush " + "LoadA", sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS6;
                }

                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    // (10) Check loadA existence
                    if (m_Project.Loads.GetArray().Length == 0)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine("(10) " + "LoadA" + " flushed with state discarded");
                        TestUtility.RemoteLogPassLine("(10) " + "LoadA" + " flushed with state discarded",
                                                      sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        mTime = DateTime.Now - sTestStartTime;
                        if (mTime.TotalMilliseconds >= 120000) // wait 120 sec
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = "(10) " + "LoadA" + " not flushed after 2 min:" + m_Project.Loads.GetArray().Length;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS400076LoadDiscard(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationD:" + sDestinationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestStartTime = DateTime.Now;
                    //mJobStatus = TestConstants.JOB_NOT_STARTED;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                /*
				if (mTestStatus == TestConstants.TEST_INITED)
				{
					CreateAndWaitUntilTransportFinished(ref mJobStatus, sTestAgvs[0], "MOVE", sSourceID, sDestinationID, sTSName,
						ref sTestResult, ref mMsg);
					if (mJobStatus == TestConstants.JOB_FINISHED)
					{
						if (sTestResult == TestConstants.TEST_PASS)
						{
							mLogger.LogPassLine(mMsg);
							TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
							mTestStatus = TestConstants.TEST_STARTING;
						}
						else
						{
							mLogger.LogFailLine(mMsg);
							TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
							mTestStatus = TestConstants.TEST_FINISHED; ;
						}
						sTestStartTime = DateTime.Now;
					}
				}
				*/
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, sDestinationID, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.RETRIEVED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    Object[] loads = m_Project.Loads.GetArray();
                    for (int i = 0; i < loads.Length; i++)
                    {
                        var load = (Load) loads[i];
                        load.Discard();
                    }
                    mTestStatus = TestConstants.TEST_RUNS1;
                    sTestStartTime = DateTime.Now;
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        bool discard = false;
                        Object[] loads = m_Project.Loads.GetArray();
                        for (int i = 0; i < loads.Length; i++)
                        {
                            var load = (Load) loads[i];
                            if (load.IsDiscarded())
                            {
                                discard = true;
                                break;
                            }
                        }

                        if (discard)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogFailLine("loads discarded");
                            TestUtility.RemoteLogFailLine("loads discarded", sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = "load not  discard";
                            ;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS305513TransOrderDelay(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);


                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "DELAY", "70");
                    TestUtility.AddSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);

                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        sTestResult = TestConstants.TEST_UNDEFINED;
                        mMsg = string.Empty;
                        m_Project.Agvs.SemiAutomatic();
                        sTestAgvs[0].Automatic();

                        //Move at 0070-01-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, null, 5, false);
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "PICK", sSourceID, null,
                                                    sTestAgvs[0].ID.ToString(), 5);

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_INITED;
                    }
                }

                //get agv current lsid
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // Wait 3 sec.
                    {
                        sAgvCurrentLSID = sTestAgvs[0].CurrentLSID.ToString();
                        mMsg = sTestAgvs[0].ID + " current lsid id:: " + sAgvCurrentLSID;
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                //check agv is not moving within 50 secs
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.CheckAgvNotMoving(ref mRunStatus, sTestAgvs[0], sAgvCurrentLSID,
                                                  sTestStartTime, 50 /*sec*/, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    // Wait until transport state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                //TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS305514TransOrderDivert(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[1].Automatic();
                    sTestAgvs[1].Suspend();

                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "DIVERT", sLocationID);

                    //Move at 0070-01-01 to 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[1].ID.ToString(), null,
                                                  null, sSourceID, sDestinationID, 5, false);
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[1].ID.ToString(), 5);

                    sTestAgvs[1].Release();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 30000) // wait 30 sec
                    {
                        mMsg = m_Transport.State.ToString();
                        if (m_Transport.State >= Transport.STATE.ASSIGNED)
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    m_Project.Transports.Cancel();
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                m_Project.Transports.Cancel();
                TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS304455LocationClosestHighest(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  DestID:" + sDestinationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    TestUtility.SetLocationPriority(ref m_Project, sSourceID, 0, ref mLogger);
                    TestUtility.SetLocationPriority(ref m_Project, sSource2ID, 8, ref mLogger);
                    TestUtility.SetLocationPriority(ref m_Project, sSource3ID, 6, ref mLogger);
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "HIGHEST", null);
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestAgvs[0].Suspend();
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    //Move at 0060-03-01 to 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);

                    // Move at 0070-01-1-01-01 to 0360-01-01 by AGV11 with priority 5.
                    transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null, null,
                                              sSource2ID, sDestinationID, 5, false);
                    transport.ID = "TransportB" + sTSName;
                    m_Transport2 = m_Project.TransportManager.NewTransport(transport);

                    // Move at 0040-01-1 to 0360-01-01 by AGV11 with priority 5.
                    transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null, null,
                                              sSource3ID, sDestinationID, 5, false);
                    transport.ID = "TransportC" + sTSName;
                    m_Transport3 = m_Project.TransportManager.NewTransport(transport);

                    sTestAgvs[0].Release();
                    sTestStartTime = DateTime.Now;

                    mTestStatus = TestConstants.TEST_STARTING;
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        mLogger.LogCreatedTransport("TransportA" + sTSName, "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        mLogger.LogCreatedTransport("TransportB" + sTSName, "MOVE", sSource2ID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        mLogger.LogCreatedTransport("TransportC" + sTSName, "MOVE", sSource3ID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    // Wait until transportB finish
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime,
                                                        sWaitTime /*min*/, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mRunStatus = TestConstants.CONTINUE;
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                /*
				if (mTestStatus == TestConstants.TEST_RUNS1)
				{
					// Wait until testAgv at Park location
					TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[0], sTestStartTime,
									 3 * sWaitTime, sAgvsInitialID[sTestAgvs[0].ID.ToString()].ToString(), ref mMsg);
					if (mRunStatus == TestConstants.CHECK_END)
					{
						if (mMsg.StartsWith("OK"))
						{
							sTestResult = TestConstants.TEST_PASS;
							mLogger.LogPassLine(mMsg);
							TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
							mRunStatus = TestConstants.CONTINUE;
							sTestStartTime = DateTime.Now;
							mTestStatus = TestConstants.TEST_RUNS2;
						}
						else
						{
							sTestResult = TestConstants.TEST_FAIL;
							mLogger.LogFailLine(mMsg);
							TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
							mTestStatus = TestConstants.TEST_FINISHED;
						}
					}
				}
				*/
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    // Wait until transportC finish
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport3, sTestStartTime,
                                                        3*sWaitTime /*min*/, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mRunStatus = TestConstants.CONTINUE;
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    // Wait until transportA finish
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        4*sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mRunStatus = TestConstants.CONTINUE;
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    DateTime timeA = m_Transport.Finished;
                    DateTime timeB = m_Transport2.Finished;
                    DateTime timeC = m_Transport3.Finished;
                    string bricht = m_Transport.ID + " finished time: " + timeA
                                    + Environment.NewLine
                                    + m_Transport2.ID + " finished time: " + timeB
                                    + Environment.NewLine
                                    + m_Transport3.ID + " finished time: " + timeC;

                    mTime = timeA - timeC;
                    TimeSpan difTime2 = timeC - timeB;
                    if (mTime.TotalMilliseconds > 0 && difTime2.TotalMilliseconds > 0)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(m_Transport.ID + " finished time: " + timeA);
                        mLogger.LogPassLine(m_Transport2.ID + " finished time: " + timeB);
                        mLogger.LogPassLine(m_Transport3.ID + " finished time: " + timeC);
                        TestUtility.RemoteLogPassLine(bricht, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(m_Transport.ID + " finished time: " + timeA);
                        mLogger.LogFailLine(m_Transport2.ID + " finished time: " + timeB);
                        mLogger.LogFailLine(m_Transport3.ID + " finished time: " + timeC);
                        TestUtility.RemoteLogFailLine(bricht, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    TestUtility.SetLocationPriority(ref m_Project, sSourceID, 0, ref mLogger);
                    TestUtility.SetLocationPriority(ref m_Project, sSource2ID, 0, ref mLogger);
                    TestUtility.SetLocationPriority(ref m_Project, sSource3ID, 0, ref mLogger);

                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    TestUtility.SetLocationPriority(ref m_Project, sSourceID, 0, ref mLogger);
                    TestUtility.SetLocationPriority(ref m_Project, sSource2ID, 0, ref mLogger);
                    TestUtility.SetLocationPriority(ref m_Project, sSource3ID, 0, ref mLogger);

                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    mLogger.LogTestException(ex.Message, ex.StackTrace);
                    TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                }
                finally
                {
                    EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                    ;
                }
            }
        }

        private void TestScenario_TS304457GroupHighestPriority(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sGroupID = sTestInputParams["sGroupID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sGroup2ID = sTestInputParams["sGroup2ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  groupID:" + sGroupID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  group2ID:" + sGroup2ID;

                    mLogger.LogMessageToFile("*******" + sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    TestUtility.SetGroupPriority(ref m_Project, sGroupID, 8, ref mLogger);
                    TestUtility.SetGroupPriority(ref m_Project, sGroup2ID, 1, ref mLogger);
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "HIGHEST", null);

                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestAgvs[0].Suspend();
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    //PICK at 0070-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, null, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    // PICK at 0030-01-1 by AGV11 with priority 5.
                    var transport2 = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(), null,
                                                   null, sSource2ID, null, 5, false);
                    transport2.ID = "TransportB" + sTSName;
                    m_Transport2 = m_Project.TransportManager.NewTransport(transport2);

                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "PICK", sSourceID, null,
                                                sTestAgvs[0].ID.ToString(), 5);
                    mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "PICK", sSource2ID, null,
                                                sTestAgvs[0].ID.ToString(), 5);

                    sTestAgvs[0].Release();
                    mRunStatus = TestConstants.CONTINUE;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    // Wait until transportA finish
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mRunStatus = TestConstants.CONTINUE;
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    string bricht = sTestAgvs[0].ID + " at: " + sTestAgvs[0].CurrentLSID;
                    mLogger.LogPassLine(bricht);
                    TestUtility.RemoteLogPassLine(bricht, sTestMonitorUsed, m_Project);

                    if (sTestAgvs[0].CurrentLSID.ToString().Equals(sSourceID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    TestUtility.SetGroupPriority(ref m_Project, sGroupID, 0, ref mLogger);
                    TestUtility.SetGroupPriority(ref m_Project, sGroup2ID, 0, ref mLogger);
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    mLogger.LogTestException(ex.Message, ex.StackTrace);
                    TestUtility.SetGroupPriority(ref m_Project, sGroupID, 0, ref mLogger);
                    TestUtility.SetGroupPriority(ref m_Project, sGroup2ID, 0, ref mLogger);
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);

                    TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                }
                finally
                {
                    EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                    ;
                }
            }
        }

        private void TestScenario_TS304456LoadClosestHighest(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTestAgvs[0].Automatic();
                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                    sTestStartTime = DateTime.Now;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                        sTestAgvs[0].Suspend();

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_INITED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        //Move from 0060-03-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null,
                                                      "LoadA", sSourceID, sDestinationID, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);

                        //Move from 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport2 = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                       null,
                                                       "LoadB", sSource2ID, sDestinationID, 5, false);
                        transport2.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport2);

                        //Move from 0040-01-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport3 = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                       null,
                                                       "LoadC", sSource3ID, sDestinationID, 5, false);
                        transport3.ID = "TransportC" + sTSName;
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport3);

                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "MOVE", sSource2ID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "MOVE", sSource3ID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);

                        TestUtility.RemoteLogPassLine("(2)Create transports", sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 8000) // wait 8 sec
                    {
                        m_Project.Transports.Deactivate();
                        m_Project.Loads.Deactivate();
                        m_Project.Loads["LoadA"].Priority = 1;
                        m_Project.Loads["LoadB"].Priority = 8;
                        m_Project.Loads["LoadC"].Priority = 1;
                        m_Project.Loads.Activate();
                        m_Project.Transports.Activate();
                        TestUtility.RemoteLogPassLine("(3)Set priority", sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 30000) // wait 30 sec
                    {
                        mMsg = string.Empty;
                        Loads loads = m_Project.Loads;
                        for (int i = 0; i < loads.Count; i++)
                        {
                            mMsg = mMsg + "   == load " + loads[i].ID + "  has priority:" + loads[i].Priority;
                        }
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestAgvs[0].Release();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilAllTransportFinished(ref mRunStatus, m_Project, sTestStartTime, 3*sWaitTime,
                                                              ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        sTestStartTime = DateTime.Now;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    DateTime timeA = m_Transport.Finished;
                    DateTime timeB = m_Transport2.Finished;
                    DateTime timeC = m_Transport3.Finished;
                    string bricht = m_Transport.ID + " finished time: " + timeA
                                    + Environment.NewLine
                                    + m_Transport2.ID + " finished time: " + timeB
                                    + Environment.NewLine
                                    + m_Transport3.ID + " finished time: " + timeC;

                    mTime = timeA - timeC;
                    TimeSpan difTime2 = timeC - timeB;
                    if (mTime.TotalMilliseconds > 0 && difTime2.TotalMilliseconds > 0)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(m_Transport.ID + " finished time: " + timeA);
                        mLogger.LogPassLine(m_Transport2.ID + " finished time: " + timeB);
                        mLogger.LogPassLine(m_Transport3.ID + " finished time: " + timeC);
                        TestUtility.RemoteLogPassLine(bricht, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogPassLine(m_Transport.ID + " finished time: " + timeA);
                        mLogger.LogPassLine(m_Transport2.ID + " finished time: " + timeB);
                        mLogger.LogPassLine(m_Transport2.ID + " finished time: " + timeC);
                        TestUtility.RemoteLogFailLine(bricht, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    TestUtility.SetLocationPriority(ref m_Project, sSourceID, 0, ref mLogger);
                    TestUtility.SetLocationPriority(ref m_Project, sSource2ID, 0, ref mLogger);
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    mLogger.LogTestException(ex.Message, ex.StackTrace);
                    TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                }
                finally
                {
                    EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                }
            }
        }

        private void TestScenario_TS305501OrderAssignmentClosest(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  testAgv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "  testAgv3:" + sTestAgvs[2].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  Destination3ID:" + sDestination3ID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestAgvs[0].Automatic();
                    sTestAgvs[1].Automatic();
                    sTestAgvs[2].Automatic();
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST", null);
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    //Move at 0240-01-01 to 0360-01-01
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sSourceID, sDestinationID);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE",
                                                    sSourceID, sDestinationID, "AGV to be assigned",
                                                    m_Transport.Priority);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mMsg = "agvID: " + sTestAgvs[0].ID + " and Transport MoverID: " + m_Transport.MoverID;
                    if (m_Transport.MoverID.ToString().Equals(sTestAgvs[0].ID.ToString()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[0], sTestStartTime,
                                                     2*sWaitTime, sAgvsInitialID[sTestAgvs[0].ID.ToString()].ToString(),
                                                     ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 6000) // wait 6 sec
                    {
                        //Move at 0060-01-01 to 0360-01-01
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sSource2ID, sDestination2ID);
                        transport.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 6000) // wait 6 sec
                    {
                        mLogger.LogCreatedTransport("TransportB" + sTSName, "MOVE",
                                                    sSource2ID, sDestination2ID, "Agv to be assigned",
                                                    m_Transport2.Priority);

                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS5;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            mTestStatus = TestConstants.TEST_RUNS6;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    mMsg = "agvID: " + sTestAgvs[1].ID + " and Transport MoverID: " + m_Transport2.MoverID;
                    if (m_Transport2.MoverID.ToString().Equals(sTestAgvs[1].ID.ToString()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS7;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS7)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[1], sTestStartTime,
                                                     2*sWaitTime, sAgvsInitialID[sTestAgvs[1].ID.ToString()].ToString(),
                                                     ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS8;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS8)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 60000) // wait 60 sec
                    {
                        //Move at 0060-13-01 to 0360-01-01
                        var transport3 = new Transport(Transport.COMMAND.MOVE, null, null, sSource3ID, sDestination3ID);
                        transport3.ID = "TransportC" + sTSName;
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport3);
                        mTestStatus = TestConstants.TEST_RUNS9;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS9)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 6000) // wait 6 sec
                    {
                        // mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "MOVE",
                        //       sSource3ID, sDestination3ID, m_Transport3.AssignedMoverID.ToString(), m_Transport3.Priority);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_AFTER_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_AFTER_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport3, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            mTestStatus = TestConstants.TEST_AFTER;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_AFTER)
                {
                    mMsg = "agvID: " + sTestAgvs[2].ID + " and Transport MoverID: " + m_Transport3.MoverID;
                    if (m_Transport3.MoverID.ToString().Equals(sTestAgvs[2].ID.ToString()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS305502TransOrderClosest(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  Destination3ID:" + sDestination3ID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST", null);
                    mMsg = string.Empty;
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestAgvs[0].Suspend();
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    //Move at 0060-03-01 to 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                  null, null, sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);

                    //Move at 0070-01-01 to 0360-01-01 by AGV11 with priority 5.
                    transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                              null, null, sSource2ID, sDestination2ID, 5, false);
                    transport.ID = "TransportB" + sTSName;
                    m_Transport2 = m_Project.TransportManager.NewTransport(transport);

                    // Move at 0040-01-01 to 0360-01-01 by AGV11 with priority 5.
                    transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                              null, null, sSource3ID, sDestination3ID, 5, false);
                    transport.ID = "TransportC" + sTSName;
                    m_Transport3 = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "MOVE", sSource2ID, sDestination2ID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "MOVE", sSource3ID, sDestination3ID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        sTestAgvs[0].Release();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport3, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport3.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport2.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2*sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }


                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    DateTime timeA = m_Transport.Finished;
                    DateTime timeB = m_Transport2.Finished;
                    DateTime timeC = m_Transport3.Finished;
                    string bricht = m_Transport.ID + " finished time: " + timeA
                                    + Environment.NewLine
                                    + m_Transport2.ID + " finished time: " + timeB
                                    + Environment.NewLine
                                    + m_Transport3.ID + " finished time: " + timeC;

                    mTime = timeA - timeB;
                    TimeSpan difTime2 = timeB - timeC;
                    if (mTime.TotalMilliseconds > 0 && difTime2.TotalMilliseconds > 0)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(m_Transport.ID + " finished time: " + timeA);
                        mLogger.LogPassLine(m_Transport2.ID + " finished time: " + timeB);
                        mLogger.LogPassLine(m_Transport3.ID + " finished time: " + timeC);
                        TestUtility.RemoteLogPassLine(bricht, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogPassLine(m_Transport.ID + " finished time: " + timeA);
                        mLogger.LogPassLine(m_Transport2.ID + " finished time: " + timeB);
                        mLogger.LogPassLine(m_Transport3.ID + " finished time: " + timeC);
                        TestUtility.RemoteLogFailLine(bricht, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS305503OrderAssignmentClosestHighest(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "testAgv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "testAgv3:" + sTestAgvs[2].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  Destination3ID:" + sDestination3ID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestAgvs[1].Automatic();
                    sTestAgvs[2].Automatic();
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    //Move at 0240-01-01 to 0360-01-01
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sSourceID, sDestinationID);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_STARTING;
                }


                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        mLogger.LogCreatedTransport("TransportA" + sTSName, "MOVE",
                                                    sSourceID, sDestinationID, "AGV TO BE Assigned ",
                                                    m_Transport.Priority);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    // Wait until Transport state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("111" + mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mMsg = "agvID: " + sTestAgvs[0].ID + " and Transport MoverID: " + m_Transport.MoverID;
                    if (m_Transport.MoverID.ToString().Equals(sTestAgvs[0].ID.ToString()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine("112" + mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 60000) // wait 60 sec
                    {
                        //Move at 0060-01-01 to 0360-01-01
                        var transport2 = new Transport(Transport.COMMAND.MOVE, null, null, sSource2ID, sDestination2ID);
                        transport2.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport2);
                        sTestStartTime = DateTime.Now;
                        mLogger.LogPassLine("113" + mMsg);
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "MOVE",
                                                    sSource2ID, sDestination2ID, m_Transport.AssignedMoverID.ToString(),
                                                    m_Transport.Priority);
                        sTestStartTime = DateTime.Now;
                        mLogger.LogPassLine("114" + mMsg);
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    // Wait until Transport state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("115" + mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mMsg = "agvID: " + sTestAgvs[1].ID + " and Transport MoverID: " + m_Transport2.MoverID;
                    if (m_Transport2.MoverID.ToString().Equals(sTestAgvs[1].ID.ToString()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine("116" + mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS6;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 60000) // wait 60 sec
                    {
                        mLogger.LogPassLine("117" + mMsg);
                        //Move at 0060-13-01 to 0360-01-01
                        var transport3 = new Transport(Transport.COMMAND.MOVE, null, null, sSource3ID, sDestination3ID);
                        transport3.ID = "TransportC" + sTSName;
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport3);
                        mTestStatus = TestConstants.TEST_RUNS7;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS7)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 60000) // wait 60 sec
                    {
                        //mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "MOVE",
                        //       sSource3ID, sDestination3ID, m_Transport3.AssignedMoverID.ToString(), m_Transport3.Priority);
                        //mLogger.LogPassLine("118" + mMsg);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS8;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS8)
                {
                    // Wait until Transport state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport3, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine("119" + mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS9;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS9)
                {
                    mMsg = "agvID: " + sTestAgvs[2].ID + " and Transport MoverID: " + m_Transport3.MoverID;
                    if (m_Transport3.MoverID.ToString().Equals(sTestAgvs[2].ID.ToString()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine("120" + mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS305504TransOrderClosestHighest(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  Destination3ID:" + sDestination3ID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    mMsg = string.Empty;
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestAgvs[0].Suspend();
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    //Move at 0060-03-01 to 0360-01-01 by AGV11 with priority 8.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                  null, null, sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);

                    //Move at 0070-01-01 to 0360-01-01 by AGV11 with priority 5.
                    transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                              null, null, sSource2ID, sDestination2ID, 8, false);
                    transport.ID = "TransportB" + sTSName;
                    m_Transport2 = m_Project.TransportManager.NewTransport(transport);

                    // Move at 0040-01-01 to 0360-01-01 by AGV11 with priority 5.
                    transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                              null, null, sSource3ID, sDestination3ID, 5, false);
                    transport.ID = "TransportC" + sTSName;
                    m_Transport3 = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "MOVE", sSource2ID, sDestination2ID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "MOVE", sSource3ID, sDestination3ID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        sTestAgvs[0].Release();
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport3, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }


                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    DateTime timeA = m_Transport.Finished;
                    DateTime timeB = m_Transport2.Finished;
                    DateTime timeC = m_Transport3.Finished;
                    string bricht = m_Transport.ID + " finished time: " + timeA
                                    + Environment.NewLine
                                    + m_Transport2.ID + " finished time: " + timeB
                                    + Environment.NewLine
                                    + m_Transport3.ID + " finished time: " + timeC;

                    mTime = timeA - timeB;
                    TimeSpan difTime2 = timeC - timeA;
                    if (mTime.TotalMilliseconds > 0 && difTime2.TotalMilliseconds > 0)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(m_Transport.ID + " finished time: " + timeA);
                        mLogger.LogPassLine(m_Transport2.ID + " finished time: " + timeB);
                        mLogger.LogPassLine(m_Transport3.ID + " finished time: " + timeC);
                        TestUtility.RemoteLogPassLine(bricht, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogPassLine(m_Transport.ID + " finished time: " + timeA);
                        mLogger.LogPassLine(m_Transport2.ID + " finished time: " + timeB);
                        mLogger.LogPassLine(m_Transport2.ID + " finished time: " + timeC);
                        TestUtility.RemoteLogFailLine(bricht, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS305505TransOrderOldest(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestination2ID = sTestInputParams["sDestination2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();
                    sDestination3ID = sTestInputParams["sDestination3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + sDestination2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;
                    sTextTestData = sTextTestData + "  Destination3ID:" + sDestination3ID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestAgvs[0].Suspend();

                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "OLDEST", null);
                    //Move at 0060-03-01 to 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportA";
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString(), 5);

                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 6000) // wait 6 sec
                    {
                        // Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSource2ID, sDestination2ID, 5, false);
                        transport.ID = "TransportB";
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSource2ID, sDestination2ID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 6000) // wait 6 sec
                    {
                        // Move at 0040-01-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSource3ID, sDestination3ID, 5, false);
                        transport.ID = "TransportC";
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSource3ID, sDestination3ID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestAgvs[0].Release();
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport3, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }


                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    DateTime timeA = m_Transport.Finished;
                    DateTime timeB = m_Transport2.Finished;
                    DateTime timeC = m_Transport3.Finished;
                    mTime = timeC - timeB;
                    TimeSpan difTime = timeB - timeA;
                    mMsg = " timeA: " + timeA + " -- timeB: " + timeB + " -- timeC: " + timeC;
                    if (mTime.TotalMilliseconds > 0 && difTime.TotalMilliseconds > 0)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS305507SchedulesDeadlockRulesVia(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "testAgv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();
                    sScheduleID = sTestInputParams["sScheduleID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Stop locationID:" + sLocationID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;
                    sTextTestData = sTextTestData + "  ViaStationID:" + sStationID;
                    sTextTestData = sTextTestData + "  Schedule:" + sScheduleID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    // Enable AREA30 deadlock schedule
                    m_Project.Schedules[sScheduleID].Enable();
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    m_Project.Agvs.Automatic();
                    //PICK at 0030-03-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, null, 5, false);
                    transport.ID = "TransportA" + sTSName;
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "PICK", sSourceID, null,
                                                sTestAgvs[0].ID.ToString(), 5);

                    mRunStatus = TestConstants.CONTINUE;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                            if (sTestAgvs[0].CurrentLSID.ToString().Equals(sSourceID))
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                mLogger.LogPassLine(mMsg);
                                TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                                //PICK at 0030-01-01 by AGV3 with priority 5.
                                var transport = new Transport(Transport.COMMAND.PICK, null, null,
                                                              sTestAgvs[1].ID.ToString(), null, null, sSourceID, null, 5,
                                                              false);
                                transport.ID = "TransportB" + sTSName;
                                m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                                mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "PICK", sSource2ID, null,
                                                            sTestAgvs[1].ID.ToString(), m_Transport.Priority);
                                mRunStatus = TestConstants.CONTINUE;
                                sTestStartTime = DateTime.Now;
                                mTestStatus = TestConstants.TEST_STARTING;
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mLogger.LogFailLine(mMsg);
                                TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                                mTestStatus = TestConstants.TEST_FINISHED;
                            }
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[1], sTestStartTime, sWaitTime,
                                                     sLocationID, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    //DROP at 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.DROP, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, null, sDestinationID, 5, false);
                    transport.ID = "TransportC" + sTSName;
                    m_Transport3 = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "DROP",
                                                null, sDestinationID, m_Transport.AssignedMoverID.ToString(),
                                                m_Transport.Priority);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS1;
                    sTestAgvs[0].SimSpeed = 500;
                    sTestAgvs[1].SimSpeed = 500;
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    if (sTestAgvs[0].DeadLocker)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mMsg = sTestAgvs[0].ID + " deadlock is True ";
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mLogger.LogPassLine(mMsg);
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                    else
                    {
                        TimeSpan timeDiff = DateTime.Now - sTestStartTime;
                        if (timeDiff.TotalMilliseconds >= 120000) // 120 sec
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = sTestAgvs[0].ID + " Deadlock not true during 2 min ";
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    if (sTestAgvs[0].DeadLockingIDs.GetArray()[0].ToString().Equals(sTestAgvs[1].ID.ToString()))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mMsg = sTestAgvs[0].ID + " deadlock id is" + sTestAgvs[1].ID;
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mLogger.LogPassLine(mMsg);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mMsg = sTestAgvs[0].ID + " has deadlockid : " + sTestAgvs[0].DeadLockingIDs.GetArray()[0];
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mLogger.LogFailLine(mMsg);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        // check Deadlock via station
                        TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[1], sTestStartTime, 2*sWaitTime,
                                                         sStationID, ref mMsg);
                        if (mRunStatus == TestConstants.CHECK_END)
                        {
                            if (mMsg.StartsWith("OK"))
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                                mLogger.LogPassLine(mMsg);
                                sTestStartTime = DateTime.Now;
                                mRunStatus = TestConstants.CONTINUE;
                                mTestStatus = TestConstants.TEST_RUNS4;
                                sTestAgvs[0].SimSpeed = 10000;
                                sTestAgvs[1].SimSpeed = 10000;
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                                mLogger.LogFailLine(mMsg);
                                mTestStatus = TestConstants.TEST_FINISHED;
                            }
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    // Wait until transportB finish
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mRunStatus = TestConstants.CONTINUE;
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    // Wait until transportC finish
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport3, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS6;
                            if (m_Project.ID.ToString().ToUpper().StartsWith("TESTOP"))
                                mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    mMsg = sTestAgvs[1].ID + " at location: " + sTestAgvs[1].CurrentLSID;
                    if (sTestAgvs[1].CurrentLSID.ToString().Equals(sSourceID))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        //REPEAT without Deadlock schedule
                        //DROP at 0360-01-01 by AGV3 with priority 5.
                        var transport = new Transport(Transport.COMMAND.DROP, null, null, sTestAgvs[1].ID.ToString(),
                                                      null, null, null, sDestinationID, 5, false);
                        transport.ID = "TransportD";
                        m_Transport = m_Project.TransportManager.NewTransport(transport);

                        mLogger.LogCreatedTransport(transport.ID.ToString(), "DROP",
                                                    null, sDestinationID, sTestAgvs[1].ID.ToString(),
                                                    m_Transport.Priority);
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_STOP;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STOP)
                {
                    // Wait until transportD finish
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;

                            // disable AREA30 deadlock schedule
                            m_Project.Schedules[sScheduleID].Disable();
                            mMsg = mMsg + "  " + sScheduleID + "  disabled";
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);

                            // repeat TransportA
                            //PICK at 0030-03-01 by AGV11 with priority 5.
                            var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(),
                                                          null, null, sSourceID, null, 5, false);
                            transport.ID = "TransportE";
                            m_Transport = m_Project.TransportManager.NewTransport(transport);
                            mLogger.LogCreatedTransport(transport.ID.ToString(), "PICK", sSourceID, null,
                                                        sTestAgvs[0].ID.ToString(), 5);

                            mRunStatus = TestConstants.CONTINUE;
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS7;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS7)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                            if (sTestAgvs[0].CurrentLSID.ToString().Equals(sSourceID))
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                mLogger.LogPassLine(mMsg);
                                TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                                //PICK at 0030-01-01 by AGV3 with priority 5.
                                var transport = new Transport(Transport.COMMAND.PICK, null, null,
                                                              sTestAgvs[1].ID.ToString(), null, null, sSourceID, null, 5,
                                                              false);
                                transport.ID = "TransportF";
                                m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                                mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "PICK", sSourceID, null,
                                                            sTestAgvs[1].ID.ToString(), m_Transport.Priority);
                                mRunStatus = TestConstants.CONTINUE;
                                sTestStartTime = DateTime.Now;
                                mTestStatus = TestConstants.TEST_RUNS8;
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mLogger.LogFailLine(mMsg);
                                TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                                mTestStatus = TestConstants.TEST_FINISHED;
                            }
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS8)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[1], sTestStartTime, sWaitTime,
                                                     sLocationID, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS9;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS9)
                {
                    //DROP at 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.DROP, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, null, sDestinationID, 5, false);
                    transport.ID = "TransportG";
                    m_Transport3 = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(m_Transport3.ID.ToString(), "DROP",
                                                null, sDestinationID, m_Transport.AssignedMoverID.ToString(),
                                                m_Transport.Priority);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_AFTER_RUNS;
                }

                if (mTestStatus == TestConstants.TEST_AFTER_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 2*60000) // wait 2*60 sec
                    {
                        if (sTestAgvs[0].DeadLocker)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mMsg = sTestAgvs[0].ID + " deadlock is True ";
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = sTestAgvs[0].ID + " Deadlock not true after 2 min ";
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[0].SimSpeed = 10000;
                    sTestAgvs[1].SimSpeed = 10000;
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                sTestAgvs[0].SimSpeed = 10000;
                sTestAgvs[1].SimSpeed = 10000;
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS305508ScheduleBattRulesQueueSimLow(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                {
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                    if (m_Project.ID.ToString().ToUpper().StartsWith("TESTOPSTEL"))
                    {
                        mLogger.LogMessageToFile("==== Test opstelling NOT continued==== ");
                        TestUtility.RemoteLogMessage("==== Test opstelling NOT continued==== ", sTestMonitorUsed,
                                                     m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  testAgv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "  testAgv3:" + sTestAgvs[2].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sLocationID = sTestInputParams["sLocationID"].ToString();

                    sTextTestData = sTextTestData + "  Disable locationID:" + sLocationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                    m_Project.Locations[sLocationID].Mode = Location.MODE.DISABLED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = "test started";
                    m_Project.Agvs.Automatic();
                    for (int i = 0; i < sTestAgvs.Length; i++)
                        sTestAgvs[i].SimBatteryLow = true;

                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    sChargingAgv = (Agv) m_Project.Agvs.GetArray()[0];
                    sQueueAgv = (Agv) m_Project.Agvs.GetArray()[0];
                    sWaitAgv = (Agv) m_Project.Agvs.GetArray()[0];
                    int counter = 0;
                    mMsg = string.Empty;
                    for (int i = 0; i < sTestAgvs.Length; i++)
                    {
                        mMsg = mMsg + "== " + sTestAgvs[i].ID + " at : "
                               + sTestAgvs[i].CurrentLSID
                               + Environment.NewLine;
                        if (sTestAgvs[i].CurrentLSID.ToString().Equals("PARK_BAT"))
                        {
                            sChargingAgv = sTestAgvs[i];
                            counter++;
                        }
                        else if (sTestAgvs[i].CurrentLSID.ToString().Equals("X0500_100"))
                        {
                            sQueueAgv = sTestAgvs[i];
                            counter++;
                        }
                        else if (
                            sTestAgvs[i].CurrentLSID.ToString().Equals(
                                sAgvsInitialID[sTestAgvs[i].ID.ToString()].ToString()))
                        {
                            sWaitAgv = sTestAgvs[i];
                        }
                        else
                        {
                            counter = 100;
                            mMsg = mMsg + "  test failed Agv " + sTestAgvs[i].ID + " at  " +
                                   sTestAgvs[i].CurrentLSID;
                            break;
                        }
                    }

                    mMsg = mMsg + "  actual queue is " + counter + Environment.NewLine
                           + " and Charging Agv is " + sChargingAgv.ID + Environment.NewLine
                           + " and Queue Agv is " + sQueueAgv.ID + Environment.NewLine
                           + " and Wait Agv is " + sWaitAgv.ID;

                    if (counter == 2)
                    {
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mLogger.LogPassLine(mMsg);
                        sTestResult = TestConstants.TEST_PASS;
                        sTestStartTime = DateTime.Now;
                        sChargingAgv.SimBatteryLow = false;

                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                    else
                    {
                        mTime = DateTime.Now - sTestStartTime;
                        if (mTime.TotalMilliseconds >= 60000) // wait 60 sec
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mMsg = string.Empty;
                    string queueAgvLSID = sQueueAgv.CurrentLSID.ToString();
                    string waitAgvLSID = sWaitAgv.CurrentLSID.ToString();
                    mTime = DateTime.Now - sTestStartTime;

                    mMsg = mMsg + "==  after " + mTime.TotalMilliseconds/1000 + "sec "
                           + " Queue Agv is now at " + queueAgvLSID + Environment.NewLine
                           + " and Wait  Agv is now at " + waitAgvLSID;
                    if (queueAgvLSID.Equals("PARK_BAT")
                        && waitAgvLSID.Equals("X0500_100"))
                    {
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mLogger.LogPassLine(mMsg);
                        sTestResult = TestConstants.TEST_PASS;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        mTime = DateTime.Now - sTestStartTime;
                        if (mTime.TotalMilliseconds >= 120000) // wait 120 sec
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    if (!m_Project.ID.ToString().ToUpper().StartsWith("TESTOPSTEL"))
                    {
                        mLogger.LogMessageToFile("==== Test opstelling NOT continued2==== ");
                        TestUtility.RemoteLogMessage("==== Test opstelling NOT continued2==== ", sTestMonitorUsed,
                                                     m_Project);
                        m_Project.Locations[sLocationID].Mode = Location.MODE.AUTOMATIC;
                        for (int i = 0; i < sTestAgvs.Length; i++)
                            sTestAgvs[i].SimBatteryLow = false;
                    }
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                    //EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    m_Project.Locations[sLocationID].Mode = Location.MODE.AUTOMATIC;
                    for (int i = 0; i < sTestAgvs.Length; i++)
                        sTestAgvs[i].SimBatteryLow = false;

                    mLogger.LogTestException(ex.Message, ex.StackTrace);
                    TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                }
                finally
                {
                    EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                    ;
                    //EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData);
                }
            }
        }

        private void TestScenario_TS330056RoutingDynamic(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sLocationID = sTestInputParams["sLocationID"].ToString();
                    sStationID = sTestInputParams["sStationID"].ToString();

                    sTextTestData = sTextTestData + "  StationID:" + sStationID;
                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Destination:" + sDestinationID;
                    sTextTestData = sTextTestData + "  LocationID:" + sLocationID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        //Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, sDestinationID, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.RETRIEVED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 1000) // wait 1 sec
                    {
                        sTestAgvs[0].SimSpeed = 500;
                        TestUtility.RemoteLogMessage("set Speed = 500", sTestMonitorUsed, m_Project);
                        mLogger.LogMessageToFile("set Speed = 500");
                        mRunStatus = TestConstants.CONTINUE;
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[0], sTestStartTime, 2*sWaitTime,
                                                     sStationID, ref mMsg);
                    mLogger.LogMessageToFile("Current position = " + sTestAgvs[0].CurrentLSID);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            sTestAgvs[0].SimSpeed = 10000;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[0], sTestStartTime, sWaitTime,
                                                     sAgvsInitialID[sTestAgvs[0].ID.ToString()].ToString(), ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    m_Project.Router.Parameters["Dynamic"].ValueAsBool = true;
                    m_Project.Stations[sStationID].Disable();
                    mLogger.LogMessageToFile("station " + sStationID + " is disabled");
                    mTestStatus = TestConstants.TEST_RUNS4;
                }

                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    //Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, sDestinationID, 5, false);
                    transport.ID = "TransportB" + sTSName;
                    m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sSourceID, sDestinationID,
                                                sTestAgvs[0].ID.ToString(), 5);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS5;
                }

                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.RETRIEVED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport2.State == Transport.STATE.RETRIEVED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS6;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS6)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        sTestAgvs[0].SimSpeed = 500;
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS7;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS7)
                {
                    TestUtility.WaitUntilAgvPassNode(ref mRunStatus, sTestAgvs[0], sTestStartTime, 3*sWaitTime,
                                                     sLocationID, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            sTestAgvs[0].SimSpeed = 10000;
                            mTestStatus = TestConstants.TEST_RUNS8;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }


                if (mTestStatus == TestConstants.TEST_RUNS8)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport2.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    try
                    {
                        sTestAgvs[0].SimSpeed = 10000;
                        m_Project.Stations[sStationID].Enable();
                        mLogger.LogMessageToFile("station " + sStationID + " is enabled");
                    }
                    finally
                    {
                        EndTestCase(sTSName, sTestResult, ref sTextTestData);
                        //EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData); ;
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    m_Project.Stations[sStationID].Enable();
                    mLogger.LogMessageToFile("station " + sStationID + " is enabled");
                    mLogger.LogTestException(ex.Message, ex.StackTrace);
                    TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                    //EndTestCaseAndUpdateFiles( sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData );
                }
                finally
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
        }

        private void TestScenario_TS300063TransOrderDoublePlay(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  testAgv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "  testAgv3:" + sTestAgvs[2].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";
                    sTextTestData = sTextTestData + "  Transport2 Type:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sGroupID = sTestInputParams["sGroupID"].ToString();
                    sGroup2ID = sTestInputParams["sGroup2ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination:" + sDestinationID;
                    sTextTestData = sTextTestData + "  GroupA=" + sGroupID;
                    sTextTestData = sTextTestData + "  GroupB=:" + sGroup2ID;

                    //TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    m_Project.Groups[sGroup2ID].MaxMovers = 1;
                    m_Project.Groups[sGroup2ID].PreAssign = true;
                    mLogger.LogMessageToFile(sGroup2ID + " MaxMovers set to 1 and PreAssign set to true");
                    TestUtility.RemoteLogMessage(sGroup2ID + " MaxMovers set to 1 and PreAssign set to true",
                                                 sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sSourceID, sDestinationID);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.RETRIEVED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.RETRIEVED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(1)" + mMsg, sTestMonitorUsed, m_Project);
                            sTestAgvs[1].Automatic();
                            sTestAgvs[2].Automatic();
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        var transport = new Transport(Transport.COMMAND.PICK, null, null, sSource2ID, null);
                        transport.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                            sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                        if (mRunStatus == TestConstants.CHECK_END)
                        {
                            if (m_Transport.State == Transport.STATE.FINISHED)
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                mLogger.LogPassLine(mMsg);
                                TestUtility.RemoteLogPassLine("(2)" + mMsg, sTestMonitorUsed, m_Project);
                                sTestStartTime = DateTime.Now;
                                mRunStatus = TestConstants.CONTINUE;
                                mTestStatus = TestConstants.TEST_RUNS2;
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mLogger.LogFailLine(mMsg);
                                TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                                mTestStatus = TestConstants.TEST_FINISHED;
                            }
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime,
                                                            sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                        if (mRunStatus == TestConstants.CHECK_END)
                        {
                            if (m_Transport.State == Transport.STATE.FINISHED)
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                mLogger.LogPassLine(mMsg);
                                TestUtility.RemoteLogPassLine("(3)" + mMsg, sTestMonitorUsed, m_Project);
                                sTestStartTime = DateTime.Now;
                                mTestStatus = TestConstants.TEST_RUNS3;
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mLogger.LogFailLine(mMsg);
                                TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                                mTestStatus = TestConstants.TEST_FINISHED;
                            }
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        mMsg = m_Transport.ID + " has " + m_Transport.AssignedMoverID
                               + " and " + m_Transport2.ID + " has " + m_Transport2.AssignedMoverID;
                        if (m_Transport.AssignedMoverID == m_Transport2.AssignedMoverID)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    m_Project.Groups[sGroup2ID].MaxMovers = 0;
                    m_Project.Groups[sGroup2ID].PreAssign = false;
                    mLogger.LogMessageToFile(sGroup2ID + "MaxMovers set to 0 and PreAssign set to false");
                    TestUtility.RemoteLogMessage(sGroup2ID + "MaxMovers set to 0 and PreAssign set to false",
                                                 sTestMonitorUsed, m_Project);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                try
                {
                    m_Project.Groups[sGroup2ID].MaxMovers = 0;
                    m_Project.Groups[sGroup2ID].PreAssign = false;
                    mLogger.LogMessageToFile(sGroup2ID + "MaxMovers set to 0 and PreAssign set to false");
                    TestUtility.RemoteLogMessage(sGroup2ID + "MaxMovers set to 0 and PreAssign set to false",
                                                 sTestMonitorUsed, m_Project);
                    mLogger.LogTestException(ex.Message, ex.StackTrace);
                    TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                }
                finally
                {
                    EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                    ;
                }
            }
        }

        private void TestScenario_TS300064DoublePlayTransReleased(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  testAgv2:" + sTestAgvs[1].ID;
                    sTextTestData = sTextTestData + "  testAgv3:" + sTestAgvs[2].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";
                    sTextTestData = sTextTestData + "  Transport2 Type:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();
                    sGroupID = sTestInputParams["sGroupID"].ToString();
                    sGroup2ID = sTestInputParams["sGroup2ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Destination:" + sDestinationID;
                    sTextTestData = sTextTestData + "  GroupA=" + sGroupID;
                    sTextTestData = sTextTestData + "  GroupB=:" + sGroup2ID;

                    //TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    m_Project.Groups[sGroup2ID].MaxMovers = 1;
                    m_Project.Groups[sGroup2ID].PreAssign = true;
                    mLogger.LogMessageToFile(sGroup2ID + "MaxMovers set to 1 and PreAssign set to true");
                    TestUtility.RemoteLogMessage(sGroup2ID + "MaxMovers set to 1 and PreAssign set to true",
                                                 sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.PICK, null, null, sSourceID, null);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime,
                                                        sWaitTime, Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (m_Transport.State == Transport.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestAgvs[1].Automatic();
                            sTestAgvs[2].Automatic();
                            sTestAgvs[1].SimSpeed = 500;
                            sTestAgvs[2].SimSpeed = 500;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        var transport = new Transport(Transport.COMMAND.PICK, null, sTestAgvs[1].ID, null, sSource2ID,
                                                      null);
                        transport.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                        sTestStartTime = DateTime.Now;

                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // wait 10 sec
                    {
                        string assignedAgv = m_Transport2.AssignedMoverID.ToString();
                        mMsg = m_Transport2.ID + " has assigned to " + assignedAgv;
                        if (assignedAgv.Equals(sTestAgvs[1].ID.ToString()) ||
                            assignedAgv.Equals(sTestAgvs[2].ID.ToString()))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    // Wait until TransportB state ASSIGNED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.ASSIGNED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                // Create TransportC
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // wait 5 sec
                    {
                        var transport = new Transport(Transport.COMMAND.DROP, null, sTestAgvs[0].ID, null, null,
                                                      sDestinationID);
                        transport.ID = "TransportC" + sTSName;
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    // Wait until TransportC state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport3, sTestStartTime, 2,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestAgvs[1].SimSpeed = 10000;
                            sTestAgvs[2].SimSpeed = 10000;
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS5;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 60000) // wait 60 sec
                    {
                        string lsid = sTestAgvs[0].CurrentLSID.ToString();
                        mMsg = sTestAgvs[0].ID + " is at : " + lsid;
                        if (lsid.Equals(sSource2ID))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    sTestAgvs[1].SimSpeed = 10000;
                    sTestAgvs[2].SimSpeed = 10000;
                    m_Project.Groups[sGroup2ID].MaxMovers = 0;
                    m_Project.Groups[sGroup2ID].PreAssign = false;
                    mLogger.LogMessageToFile(sGroup2ID + "MaxMovers set to 0 and PreAssign set to false");
                    TestUtility.RemoteLogMessage(sGroup2ID + "MaxMovers set to 0 and PreAssign set to false",
                                                 sTestMonitorUsed, m_Project);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                try
                {
                    sTestAgvs[1].SimSpeed = 10000;
                    sTestAgvs[2].SimSpeed = 10000;
                    m_Project.Groups[sGroup2ID].MaxMovers = 0;
                    m_Project.Groups[sGroup2ID].PreAssign = false;
                    mLogger.LogMessageToFile(sGroup2ID + "MaxMovers set to 0 and PreAssign set to false");
                    TestUtility.RemoteLogMessage(sGroup2ID + "MaxMovers set to 0 and PreAssign set to false",
                                                 sTestMonitorUsed, m_Project);
                    mLogger.LogTestException(ex.Message, ex.StackTrace);
                    TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                }
                finally
                {
                    EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                    ;
                }
            }
        }

        private void TestScenario_TS830001DBSQLSERVERStopStart(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.Automatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;

                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + "MOVE";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    bool fileBufferAvtiveFlag = m_Project.RuntimeDbSerializers["SQLSERVER"].FileBufferingActiv;
                    mMsg = "fileBufferAvtiveFlag is " + fileBufferAvtiveFlag.ToString();
                    TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;

                    string machine = Environment.MachineName;
                    if (!machine.ToUpper().StartsWith("EPIATESTSERVER1"))
                    {
                        sTextTestData = "DB not tested, test machine is  " + machine;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        bool fileBufferAvtiveFlag = m_Project.RuntimeDbSerializers["SQLSERVER"].FileBufferingActiv;
                        mMsg = "fileBufferAvtiveFlag is " + fileBufferAvtiveFlag.ToString();
                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_INITED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    try // Stop SQL SERVER
                    {
                        mMsg = "Stop SQL SERVER";
                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                        var proc = new Process();
                        proc.EnableRaisingEvents = false;
                        proc.StartInfo.FileName = "net";
                        proc.StartInfo.Arguments = " Stop MSSQLSERVER";
                        proc.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                        proc.StartInfo.ErrorDialog = true;
                        proc.Start();
                        proc.WaitForExit();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + "  x:x  " + ex.StackTrace);
                    }
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 50000) // 50 seconde
                    {
                        //Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, sDestinationID, 5, false);
                        m_Transport = m_Project.TransportManager.NewTransport(transport);

                        transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, sDestinationID, 5, false);
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport);

                        transport = new Transport(Transport.COMMAND.MOVE, null, null, sTestAgvs[0].ID.ToString(), null,
                                                  null, sSourceID, sDestinationID, 5, false);
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport);

                        mMsg = "TransID---" + m_Transport.ID + Environment.NewLine
                               + "\t" + "TransID---" + m_Transport2.ID + Environment.NewLine
                               + "\t" + "TransID3---" + m_Transport3.ID + Environment.NewLine;

                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        string dbBufferFile = Path.Combine(m_Project.DbSerializingBufferDir.FullName, "DBBuffering.bin");
                        var dbFileInfo = new FileInfo(dbBufferFile);
                        long dbBufferSize = dbFileInfo.Length;
                        mMsg = "DBBuffering.bin size is:" + dbBufferSize + Environment.NewLine;
                        bool fileBufferAvtiveFlag = m_Project.RuntimeDbSerializers["SQLSERVER"].FileBufferingActiv;
                        if (fileBufferAvtiveFlag)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mMsg = mMsg + "  fileBufferAvtiveFlag is " + fileBufferAvtiveFlag.ToString();
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = mMsg + "  fileBufferAvtiveFlag is " + fileBufferAvtiveFlag.ToString();
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 8000) // 8 seconde
                    {
                        mMsg = "Start SQL SERVER";
                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                        try
                        {
                            Thread.Sleep(1000);
                            var proc = new Process();
                            proc.EnableRaisingEvents = false;
                            proc.StartInfo.FileName = "net";
                            proc.StartInfo.Arguments = " Start MSSQLSERVER";
                            proc.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                            proc.StartInfo.ErrorDialog = true;
                            proc.Start();
                            proc.WaitForExit();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
                        }

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2) // wait until flag false
                {
                    mTime = DateTime.Now - sTestStartTime;
                    bool fileBufferAvtiveFlag = m_Project.RuntimeDbSerializers["SQLSERVER"].FileBufferingActiv;
                    mMsg = "fileBufferAvtiveFlag is " + fileBufferAvtiveFlag.ToString();
                    if (!fileBufferAvtiveFlag) // 8 seconde
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        if (mTime.TotalMilliseconds >= 10000) // 10 seconde
                        {
                            mMsg = mMsg + "  max time expired: 10 sec";
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestResult = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 80000) // 80 seconde
                    {
                        string connectionString1 =
                            m_Project.RuntimeDbSerializers["SQLSERVER"].OleDbConnectionString.ToString();

                        mMsg = "connect to SQL SERVER: " + connectionString1;
                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                        string connectionString =
                            "Integrated Security=SSPI;"
                            + "Persist Security Info=False;"
                            + "Initial Catalog=Etricc_Eurobaltic;"
                            + "Data Source=EPIATESTSERVER1";

                        myConnection = new SqlConnection(connectionString);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        mMsg = "Open connection";
                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                        mTestStatus = TestConstants.TEST_RUNS5;
                        try
                        {
                            myConnection.Open();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.ToString());
                            MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        sTestStartTime = DateTime.Now;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        mTestStatus = TestConstants.TEST_RUNS5;
                        try
                        {
                            string sqlCommand
                                = "select * from Egemin_EPIA_WCS_Transportation_Transport WHERE DB_Context  like '%" +
                                  m_Transport.ID + "%'";

                            SqlDataReader myReader = null;
                            var myCommand = new SqlCommand(sqlCommand, myConnection);

                            myReader = myCommand.ExecuteReader();

                            mMsg = "sql command:" + sqlCommand + Environment.NewLine;

                            //TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                            if (myReader.HasRows)
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                mMsg = mMsg + "\t" + "has fields :" + myReader.FieldCount;
                                mMsg = mMsg + "\t" + "has record:" + myReader.HasRows.ToString();
                                mLogger.LogPassLine(mMsg);
                                TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mMsg = mMsg + "\t" + "has fields :" + myReader.FieldCount + Environment.NewLine;
                                mMsg = mMsg + "\t" + "has record:" + myReader.HasRows.ToString();
                                mLogger.LogFailLine(mMsg);
                                TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            }
                            mTestStatus = TestConstants.TEST_AFTER_RUNS;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.ToString());
                            MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        sTestStartTime = DateTime.Now;
                    }
                }

                if (mTestStatus == TestConstants.TEST_AFTER_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        try
                        {
                            mMsg = "close connection:";
                            TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                            myConnection.Close();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.ToString());
                            MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData);
                //EndTestCase(sTSName, sTestResult, ref sTextTestData); 
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                //EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData);
            }
        }

        //  TestScenario_TS300080TransPickFromGroup
        private void TestScenario_TS300080TransPickFromGroup(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sGroupID = sTestInputParams["sGroupID"].ToString();
                    sSourceID = sTestInputParams["sSourceID"].ToString();

                    sTextTestData = sTextTestData + "  GroupID:" + sGroupID;
                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    mLogger.LogMessageToFile(sTextTestData);

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;

                    sTestResult = TestConstants.TEST_UNDEFINED;
                    if (sProjectID == TestConstants.PROJECT_TESTOPSTELLING)
                        mTestStatus = TestConstants.TEST_FINISHED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Pick at 0070-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sGroupID, null, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "pick", sGroupID, null,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //3.	Wait until TransportA state WAIT FOR SOURCE
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.WAIT_SOURCE, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.	Edit TransportA source to location at 0070-01-01-01-01
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    m_Transport.SourceID = sSourceID;
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS1;
                }
                //5.	Wait until TransportA state FINISHED 
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        //  TestScenario_TS300081TransDropToGroup
        private void TestScenario_TS300081TransDropToGroup(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + sTestAgvs[0].ID;
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sGroupID = sTestInputParams["sGroupID"].ToString();
                    sDestinationID = sTestInputParams["sDestinationID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  GroupID:" + sGroupID;
                    sTextTestData = sTextTestData + "  DestinationID:" + sDestinationID;

                    mLogger.LogMessageToFile(sTextTestData);

                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;

                    if (sProjectID == TestConstants.PROJECT_TESTOPSTELLING)
                        mTestStatus = TestConstants.TEST_FINISHED;
                }
                //1.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    m_Project.Agvs.SemiAutomatic();
                    sTestAgvs[0].Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create TransportA Pick at 0070-01-01-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport = new Transport(Transport.COMMAND.PICK, null, null, sTestAgvs[0].ID.ToString(),
                                                      null, null, sSourceID, null, 5, false);
                        transport.ID = "TransportA" + sTSName;
                        m_Transport = m_Project.TransportManager.NewTransport(transport);
                        mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "pick", sSourceID, null,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }
                //3.	Wait until TransportA state FINISHED
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.	Create TransportB Pick at 0070-06-02 by AGV11
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // wait 3 sec
                    {
                        var transport2 = new Transport(Transport.COMMAND.DROP, null, null, sTestAgvs[0].ID.ToString(),
                                                       null, null, null, sGroupID, 5, false);
                        transport2.ID = "TransportB" + sTSName;
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport2);
                        mLogger.LogCreatedTransport(m_Transport2.ID.ToString(), "DROP", null, sGroupID,
                                                    sTestAgvs[0].ID.ToString(), 5);
                        sTestStartTime = DateTime.Now;
                        mRunStatus = TestConstants.CONTINUE;
                        mTestStatus = TestConstants.TEST_RUNS1;
                    }
                }
                //5.	Wait until TransportB state WAIT FOR DESTINATION
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.WAIT_DESTINATION, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //6.	Edit TransportA destination to location at 0070-06-02-01-01
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    m_Transport2.DestinationID = sDestinationID;
                    mMsg = "Edit TransportA destination to location at " + sDestinationID;
                    mLogger.LogPassLine(mMsg);
                    TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS3;
                }
                //7.	Wait until TransportB state FINISHED 
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, sWaitTime,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS200080AgvModeSemiToAuto(string sTSName)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "   Agv:" + sTestAgvs[0].ID;
                    sTextTestData = sTextTestData + "   Job Type:" + "PICK";

                    sTestInputParams = TestData.GetTestInputParams(m_Project, sTSName);
                    sSourceID = sTestInputParams["sSourceID"].ToString();
                    sSource2ID = sTestInputParams["sSource2ID"].ToString();
                    sSource3ID = sTestInputParams["sSource3ID"].ToString();

                    sTextTestData = sTextTestData + "  SourceID:" + sSourceID;
                    sTextTestData = sTextTestData + "  Source2ID:" + sSource2ID;
                    sTextTestData = sTextTestData + "  Source3ID:" + sSource3ID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }
                //1.	Set Agv11 mode semi-automatic
                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestAgvs[0].SemiAutomatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }
                //2.	Create JobA Pick at 0070-01-01-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    Job job = TestUtility.CreateTestJob("JobA-" + sTSName, "PICK", sSourceID, sTSName + " jobA",
                                                        sProjectID);
                    m_Job = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job);
                    mLogger.LogCreatedJob("JobA-" + sTSName, "PICK", sSourceID, sTestAgvs[0].ID.ToString());
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_STARTING;
                }
                //3.	Wait until JobA state BUSY and Agv11 not at park location
                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    TestUtility.WaitUntilJobState(ref mRunStatus, m_Job, Job.STATE.BUSY, sTestStartTime, sWaitTime,
                                                  ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //4.	Create JobB Pick at 0070-02-01-01-01 by AGV11
                //5.	Create JobC Pick at 0070-03-01-01-01 by AGV11
                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    Job job2 = TestUtility.CreateTestJob("JobB-" + sTSName, "PICK", sSource2ID, sTSName + " jobB",
                                                         sProjectID);
                    m_Job2 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job2);
                    mLogger.LogCreatedJob("JobB-" + sTSName, "PICK", sSource2ID, sTestAgvs[0].ID.ToString());

                    Job job3 = TestUtility.CreateTestJob("JobC-" + sTSName, "PICK", sSource3ID, sTSName + " jobC",
                                                         sProjectID);
                    m_Job3 = m_Project.Agvs[sTestAgvs[0].ID.ToString()].NewJob(job3);
                    mLogger.LogCreatedJob("JobC-" + sTSName, "PICK", sSource3ID, sTestAgvs[0].ID.ToString());

                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS1;
                }
                //6.	Check JobB and JobC state PENDING
                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 10000) // 10 seconde
                    {
                        mMsg = m_Job2.ID + "  has state : " + m_Job2.State.ToString();
                        mMsg += " and " + m_Job3.ID + "  has state : " + m_Job3.State.ToString();
                        if (m_Job2.State == Job.STATE.PENDING && m_Job3.State == Job.STATE.PENDING)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //7.	Set Agv11 mode automatic
                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    mMsg = sTestAgvs[0].ID + "  has mode automatic ";
                    sTestAgvs[0].Automatic();
                    mLogger.LogPassLine(mMsg);
                    TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_RUNS3;
                }
                //8.	Wait 5 sec
                //9.	Check JobB and JobC state FINISHED
                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        mMsg = m_Job2.ID + "  has state : " + m_Job2.State.ToString();
                        mMsg += " and " + m_Job3.ID + "  has state : " + m_Job3.State.ToString();
                        if (m_Job2.State == Job.STATE.FINISHED && m_Job3.State == Job.STATE.FINISHED)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_RUNS4;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //10.	Wait until Agv11 has Park JOB
                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    if (TestUtility.CheckAgvHasJobWithType(sTestAgvs[0], "PARK"))
                    {
                        mMsg = " testAgv has  park job : ";
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS5;
                    }
                    else
                    {
                        mTime = DateTime.Now - sTestStartTime;
                        if (mTime.TotalMilliseconds >= sWaitTime*60*1000)
                        {
                            mMsg = "after " + sWaitTime + "  min still nopark job : ";
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }
                //11.	Wait until Agv11 return ParkLocation
                //ToDo should change to any Park location 
                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    if (TestUtility.IsAgvAtParkLocation(sTestAgvs[0], m_Project, ref mMsg))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                    else
                    {
                        mTime = DateTime.Now - sTestStartTime;
                        if (mTime.TotalMilliseconds >= sWaitTime*60*1000)
                        {
                            mMsg = "after " + sWaitTime + "min, " + mMsg + " but still at : " + sTestAgvs[0].CurrentLSID;
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    //EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                //EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData); ;
                EndTestCaseAndUpdateFiles(sTSName, sTestResult, ref sTextTestData);
            }
        }

        //========================================================================================================================

        private void TestScenario_TS305506ScheduleBattRulesQueueWP(string sTSName, Agv testAgv, Agv testAgv2,
                                                                   Agv testAgv3)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + testAgv.ID;
                    sTextTestData = sTextTestData + "  testAgv2:" + testAgv2.ID;
                    sTextTestData = sTextTestData + "  testAgv3:" + testAgv3.ID;

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    m_Project.Agvs.Automatic();
                    //TestUtility.SetSchedulesRules(ref m_Project, "*.BATT", "QUEUE", "PARK_BAT 2");

                    string args = "PARK_BAT";
                    int plusMin = 1;
                    WeekPlan wp
                        = TestUtility.CreateWeekPlan(sTSName, args, 0 /*today+0)*/, plusMin);

                    testAgv.BatteryChargePlan.Deactivate();
                    testAgv.BatteryChargePlan.Clear();
                    testAgv.BatteryChargePlan.Add(wp, true);
                    testAgv.BatteryChargePlan.Activate();

                    testAgv2.BatteryChargePlan.Deactivate();
                    testAgv2.BatteryChargePlan.Clear();
                    testAgv2.BatteryChargePlan.Add(wp, true);
                    testAgv2.BatteryChargePlan.Activate();

                    testAgv3.BatteryChargePlan.Deactivate();
                    testAgv3.BatteryChargePlan.Clear();
                    testAgv3.BatteryChargePlan.Add(wp, true);
                    testAgv3.BatteryChargePlan.Activate();

                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 150000) // wait 150 sec
                    {
                        int batts = 0;
                        if (TestUtility.CheckAgvHasJobWithType(testAgv, "Batt"))
                            batts++;
                        if (TestUtility.CheckAgvHasJobWithType(testAgv2, "Batt"))
                            batts++;
                        if (TestUtility.CheckAgvHasJobWithType(testAgv3, "Batt"))
                            batts++;

                        mMsg = "total batt jobs is " + batts;
                        sTextTestData = sTextTestData + " ( total batt jobs: " + batts + ")";

                        if (batts < 2)
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogFailLine(mMsg);
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mLogger.LogPassLine(mMsg);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    //TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                }
            }
            catch (Exception ex)
            {
                //TestUtility.SetSchedulesRules(ref m_Project, "*.PICK", "CLOSEST_HIGHEST", null);
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS242099WeekPlanCalibrationMultipleTrigger(string sTSName, Agv[] agvs)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    for (int i = 0; i < agvs.Length; i++)
                    {
                        sTextTestData = sTextTestData + "testAgv" + i + " :  " + agvs[i].ID;
                    }
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    //int numAgvs = m_Project.Agvs.GetArray().Length;
                    int numAgvs = agvs.Length;
                    //Object[] agvs = m_Project.Agvs.GetArray();

                    int plusMin = 1;
                    string args = "CX01";
                    WeekPlan wp
                        = TestUtility.CreateWeekPlan(sTSName, args, 0 /*today+0*/, plusMin);

                    for (int i = 0; i < numAgvs; i++)
                    {
                        agvs[i].CalibrationPlan.Deactivate();
                        agvs[i].CalibrationPlan.Clear();
                        agvs[i].CalibrationPlan.Add(wp, true);
                        agvs[i].CalibrationPlan.Activate();
                    }
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    for (int i = 0; i < agvs.Length; i++)
                    {
                        agvs[i].Restart();
                    }
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 180000) // 180 seconde
                    {
                        int waits = 0;
                        int numAgvs = agvs.Length;
                        for (int i = 0; i < numAgvs; i++)
                        {
                            if (TestUtility.CheckAgvHasJobWithType(agvs[i], "Wait"))
                                waits++;
                        }
                        sTextTestData = sTextTestData + "  ( total Wait job :" + waits + " )";


                        mMsg = "total wait job created " + waits;
                        if (waits == numAgvs)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS830002DBBufferingNotEmptyAtStartup(string sTSName, Agv testAgv, string transType,
                                                                       string sourceID, string destID)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    m_Project.Agvs.Automatic();
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    sTextTestData = sTextTestData + "  SourceID:" + sourceID;
                    sTextTestData = sTextTestData + "  Transport Type:" + transType;
                    sTextTestData = sTextTestData + "  DestinationID:" + destID;

                    bool fileBufferAvtiveFlag = m_Project.RuntimeDbSerializers["SQLSERVER"].FileBufferingActiv;
                    mMsg = "fileBufferAvtiveFlag is " + fileBufferAvtiveFlag.ToString();

                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;

                    string machine = Environment.MachineName;
                    if (!machine.ToUpper().StartsWith("EPIATESTSERVER1"))
                    {
                        sTextTestData = "DB not tested, test machine is  " + machine;
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        string dbBufferFile = Path.Combine(m_Project.DbSerializingBufferDir.FullName, "DBBuffering.bin");
                        var dbFileInfo = new FileInfo(dbBufferFile);
                        long dbBufferSize = dbFileInfo.Length;
                        mMsg = "DBBuffering.bin size is:" + dbBufferSize + Environment.NewLine;
                        bool fileBufferAvtiveFlag = m_Project.RuntimeDbSerializers["SQLSERVER"].FileBufferingActiv;
                        if (fileBufferAvtiveFlag)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mMsg = mMsg + "  fileBufferAvtiveFlag is " + fileBufferAvtiveFlag.ToString();
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        }
                        mTestStatus = TestConstants.TEST_INITED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    try // Stop SQL SERVER
                    {
                        mMsg = "Stop SQL SERVER";
                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                        var proc = new Process();
                        proc.EnableRaisingEvents = false;
                        proc.StartInfo.FileName = "net";
                        proc.StartInfo.Arguments = " Stop MSSQLSERVER";
                        proc.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                        proc.StartInfo.ErrorDialog = true;
                        proc.Start();
                        proc.WaitForExit();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + "  x:x  " + ex.StackTrace);
                    }
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_STARTING;
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 50000) // 50 seconde
                    {
                        //Move at 0070-01-01-01-01 to 0360-01-01 by AGV11 with priority 5.
                        var transport = new Transport(Transport.COMMAND.MOVE, null, null, testAgv.ID.ToString(), null,
                                                      null, sourceID, destID, 5, false);
                        m_Transport = m_Project.TransportManager.NewTransport(transport);

                        transport = new Transport(Transport.COMMAND.MOVE, null, null, testAgv.ID.ToString(), null, null,
                                                  sourceID, destID, 5, false);
                        m_Transport2 = m_Project.TransportManager.NewTransport(transport);

                        transport = new Transport(Transport.COMMAND.MOVE, null, null, testAgv.ID.ToString(), null, null,
                                                  sourceID, destID, 5, false);
                        m_Transport3 = m_Project.TransportManager.NewTransport(transport);

                        mMsg = "TransID---" + m_Transport.ID + Environment.NewLine
                               + "\t" + "TransID---" + m_Transport2.ID + Environment.NewLine
                               + "\t" + "TransID3---" + m_Transport3.ID + Environment.NewLine;

                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        string dbBufferFile = Path.Combine(m_Project.DbSerializingBufferDir.FullName, "DBBuffering.bin");
                        var dbFileInfo = new FileInfo(dbBufferFile);
                        long dbBufferSize = dbFileInfo.Length;
                        mMsg = "DBBuffering.bin size is:" + dbBufferSize + Environment.NewLine;
                        bool fileBufferAvtiveFlag = m_Project.RuntimeDbSerializers["SQLSERVER"].FileBufferingActiv;
                        if (fileBufferAvtiveFlag)
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mMsg = mMsg + "  fileBufferAvtiveFlag is " + fileBufferAvtiveFlag.ToString();
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS1;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mMsg = mMsg + "  fileBufferAvtiveFlag is " + fileBufferAvtiveFlag.ToString();
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS1)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 8000) // 8 seconde
                    {
                        mMsg = "Start SQL SERVER";
                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                        try
                        {
                            Thread.Sleep(1000);
                            var proc = new Process();
                            proc.EnableRaisingEvents = false;
                            proc.StartInfo.FileName = "net";
                            proc.StartInfo.Arguments = " Start MSSQLSERVER";
                            proc.StartInfo.WorkingDirectory = Directory.GetCurrentDirectory();
                            proc.StartInfo.ErrorDialog = true;
                            proc.Start();
                            proc.WaitForExit();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
                        }

                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS2;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2) // wait until flag false
                {
                    mTime = DateTime.Now - sTestStartTime;
                    bool fileBufferAvtiveFlag = m_Project.RuntimeDbSerializers["SQLSERVER"].FileBufferingActiv;
                    mMsg = "fileBufferAvtiveFlag is " + fileBufferAvtiveFlag.ToString();
                    if (!fileBufferAvtiveFlag) // 8 seconde
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS3;
                    }
                    else
                    {
                        if (mTime.TotalMilliseconds >= 10000) // 10 seconde
                        {
                            mMsg = mMsg + "  max time expired: 10 sec";
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            sTestResult = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 80000) // 80 seconde
                    {
                        string connectionString1 =
                            m_Project.RuntimeDbSerializers["SQLSERVER"].OleDbConnectionString.ToString();

                        mMsg = "connect to SQL SERVER: " + connectionString1;
                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                        string connectionString =
                            "Integrated Security=SSPI;"
                            + "Persist Security Info=False;"
                            + "Initial Catalog=Etricc_Eurobaltic;"
                            + "Data Source=EPIATESTSERVER1";

                        myConnection = new SqlConnection(connectionString);
                        sTestStartTime = DateTime.Now;
                        mTestStatus = TestConstants.TEST_RUNS4;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS4)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        mMsg = "Open connection";
                        TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                        mTestStatus = TestConstants.TEST_RUNS5;
                        try
                        {
                            myConnection.Open();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.ToString());
                            MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        sTestStartTime = DateTime.Now;
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS5)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        mTestStatus = TestConstants.TEST_RUNS5;
                        try
                        {
                            string sqlCommand
                                = "select * from Egemin_EPIA_WCS_Transportation_Transport WHERE DB_Context  like '%" +
                                  m_Transport.ID + "%'";

                            SqlDataReader myReader = null;
                            var myCommand = new SqlCommand(sqlCommand, myConnection);

                            myReader = myCommand.ExecuteReader();

                            mMsg = "sql command:" + sqlCommand + Environment.NewLine;

                            //TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);

                            if (myReader.HasRows)
                            {
                                sTestResult = TestConstants.TEST_PASS;
                                mMsg = mMsg + "\t" + "has fields :" + myReader.FieldCount;
                                mMsg = mMsg + "\t" + "has record:" + myReader.HasRows.ToString();
                                mLogger.LogPassLine(mMsg);
                                TestUtility.RemoteLogPassLine(mMsg, sTestMonitorUsed, m_Project);
                            }
                            else
                            {
                                sTestResult = TestConstants.TEST_FAIL;
                                mMsg = mMsg + "\t" + "has fields :" + myReader.FieldCount + Environment.NewLine;
                                mMsg = mMsg + "\t" + "has record:" + myReader.HasRows.ToString();
                                mLogger.LogFailLine(mMsg);
                                TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                            }
                            mTestStatus = TestConstants.TEST_AFTER_RUNS;
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.ToString());
                            MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                        sTestStartTime = DateTime.Now;
                    }
                }

                if (mTestStatus == TestConstants.TEST_AFTER_RUNS)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 3000) // 3 seconde
                    {
                        try
                        {
                            mMsg = "close connection:";
                            TestUtility.RemoteLogMessage(mMsg, sTestMonitorUsed, m_Project);
                            myConnection.Close();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.ToString());
                            MessageBox.Show(e.Message + Environment.NewLine + e.StackTrace);
                        }
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS24XX01WeekPlanActiveCheck(string sTSName, Agv testAgv, string typeWP,
                                                              ref int testStatus, string args)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                {
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);
                }

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + testAgv.ID;
                    sTextTestData = sTextTestData + "  args:" + args;
                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    testAgv.Automatic();
                    sTestStartTime = DateTime.Now;
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 5000) // 5 seconde
                    {
                        int plusMin = 1;
                        WeekPlan wp
                            = TestUtility.CreateWeekPlan(sTSName, args, 0 /*today+0)*/, plusMin);
                        if (typeWP.ToLower().StartsWith("wpbatt"))
                        {
                            testAgv.BatteryChargePlan.Deactivate();
                            testAgv.BatteryChargePlan.Clear();
                            testAgv.BatteryChargePlan.Add(wp, true);
                            testAgv.BatteryChargePlan.Activate();
                        }
                        else if (typeWP.ToLower().StartsWith("wpcalib"))
                        {
                            testAgv.CalibrationPlan.Deactivate();
                            testAgv.CalibrationPlan.Clear();
                            testAgv.CalibrationPlan.Add(wp, true);
                            testAgv.CalibrationPlan.Activate();
                        }
                        sTestStartTime = DateTime.Now;
                        TestUtility.RemoteLogMessage(typeWP + " added", sTestMonitorUsed, m_Project);
                        mTestStatus = TestConstants.TEST_STARTING;
                    }
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    mTime = DateTime.Now - sTestStartTime;
                    if (mTime.TotalMilliseconds >= 15000) // 15 seconde
                    {
                        bool activeWP = testAgv.IsAvailable(Job.TYPE.UNDEFINED, null, true);
                        string msg = "wp is active " + activeWP;
                        TestUtility.RemoteLogMessage(msg, sTestMonitorUsed, m_Project);
                        mLogger.LogMessageToFile(msg);

                        if (activeWP)
                            sTestResult = TestConstants.TEST_PASS;
                        else
                            sTestResult = TestConstants.TEST_FAIL;

                        sTextTestData = sTextTestData + "  WP active:" + activeWP.ToString();
                        mTestStatus = TestConstants.TEST_FINISHED;
                    }
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                {
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                    ;
                }
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        private void TestScenario_TS300055TransOrderPriority(string sTSName, Agv testAgv, string transType,
                                                             string sourceID, string destID, string source2ID,
                                                             string dest2ID)
        {
            try
            {
                if (mTestStatus == TestConstants.TEST_NOT_STARTED)
                    CleanUP(sTSName, sTestAgvs, sAgvsInitialID, sAgvsDefaultDropID);

                if (mTestStatus == TestConstants.TEST_CLEARNED)
                {
                    sTextTestData = sTextTestData + "testAgv:" + testAgv.ID;
                    sTextTestData = sTextTestData + "  Transport Type:" + transType;
                    sTextTestData = sTextTestData + "  SourceID:" + sourceID;
                    sTextTestData = sTextTestData + "  DestinationID:" + destID;
                    sTextTestData = sTextTestData + "  Source2ID:" + source2ID;
                    sTextTestData = sTextTestData + "  Destination2ID:" + dest2ID;
                    mLogger.LogMessageToFile(sTextTestData);
                    TestUtility.RemoteLogMessage(sTextTestData, sTestMonitorUsed, m_Project);

                    mTestStatus = TestConstants.TEST_INPUT_TEXT_ADDED;
                }

                if (mTestStatus == TestConstants.TEST_INPUT_TEXT_ADDED)
                {
                    sTestResult = TestConstants.TEST_UNDEFINED;
                    mMsg = string.Empty;
                    m_Project.Agvs.SemiAutomatic();
                    testAgv.Automatic();
                    testAgv.Suspend();
                    mTestStatus = TestConstants.TEST_INITED;
                }

                if (mTestStatus == TestConstants.TEST_INITED)
                {
                    //Move at 0070-01-01 to 0360-01-01 by AGV11 with priority 8.
                    var transport = new Transport(Transport.COMMAND.MOVE, null, null, testAgv.ID.ToString(), null, null,
                                                  sourceID, destID, 8, false);
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    // Move at 0040-01-01 to 0360-01-01 by AGV11 with priority 1.
                    transport = new Transport(Transport.COMMAND.MOVE, null, null, testAgv.ID.ToString(), null, null,
                                              source2ID, dest2ID, 1, false);
                    m_Transport2 = m_Project.TransportManager.NewTransport(transport);
                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", sourceID, destID,
                                                testAgv.ID.ToString(), 1);
                    mLogger.LogCreatedTransport(m_Transport.ID.ToString(), "MOVE", source2ID, dest2ID,
                                                testAgv.ID.ToString(), 8);

                    mTestStatus = TestConstants.TEST_STARTING;
                }

                if (mTestStatus == TestConstants.TEST_STARTING)
                {
                    testAgv.Release();
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mTestStatus = TestConstants.TEST_RUNS;
                }

                if (mTestStatus == TestConstants.TEST_RUNS)
                {
                    // Wait until Transport state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport, sTestStartTime, 2,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mRunStatus = TestConstants.CONTINUE;
                            mTestStatus = TestConstants.TEST_RUNS2;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS2)
                {
                    // (3) Wait until TransportA state FINISHED
                    TestUtility.WaitUntilTransportState(ref mRunStatus, m_Transport2, sTestStartTime, 2,
                                                        Transport.STATE.FINISHED, ref mMsg);
                    if (mRunStatus == TestConstants.CHECK_END)
                    {
                        if (mMsg.StartsWith("OK"))
                        {
                            sTestResult = TestConstants.TEST_PASS;
                            mLogger.LogPassLine(mMsg);
                            TestUtility.RemoteLogPassLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            sTestStartTime = DateTime.Now;
                            mTestStatus = TestConstants.TEST_RUNS3;
                        }
                        else
                        {
                            sTestResult = TestConstants.TEST_FAIL;
                            mLogger.LogFailLine(mMsg);
                            TestUtility.RemoteLogFailLine("(3) " + mMsg, sTestMonitorUsed, m_Project);
                            mTestStatus = TestConstants.TEST_FINISHED;
                        }
                    }
                }

                if (mTestStatus == TestConstants.TEST_RUNS3)
                {
                    DateTime time1 = m_Transport.Finished;
                    DateTime time2 = m_Transport2.Finished;
                    mTime = time2 - time1;
                    if (mTime.TotalMilliseconds > 0)
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(" time1: " + time1 + " and time2: " + time2);
                        TestUtility.RemoteLogPassLine("time1: " + time1 + " and time2: " + time2, sTestMonitorUsed,
                                                      m_Project);
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_FAIL;
                        mLogger.LogFailLine("time1: " + time1 + " and time2: " + time2);
                        TestUtility.RemoteLogFailLine("time1: " + time1 + " and time2: " + time2, sTestMonitorUsed,
                                                      m_Project);
                    }
                    mTestStatus = TestConstants.TEST_FINISHED;
                }

                if (mTestStatus == TestConstants.TEST_FINISHED)
                    EndTestCase(sTSName, sTestResult, ref sTextTestData);
                ;
            }
            catch (Exception ex)
            {
                mLogger.LogTestException(ex.Message, ex.StackTrace);
                TestUtility.RemoteLogTestException(ex.Message, sTestMonitorUsed, m_Project);
                EndTestCase(sTSName, TestConstants.TEST_EXCEPTION, ref sTextTestData);
                ;
            }
        }

        #endregion //------ END Test Scenarios ------------------------------------------

        # region CleanUP

        protected void CleanUP(string testScenario, Agv[] testAgvs, Hashtable agvsInitialID, Hashtable agvsDefDropID)
        {
            if (mCleanUPStatus == TestConstants.CLEANUP_NOT_STARTED)
            {
                TestUtility.RemoteLogTestReset(testScenario + "==" + m_Project.ID, sTestMonitorUsed, m_Project);
                mCleanUPStatus = TestConstants.CLEANUP_INIT_INPUT;
            }

            if (mCleanUPStatus == TestConstants.CLEANUP_INIT_INPUT)
            {
                m_Project.Agvs.Removed();
                sTestStartTime = DateTime.Now;
                mCleanUPStatus = TestConstants.CLEANUP_STARTING;
            }


            // Set All Agvs semi-automatic
            if (mCleanUPStatus == TestConstants.CLEANUP_STARTING)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 5000) // 5 second
                {
                    mMsg = "(2) Set All Agvs mode semi-automatic";
                    TestUtility.RemoteLogTestReset(mMsg, sTestMonitorUsed, m_Project);
                    for (int i = 0; i < m_Project.Agvs.GetArray().Length; i++)
                    {
                        var agv = (Agv) m_Project.Agvs.GetArray()[i];
                        agv.SemiAutomatic();
                    }
                    mLogger.LogMessageToFile(mMsg);
                    sTestStartTime = DateTime.Now;
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS;
                }
            }

            // restart All Agvs
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 8000) // 8 second
                {
                    mMsg = "(3) Restart All Agvs";
                    TestUtility.RemoteLogTestReset(mMsg, sTestMonitorUsed, m_Project);
                    m_Project.Agvs.Restart();

                    mLogger.LogMessageToFile(mMsg);
                    mRunStatus = TestConstants.CONTINUE;
                    sTestStartTime = DateTime.Now;
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS1;
                }
            }
            // Wait until All Agvs state READY Or READYCHARGING
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS1)
            {
                TestUtility.WaitUntilAgvsStateReadyOrReadyCharging(ref mRunStatus, sTestAgvs, sTestStartTime, sWaitTime,
                                                                   ref mMsg);
                if (mRunStatus == TestConstants.CHECK_END)
                {
                    if (mMsg.StartsWith("OK"))
                    {
                        sTestStartTime = DateTime.Now;
                        mCleanUPStatus = TestConstants.CLEANUP_RUNS2;
                    }
                    else
                        throw new Exception("Clean Up Failed during Agvs restart:" + mMsg);
                }
            }
            // Cancel All Unfinished Transports
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS2)
            {
                Transports transports = m_Project.Transports;
                for (int i = 0; i < transports.Count; i++)
                {
                    var tr = (Transport) transports.GetArray()[i];
                    if (tr.State < Transport.STATE.FINISHED)
                        tr.Cancel();
                }
                sTestStartTime = DateTime.Now;
                mCleanUPStatus = TestConstants.CLEANUP_RUNS3;
            }

            // Flush All Transports
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS3)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 2000) // 2 second
                {
                    mMsg = "(4) Flush All transports";
                    TestUtility.RemoteLogTestReset(mMsg, sTestMonitorUsed, m_Project);
                    m_Project.Transports.FlushAll();

                    mLogger.LogMessageToFile(mMsg);
                    sTestStartTime = DateTime.Now;
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS4;
                }
            }

            // Cleanup Jobs
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS4)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 1000) // 1 second
                {
                    TestUtility.RemoteLogTestReset("(5)Cleanup All jobs", sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile("Cleanup All jobs");
                    Agvs agvs = m_Project.Agvs;
                    for (int i = 0; i < agvs.Count; i++)
                    {
                        var agv = (Agv) m_Project.Agvs.GetArray()[i];
                        CleanUpAgvAllJobs(agv);
                    }
                    mLogger.LogMessageToFile("All Agv jobs Cleaned up");
                    sTestStartTime = DateTime.Now;
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS6;
                }
            }

            // 1 first return park place
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS6)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 2000) // 2 second
                {
                    mLogger.LogMessageToFile(" First All agvs return to parking");
                    // Check Park place and return to parking
                    Agvs agvs = m_Project.Agvs;
                    for (int i = 0; i < agvs.Count; i++)
                    {
                        var agv = (Agv) m_Project.Agvs.GetArray()[i];
                        mLogger.LogMessageToFile(agv.ID + " current at " + agv.CurrentLSID
                                                 + " should park to  " + agvsInitialID[agv.ID.ToString()]
                                                 + " it has mode  " + agv.Mode.ToString()
                                                 + " it has state: " + agv.State.ToString());
                        CheckAgvParkLocationAndSendToPark(agv, agvsInitialID[agv.ID.ToString()].ToString());
                    }
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS8;
                }
            }


            // wait until Agvs at Initial positions
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS8)
            {
                TestUtility.WaitUntilAgvsAtInitialPositions(ref mRunStatus, m_Project, sTestStartTime, 3, agvsInitialID,
                                                            ref mMsg);
                if (mRunStatus == TestConstants.CHECK_END)
                {
                    if (mMsg.StartsWith("OK"))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogTestReset("(6)" + mMsg, sTestMonitorUsed, m_Project);
                        mCleanUPStatus = TestConstants.CLEANUP_RUNS9;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_EXCEPTION;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mCleanUPStatus = TestConstants.CLEANUP_FINISHED;
                    }
                }
            }

            // Drop Loads
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS9)
            {
                mLogger.LogMessageToFile("CLEANUP_RUNS9:drop load");
                Agvs agvs = m_Project.Agvs;
                for (int i = 0; i < agvs.Count; i++)
                {
                    var agv = (Agv) m_Project.Agvs.GetArray()[i];
                    if (agv.Loaded)
                        DropLoadWhenLoaded(agv, agvsDefDropID[agv.ID.ToString()].ToString());
                }
                sTestStartTime = DateTime.Now;
                mRunStatus = TestConstants.CONTINUE;
                TestUtility.RemoteLogTestReset("(7)CLEANUP_RUNS9:drop loads", sTestMonitorUsed, m_Project);
                mCleanUPStatus = TestConstants.CLEANUP_RUNS10;
            }


            // Wait Until All Loads Droped
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS10)
            {
                TestUtility.WaitUntilAgvsEmpty(ref mRunStatus, m_Project, sTestStartTime, sWaitTime, ref mMsg);
                if (mRunStatus == TestConstants.CHECK_END)
                {
                    if (mMsg.StartsWith("OK"))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogTestReset("(8)" + mMsg, sTestMonitorUsed, m_Project);
                        sTestStartTime = DateTime.Now;
                        mCleanUPStatus = TestConstants.CLEANUP_RUNS12;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_EXCEPTION;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mCleanUPStatus = TestConstants.CLEANUP_FINISHED;
                    }
                }
            }

            // first return park place again
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS12)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 2000) // 2 second
                {
                    // Check Park place and return to parking
                    Agvs agvs = m_Project.Agvs;
                    for (int i = 0; i < agvs.Count; i++)
                    {
                        var agv = (Agv) m_Project.Agvs.GetArray()[i];
                        CheckAgvParkLocationAndSendToPark(agv, agvsInitialID[agv.ID.ToString()].ToString());
                    }
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS13;
                }
            }


            // wait until Agvs at Initial positions
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS13)
            {
                TestUtility.WaitUntilAgvsAtInitialPositions(ref mRunStatus, m_Project, sTestStartTime, 3, agvsInitialID,
                                                            ref mMsg);
                if (mRunStatus == TestConstants.CHECK_END)
                {
                    if (mMsg.StartsWith("OK"))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        sTestStartTime = DateTime.Now;
                        TestUtility.RemoteLogTestReset("(9)" + mMsg, sTestMonitorUsed, m_Project);
                        mCleanUPStatus = TestConstants.CLEANUP_RUNS14;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_EXCEPTION;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mCleanUPStatus = TestConstants.CLEANUP_FINISHED;
                    }
                    sTestStartTime = DateTime.Now;
                }
            }


            // Discard All loads
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS14)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 2000) // 2 second
                {
                    Object[] loads = m_Project.Loads.GetArray();
                    for (int i = 0; i < loads.Length; i++)
                    {
                        var load = (Load) loads[i];
                        load.Discard();
                    }
                    sTestStartTime = DateTime.Now;
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS15;
                }
            }

            // 5 Flush all loads
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS15)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 2000) // 2 second
                {
                    mLogger.LogMessageToFile("Flush all loads");
                    m_Project.Loads.FlushAll();
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS16;
                    sTestStartTime = DateTime.Now;
                }
            }

            // Cleanup Jobs
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS16)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 1000) // 1 second
                {
                    TestUtility.RemoteLogTestReset("(10)Cleanup All jobs", sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile("Cleanup All jobs");
                    Agvs agvs = m_Project.Agvs;
                    for (int i = 0; i < agvs.Count; i++)
                    {
                        var agv = (Agv) m_Project.Agvs.GetArray()[i];
                        CleanUpAgvAllJobs(agv);
                    }
                    mLogger.LogMessageToFile("All Agv jobs Cleaned up");
                    mLogger.LogTestRunStartup(testScenario);

                    m_Project.Agvs.Removed();
                    sTestStartTime = DateTime.Now;
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS17;
                }
            }

            // Set All Agvs mode semi-automatic
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS17)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 1000) // 1 second
                {
                    TestUtility.RemoteLogTestReset("(11)CLEANUP_RUNS17", sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile("CLEANUP_RUNS17");

                    for (int i = 0; i < m_Project.Agvs.GetArray().Length; i++)
                    {
                        var agv = (Agv) m_Project.Agvs.GetArray()[i];
                        agv.SemiAutomatic();
                    }


                    sTestStartTime = DateTime.Now;
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS18;
                }
            }

            // Restart TestAgvs 
            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS18)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 5000) // 5 second
                {
                    TestUtility.RemoteLogTestReset("(12)CLEANUP_RUNS18: restart testAgvs", sTestMonitorUsed, m_Project);
                    mLogger.LogMessageToFile("CLEANUP_RUNS18 restart testAgvs");
                    for (int i = 0; i < testAgvs.Length; i++)
                    {
                        testAgvs[i].Restart();
                    }
                    mLogger.LogTestRunStartup(testScenario);
                    sTestStartTime = DateTime.Now;
                    mRunStatus = TestConstants.CONTINUE;
                    mCleanUPStatus = TestConstants.CLEANUP_RUNS19;
                }
            }

            if (mCleanUPStatus == TestConstants.CLEANUP_RUNS19)
            {
                TestUtility.WaitUntilAgvsStateReadyOrReadyCharging(ref mRunStatus, testAgvs, sTestStartTime, sWaitTime,
                                                                   ref mMsg);
                if (mRunStatus == TestConstants.CHECK_END)
                {
                    if (mMsg.StartsWith("OK"))
                    {
                        sTestResult = TestConstants.TEST_PASS;
                        mLogger.LogPassLine(mMsg);
                        TestUtility.RemoteLogTestReset("(13)" + mMsg, sTestMonitorUsed, m_Project);
                        mCleanUPStatus = TestConstants.CLEANUP_FINISHED;
                    }
                    else
                    {
                        sTestResult = TestConstants.TEST_EXCEPTION;
                        mLogger.LogFailLine(mMsg);
                        TestUtility.RemoteLogFailLine(mMsg, sTestMonitorUsed, m_Project);
                        mCleanUPStatus = TestConstants.CLEANUP_FINISHED;
                    }
                    sTestStartTime = DateTime.Now;
                }
            }


            if (mCleanUPStatus == TestConstants.CLEANUP_FINISHED)
            {
                sTestStartTime = DateTime.Now;
                //mJobStatus = TestConstants.JOB_NOT_STARTED;
                mCleanUPStatus = TestConstants.CLEANUP_NOT_STARTED;
                mTestStatus = TestConstants.TEST_CLEARNED;
            }
        }

        protected void CleanUpAgvAllJobs(Agv agv)
        {
            for (int i = 0; i < agv.Jobs.GetArray().Length; i++)
            {
                var job = (Job) agv.Jobs.GetArray()[i];
                if (job.State < Job.STATE.FINISHED)
                {
                    TestUtility.RemoteLogPassLine("Cancel :" + job.ID, sTestMonitorUsed, m_Project);
                    job.Cancel();
                }
            }
            agv.Jobs.FlushAll();
        }

        protected void CheckAgvParkLocationAndSendToPark(Agv agv, string parkID)
        {
            Job job;
            if (!agv.CurrentLSID.ToString().Equals(parkID))
            {
                job = TestUtility.CreateTestJob("PARK", parkID, agv.ID + " CleanUP ");
                m_Project.Agvs[agv.ID.ToString()].NewJob(job);
            }
        }

        protected void DropLoadWhenLoaded(Agv agv, string destID)
        {
            if (agv.Loaded)
            {
                Job job = TestUtility.CreateTestJob("DROP", destID, agv.ID + " CleanUP: drop load");
                m_Project.Agvs[agv.ID.ToString()].NewJob(job);
            }
        }

        # endregion//--------- CleanUP --------------

        #region //------ Test Help ------------------------------------------

        /// <summary>
        /// CreateAndWaitUntilJobFinished
        /// </summary>
        /*protected void CreateAndWaitUntilJobFinished(ref int status, Agv agv, string jobType,
				string LocationID, string TestScenarioID, ref int result, ref string msg)
		{
			Job job;
			if (status == TestConstants.JOB_NOT_STARTED)
			{
				job = TestUtility.CreateTestJob(jobType, LocationID, TestScenarioID);
				Job thejob = m_Project.Agvs[agv.ID.ToString()].NewJob(job);
				sTestStartTime = DateTime.Now;
				sJobStartTime = DateTime.Now;
				mLogger.LogCreatedJob(thejob.ID.ToString(), jobType, LocationID, agv.ID.ToString());
				status = TestConstants.JOB_CREATED;
			}

			if (status == TestConstants.JOB_CREATED)							// wait until job finished
			{
				mTime = DateTime.Now - sTestStartTime;
				if (mTime.TotalMilliseconds >= 3000)							// wait 3 sec
				{
					string lsid = agv.CurrentLSID.ToString();
					if (lsid.Equals(LocationID))
					{
						result = TestConstants.TEST_PASS;
						msg = string.Empty;
						status = TestConstants.JOB_FINISHED;
					}
					else
					{
						mTime = DateTime.Now - sJobStartTime;
						if (mTime.TotalMilliseconds >= 100000)					// wait 100 sec
						{
							result = TestConstants.TEST_FAIL;
							msg = agv.ID.ToString() + " at : " + lsid;
							status = TestConstants.JOB_FINISHED;
						}
					}
					sTestStartTime = DateTime.Now;
				}
			}
		}
		// end CreateAndWaitUntilJobFinished
		*/
        /*
		protected void CreateAndWaitUntilTransportFinished(ref int status, Agv agv, string transType, string sourceID,
			string destID, string TestScenarioID, ref int result, ref string msg)
		{
			if (status == TestConstants.JOB_NOT_STARTED)
			{
				if (agv != null)
					agv.Automatic();

				m_Transport = TestUtility.CreateTestTransport(transType, agv, sourceID, destID, TestScenarioID, ref m_Project, ref m_Transport);
				sTestStartTime = DateTime.Now;
				sJobStartTime = DateTime.Now;
				if (agv != null)
					mLogger.LogCreatedTransport(m_Transport.ID.ToString(), transType, sourceID, destID, agv.ID.ToString());
				status = TestConstants.JOB_CREATED;
			}

			if (status == TestConstants.JOB_CREATED)						// wait until transport finished 
			{
				mTime = DateTime.Now - sTestStartTime;
				if (mTime.TotalMilliseconds >= 3000)						// wait 3 second
				{
					string state = m_Transport.State.ToString();
					if (m_Transport.State >= Transport.STATE.FINISHED)
					{
						msg = "transport finished";
						result = TestConstants.TEST_PASS;
						status = TestConstants.JOB_FINISHED;
					}
					else
					{
						mTime = DateTime.Now - sJobStartTime;
						if (mTime.TotalMilliseconds >= 100000)  // 100 second
						{
							msg = "transport not finished state:" + state;
							result = TestConstants.TEST_FAIL;
							status = TestConstants.JOB_FINISHED;
						}
					}
					sTestStartTime = DateTime.Now;
				}
			}
		}
		*/
        protected void EndTestCase(string sTSName, int result, ref string testData)
        {
            for (int i = 0; i < sTestAgvs.Length; i++)
            {
                if (sTestAgvs[i] != null)
                {
                    if (sTestAgvs[i].Mode == Mover.MODE.REMOVED)
                        mLogger.LogMessageToFile("\t\t------    " + sTestAgvs[i].ID + "\t mode is removed.\t ");
                    else
                        mLogger.LogMessageToFile("\t\t------    " + sTestAgvs[i].ID + "\t pos.\t " +
                                                 sTestAgvs[i].CurrentLSID + " \t\t- Loaded : " +
                                                 sTestAgvs[i].Loaded.ToString());
                }
                else
                    mLogger.LogMessageToFile("\t\t------   testAgv" + (i + 1) + " is null");
            }


            if (sTestID == 0)
            {
                mTestStatus = TestConstants.TEST_NOT_STARTED;

                mLogger.LogMessageToFile("\tEnd Case(" + sCounter + ")-----  Test " + sTSName + " :   end :next one ");
                mLogger.LogMessageToFile(" ");
                TestUtility.RemoteLogMessage(
                    "\tEnd Case(" + sCounter + ")--- Test " + sTSName + " :   end :next one\n\n ", sTestMonitorUsed,
                    m_Project);

                sCounter++;
                sTotalTestCounter++;

                try
                {
                    if (result == TestConstants.TEST_FAIL)
                        sTotalFailCounter++;
                    else if (result == TestConstants.TEST_PASS)
                        sTotalPassCounter++;
                    else if (result == TestConstants.TEST_EXCEPTION)
                        sTotalExceptionCounter++;
                    else if (result == TestConstants.TEST_UNDEFINED)
                        sTotalUntestedCounter++;

                    //WriteExcelWorkSheetTestResultOfThisCase( result, sCounter, sTSName, testData );
                    string time = DateTime.Now.ToString("HH:mm:ss");
                    int row = sCounter;
                    //if (PCName.ToUpper().StartsWith(PC_TEAMTESTETRICC5))
                    //{
                    xlsBody[row, 0] = time;
                    xlsBody[row, 1] = sTSName;
                    xlsBody[row, 2] = "" + result;
                    xlsBody[row, 3] = testData;
                    //}
                    /*else
					{
						xSheet.Cells.set_Item(row, 1, time);
						xSheet.Cells.set_Item(row, 2, sTSName);
						xRange = xSheet.get_Range("B" + row, "B" + row);
						switch (result)
						{
							case TestConstants.TEST_PASS:
								xRange.Interior.ColorIndex = TestConstants.EXCEL_GREEN;
								break;
							case TestConstants.TEST_FAIL:
								xRange.Interior.ColorIndex = TestConstants.EXCEL_RED;
								xSheet.Cells.set_Item(row, 3, testData);
								break;
							case TestConstants.TEST_EXCEPTION:
								xRange.Interior.ColorIndex = TestConstants.EXCEL_PINK;
								break;
							case TestConstants.TEST_UNDEFINED:
								xRange.Interior.ColorIndex = TestConstants.EXCEL_YELLOW;
								xSheet.Cells.set_Item(row, 3, testData);
								break;
						}
					}*/
                }
                catch (Exception ex)
                {
                    mLogger.LogTestException("ExcelException:" + ex.Message, ex.StackTrace);
                }
                testData = string.Empty;
            }
            else
            {
                mLogger.LogMessageToFile("\t\t------    Test " + sTSName + " :   end (no next test)");
                mLogger.LogMessageToFile(" ");
                TestUtility.RemoteLogMessage("\t------    Test " + sTSName + " :   end (no next test)\n\n",
                                             sTestMonitorUsed, m_Project);
                mTestStatus = 99;
            }
            mCleanUPStatus = TestConstants.CLEANUP_NOT_STARTED;
            Thread.Sleep(2000);
        }

        protected void EndTestCaseAndUpdateFiles(string sTSName, int result, ref string testData)
        {
            if (mJobStatus == TestConstants.JOB_NOT_STARTED)
            {
                for (int i = 0; i < sTestAgvs.Length; i++)
                {
                    if (sTestAgvs[i] != null)
                    {
                        if (sTestAgvs[i].Mode == Mover.MODE.REMOVED)
                            mLogger.LogMessageToFile("\t\t------    " + sTestAgvs[i].ID + "\t mode is removed.\t ");
                        else
                            mLogger.LogMessageToFile("\t\t------    " + sTestAgvs[i].ID + "\t pos.\t " +
                                                     sTestAgvs[i].CurrentLSID + " \t\t- Loaded : " +
                                                     sTestAgvs[i].Loaded.ToString());
                    }
                    else
                        mLogger.LogMessageToFile("\t\t------   testAgv" + (i + 1) + " is null");
                }

                mLogger.LogMessageToFile("\t\t------    Test " + sTSName + " :   end :last one ");
                mLogger.LogMessageToFile(" End of the Test");
                TestUtility.RemoteLogMessage("\t------    Test " + sTSName + " :   end :last one\n\n ", sTestMonitorUsed,
                                             m_Project);

                sTestStartTime = DateTime.Now;
                mJobStatus = TestConstants.JOB_CREATED;
            }

            string xNameLog = string.Empty;
            string excelpath = string.Empty;
            string root = m_Project.Facilities["Tests"].Parameters["TestCenterRoot"].ValueAsString;
            if (mJobStatus == TestConstants.JOB_CREATED)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 3000) // 3 second
                {
                    TestUtility.RemoteLogMessage("\t------    Test Finished1", sTestMonitorUsed, m_Project);
                    if (sTestID == 0)
                    {
                        sTotalTestCounter++;
                        try
                        {
                            // last test case in excel sheet 
                            if (result == TestConstants.TEST_FAIL)
                                sTotalFailCounter++;
                            else if (result == TestConstants.TEST_PASS)
                                sTotalPassCounter++;
                            else if (result == TestConstants.TEST_EXCEPTION)
                                sTotalExceptionCounter++;
                            else if (result == TestConstants.TEST_UNDEFINED)
                                sTotalUntestedCounter++;

                            //WriteExcelWorkSheetTestResultOfThisCase(  result, sCounter, sTSName, testData );
                            string time = DateTime.Now.ToString("HH:mm:ss");
                            int row = sCounter + 1;
                            string xNameXLS = string.Empty;

                            //if (PCName.ToUpper().StartsWith(PC_TEAMTESTETRICC5))
                            //{
                            mLogger.LogMessageToFile("WWWWWWWWWWWWWWWWW row :" + row);
                            xlsBody[row, 0] = time;
                            xlsBody[row, 1] = sTSName;
                            xlsBody[row, 2] = "" + result;
                            xlsBody[row, 3] = testData;
                            for (int i = 0; i < row + 1; i++)
                            {
                                mLogger.LogMessageToFile(" row[" + i + ",0] :" + xlsBody[i, 0]);
                                mLogger.LogMessageToFile(" row[" + i + ",1] :" + xlsBody[i, 1]);
                                mLogger.LogMessageToFile(" row[" + i + ",2] :" + xlsBody[i, 2]);
                                mLogger.LogMessageToFile(" row[" + i + ",3] :" + xlsBody[i, 3]);
                            }

                            var yApp = new Microsoft.Office.Interop.Excel.Application();
                            Workbooks yBooks = yApp.Workbooks;
                            Workbook yBook = yBooks.Add(Type.Missing);
                            var ySheet = (Worksheet) yBook.Worksheets[1];
                            Range yRange;
                            yApp.Visible = testinfo.excelShow_11;
                            yApp.Interactive = true;

                            // Header
                            string today = DateTime.Now.ToString("MMMM-dd");
                            ySheet.Cells.set_Item(1, 1, today);
                            ySheet.Cells.set_Item(1, 2, "Test Scenarios");
                            for (int i = 1; i < 9; i++)
                            {
                                ySheet.Cells.set_Item(i, 3, xlsHeader[i - 1]);
                                Thread.Sleep(1000);
                            }

                            // Body
                            for (int i = 2; i < row + 1; i++)
                            {
                                ySheet.Cells.set_Item(i, 1, xlsBody[i, 0]);
                                ySheet.Cells.set_Item(i, 2, xlsBody[i, 1]);
                                yRange = ySheet.get_Range("B" + i, "B" + i);
                                switch (xlsBody[i, 2])
                                {
                                    case "1001": //TestConstants.TEST_PASS:
                                        yRange.Interior.ColorIndex = TestConstants.EXCEL_GREEN;
                                        break;
                                    case "1002": //TestConstants.TEST_FAIL:
                                        yRange.Interior.ColorIndex = TestConstants.EXCEL_RED;
                                        ySheet.Cells.set_Item(row, 3, xlsBody[i, 2]);
                                        break;
                                    case "1003": //TestConstants.TEST_EXCEPTION:
                                        yRange.Interior.ColorIndex = TestConstants.EXCEL_PINK;
                                        break;
                                    case "1000": //TestConstants.TEST_UNDEFINED:
                                        yRange.Interior.ColorIndex = TestConstants.EXCEL_YELLOW;
                                        ySheet.Cells.set_Item(row, 3, xlsBody[i, 2]);
                                        break;
                                }
                                Thread.Sleep(1000);
                            }
                            // Foot

                            row = sCounter + 1;

                            TestUtility.AddTestTotalCounterToExcel(ref ySheet, row + 1,
                                                                   sTotalTestCounter, sTotalPassCounter,
                                                                   sTotalFailCounter);
                            TestUtility.AddLengendeToExcel(ref ySheet, row + 7);

                            string xPath = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + TEST_RESULT_FILE_SURFIX
                                           + "-" + Environment.MachineName;
                            xNameXLS = xPath + ".xls";
                            xNameLog = xPath + ".log";

                            sXLSPath = Path.Combine(root, xPath + ".xls");

                            // Save the Workbook and quit Excel.
                            mLogger.LogMessageToFile("save excel 1:sXLSPath :" + sXLSPath);
                            object missing = Missing.Value;
                            yBook.SaveAs(sXLSPath, XlFileFormat.xlWorkbookNormal,
                                         missing, missing, missing, missing,
                                         XlSaveAsAccessMode.xlNoChange,
                                         missing, missing, missing, missing, missing);
                            //1, true, missing, missing, missing); 

                            if (yBook != null) yBook.Close(true, xPath, false);
                            if (yBooks != null) yBooks.Close();
                            yApp.Quit();
                            //}
                            /*else
							{
								xSheet.Cells[row, 1] = time;
								xSheet.Cells[row, 2] = sTSName;
								//xSheet.Cells.set_Item(row, 1, time);
								//xSheet.Cells.set_Item(row, 2, sTSName);
								xRange = xSheet.get_Range("B" + row, "B" + row);
								switch (result)
								{
									case TestConstants.TEST_PASS:
										xRange.Interior.ColorIndex = TestConstants.EXCEL_GREEN;
										break;
									case TestConstants.TEST_FAIL:
										xRange.Interior.ColorIndex = TestConstants.EXCEL_RED;
										xSheet.Cells.set_Item(row, 3, testData);
										break;
									case TestConstants.TEST_EXCEPTION:
										xRange.Interior.ColorIndex = TestConstants.EXCEL_PINK;
										break;
									case TestConstants.TEST_UNDEFINED:
										xRange.Interior.ColorIndex = TestConstants.EXCEL_YELLOW;
										xSheet.Cells.set_Item(row, 3, testData);
										break;
								}

								row = sCounter + 1;

								TestUtility.AddTestTotalCounterToExcel(ref xSheet, row + 1,
									sTotalTestCounter, sTotalPassCounter, sTotalFailCounter);
								TestUtility.AddLengendeToExcel(ref xSheet, ref xRange, row + 7);

								string xPath = System.DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + TEST_RESULT_FILE_SURFIX
									+ "-" + System.Environment.MachineName;
								xNameXLS = xPath + ".xls";
								xNameLog = xPath + ".log";

								sXLSPath = System.IO.Path.Combine(root, xPath + ".xls");

								// Save the Workbook and quit Excel.
								object missing = System.Reflection.Missing.Value;
								xBook.SaveAs(sXLSPath, Excel.XlFileFormat.xlWorkbookNormal,
												missing, missing, missing, missing,
												Excel.XlSaveAsAccessMode.xlNoChange,
												missing, missing, missing, missing, missing);
								//1, true, missing, missing, missing); 

								if (xBook != null) xBook.Close(true, xPath, false);
								if (xBooks != null) xBooks.Close();
								xApp.Quit();
							}*/

                            //Save xPath to G driver if machine is TestPC
                            if (testinfo.installedApp_1.Equals("Installed:Etricc5"))
                            {
                                mLogger.LogMessageToFile("copy excel ---> installed App is:" + testinfo.installedApp_1);

                                #region  // save excel and log to Xdrive only for "Etricc5"

                                //if (System.Environment.MachineName.StartsWith("EPIATESTPC"))
                                //{
                                string dir = testinfo.buildInstallScriptDir_7;
                                string path = Path.GetFullPath(dir);

                                //System.Windows.MessageBox.Show("path 1:" + path);

                                int ib = path.LastIndexOf("\\");
                                while (ib > 0)
                                {
                                    string y = path.Substring(ib + 1);
                                    //System.Windows.MessageBox.Show("y :" + y);
                                    if (y.StartsWith("Etricc") && y.IndexOf("_") > 0 && y.IndexOf(".") > 0)
                                    {
                                        //System.Windows.MessageBox.Show("path 2:" + path);
                                        break;
                                    }
                                    else
                                    {
                                        path = path.Substring(0, ib);
                                        //System.Windows.MessageBox.Show("xxxxx path 2:" + path);
                                        ib = path.LastIndexOf("\\");
                                    }
                                }

                                excelpath = path + "\\TestResults";
                                mLogger.LogMessageToFile("save excel to x:Driver :" + excelpath);

                                if (Directory.Exists(excelpath))
                                {
                                    mLogger.LogMessageToFile(" X:Driver directory exist :" + excelpath);
                                }
                                else
                                {
                                    mLogger.LogMessageToFile(" X:Driver directory NOT exist :" + excelpath);

                                    int ret = OpenDriveMap(@"\\Teamsystem\Team Systems Builds", "X:");
                                    if (ret == 0 || ret == 85)
                                    {
                                        mLogger.LogMessageToFile("OPEN MAP DRIVE OK:");
                                        if (Directory.Exists(excelpath))
                                        {
                                            mLogger.LogMessageToFile(" X:Driver directory exist :" + excelpath);
                                        }
                                        else
                                        {
                                            mLogger.LogMessageToFile(" X:Driver directory NOT  exist AGAIN :" +
                                                                     excelpath);
                                        }
                                    }
                                    else
                                    {
                                        mLogger.LogMessageToFile("OpenDriveMap failed with error code:" + ret);
                                        ret = CreateDriveMap(excelpath);
                                        if (ret != 0)
                                        {
                                            mLogger.LogMessageToFile("CreateDriveMap failed with error code:" + ret);
                                            //return false;
                                        }
                                    }
                                }

                                CopyResult(root, excelpath, xNameXLS);
                                mLogger.LogMessageToFile("test results saved on Marvel_Technologies_server ");

                                mLogger.LogMessageToFile("Copy test log file ");
                                CopyLogFile("C:\\EtriccTests", excelpath, TEST_CENTER_LOG_FILE, xNameLog);
                                //}

                                #endregion
                            }
                            else
                                mLogger.LogMessageToFile("installed App is:" + testinfo.installedApp_1);
                        }
                        catch (Exception ex)
                        {
                            mLogger.LogTestException("xxxxxxxxx:" + ex.Message, ex.StackTrace);
                            //mJobStatus = TestConstants.JOB_FINISHED2;
                            sCounter = -1;
                        }
                    }
                    sTestStartTime = DateTime.Now;
                    mJobStatus = TestConstants.JOB_FINISHED;
                }
            }

            if (mJobStatus == TestConstants.JOB_FINISHED)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 3000) // 3 second
                {
                    if (sTestID == 0)
                    {
                        string str1 = "<html><body><b><center>Test Overview</center></b><br><br><table col=" +
                                      '"' + "5" + '"' + " > <tr><th></th><th>Total Tests:</th><th>&nbsp;</th><th>"
                                      + sTotalTestCounter + "<" + testinfo.buildInstallScriptDir_7 + ">"
                                      + "</th>  <th></th> </tr>"
                                      + "<tr><td><br></td><td></td><td></td><td></td><td></td></tr>"
                                      +
                                      "                            <tr><td></td>	<td>Pass:     </td> <td></td>       <td>"
                                      + sTotalPassCounter
                                      +
                                      "</td><td></td><td></td></tr><tr><td></td>	<td>Fail:     </td>  <td></td>	    <td>"
                                      + sTotalFailCounter
                                      +
                                      "</td><td></td><td></td></tr><tr><td></td>	<td>Exception:</td>  <td></td>	    <td>"
                                      + sTotalExceptionCounter
                                      +
                                      "</td><td></td><td></td></tr><tr><td></td>	<td>Untested:</td>   <td></td>	    <td>"
                                      + sTotalUntestedCounter
                                      + "</td><td></td>	<td></td></tr></table><br><br></body></html>";

                        string TextStatistics = "       Test Overview   " + Environment.NewLine;
                        TextStatistics = TextStatistics + "Total Test Cases:     " + sTotalTestCounter +
                                         Environment.NewLine;
                        TextStatistics = TextStatistics + "Pass:                 " + sTotalPassCounter +
                                         Environment.NewLine;
                        TextStatistics = TextStatistics + "Fail:                 " + sTotalFailCounter +
                                         Environment.NewLine;
                        TextStatistics = TextStatistics + "Exception:            " + sTotalExceptionCounter +
                                         Environment.NewLine;
                        TextStatistics = TextStatistics + "Untested:             " + sTotalUntestedCounter +
                                         Environment.NewLine;
                        TextStatistics = TextStatistics + Environment.NewLine;

                        TestUtility.RemoteLogMessage(TextStatistics, sTestMonitorUsed, m_Project);

                        TextStatistics = str1;
                        // send email
                        TestUtility.SendTestResultToDevelopers(sXLSPath, sPrjLayout[sProjectID.ToString()].ToString(),
                                                               testinfo.buildType_3, ref mLogger, sTotalFailCounter,
                                                               TextStatistics, string.Empty, testinfo.sendMail_10);


                        // check cmd proc. if exist close it --> clearn up screen
                        //TestUtility.CloseProcesses("cmd");
                        // check excel proc. if exist close it --> clearn up screen
                        //TestUtility.KillProcesses("EXCEL");
                    }
                    sTestStartTime = DateTime.Now;
                    mJobStatus = TestConstants.JOB_FINISHED1;
                }
            }

            if (mJobStatus == TestConstants.JOB_FINISHED1)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 3000) // 3 second
                {
                    if (sTestID == 0)
                    {
                        if (PCName.ToUpper().StartsWith(PC_TEAMTESTETRICC5))
                            Thread.Sleep(1000);
                        else
                        {
                            TestUtility.RemoteLogMessage("\t------    Test Finished3", sTestMonitorUsed, m_Project);
                            /*while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xRange) > 0)
								xRange = null;
							while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xSheet) > 0)
								xSheet = null;
							while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xBook) > 0)
								xBook = null;
							while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xBooks) > 0)
								xBooks = null;
							while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xApp) > 0)
								xApp = null;

							xRange = null;
							xSheet = null;
							xBook = null;
							xBooks = null;
							xApp = null;*/
                        }
                        GC.GetTotalMemory(false);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        GC.GetTotalMemory(true);

                        mLogger.LogMessageToFile("test results saved ");

                        Process[] pExcel = Process.GetProcessesByName("EXCEL");
                        try
                        {
                            for (int i = 0; i < pExcel.Length; i++)
                                pExcel[i].Kill();
                        }
                        catch
                        {
                        }
                    }
                    sTestStartTime = DateTime.Now;
                    mJobStatus = TestConstants.JOB_FINISHED2;
                    //sCounter = -1;
                }
            }

            if (mJobStatus == TestConstants.JOB_FINISHED2)
            {
                mTime = DateTime.Now - sTestStartTime;
                if (mTime.TotalMilliseconds >= 3000) // 3 second
                {
                    // temp root set to string "C:\\EtriccTests"
                    string setupinfo = TestUtility.UpdateSetupInfoFile(ref mLogger, "C:\\EtriccTests" /*sRoot*/,
                                                                       testinfo.installedApp_1, sTotalFailCounter);
                    //mLogger.LogMessageToFile("Copy test log file ");
                    //CopyLogFile("C:\\EtriccTests", excelpath, TEST_CENTER_LOG_FILE, xNameLog);
                    sCounter = -1;
                }
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
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

        /// <summary>
        /// Copies the new setup to the local machine
        /// </summary>
        /// <returns></returns>
        public bool CopyResult(string fromPath, string toPath, string fileName)
        {
            TestUtility.RemoteLogMessage("from path :" + fromPath, sTestMonitorUsed, m_Project);
            TestUtility.RemoteLogMessage("to path :" + toPath, sTestMonitorUsed, m_Project);
            TestUtility.RemoteLogMessage("file name :" + fileName, sTestMonitorUsed, m_Project);

            if (toPath.StartsWith(@"\\"))
            {
                //if the first action fails try to logon to the server
                if (CreateDriveMap(toPath) != 0)
                {
                    return false;
                }
            }

            if (!Directory.Exists(fromPath))
            {
                Directory.CreateDirectory(fromPath);
            }

            if (Directory.Exists(toPath))
            {
                TestUtility.RemoteLogMessage("to Path exist :OK", sTestMonitorUsed, m_Project);
            }
            else
                Directory.CreateDirectory(toPath);

            TestUtility.RemoteLogMessage("topath checked :" + fileName, sTestMonitorUsed, m_Project);

            try
            {
                string xFile = Path.Combine(fromPath, fileName);
                TestUtility.RemoteLogMessage("copyed file :" + xFile, sTestMonitorUsed, m_Project);

                var file = new FileInfo(xFile);
                file.CopyTo(Path.Combine(toPath, file.Name));
                TestUtility.RemoteLogMessage("Copied result from " + fromPath + " to " + toPath, sTestMonitorUsed,
                                             m_Project);
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException("FromPath=" + fromPath + "   " + ex + "\r\n" + ex.StackTrace,
                                                   sTestMonitorUsed, m_Project);
                return false;
            }
            return true;
        }

        public bool CopyLogFile(string srcPath, string destPath, string srcFileName, string destFileName)
        {
            if (srcPath.StartsWith(@"\\"))
            {
                //if the first action fails try to logon to the server
                if (CreateDriveMap(destPath) != 0)
                {
                    return false;
                }
            }

            if (!Directory.Exists(srcPath))
            {
                Directory.CreateDirectory(srcPath);
            }

            if (Directory.Exists(destPath))
            {
                TestUtility.RemoteLogMessage("dest Path exist :OK", sTestMonitorUsed, m_Project);
            }
            else
                Directory.CreateDirectory(destPath);

            TestUtility.RemoteLogMessage("destpath checked :" + srcFileName, sTestMonitorUsed, m_Project);

            try
            {
                string xFile = Path.Combine(srcPath, srcFileName);
                TestUtility.RemoteLogMessage("copyed file :" + xFile, sTestMonitorUsed, m_Project);

                mLogger.LogMessageToFile("copyed file :" + xFile);
                mLogger.LogMessageToFile("dest file :" + Path.Combine(destPath, destFileName));

                var file = new FileInfo(xFile);
                file.CopyTo(Path.Combine(destPath, destFileName));
                TestUtility.RemoteLogMessage("Copied result from " + srcPath + " to " + destPath, sTestMonitorUsed,
                                             m_Project);
            }
            catch (Exception ex)
            {
                TestUtility.RemoteLogTestException("FromPath=" + srcPath + "   " + ex + "\r\n" + ex.StackTrace,
                                                   sTestMonitorUsed, m_Project);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Create a drive mapping to the destination
        /// </summary>
        /// <param name="Destination">Full drive path</param>
        public static int CreateDriveMap(string Destination)
        {
            if ((Destination == null) || (Destination == ""))
                return -1;

            var netResource = new NETRESOURCEA[1];
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
        /// Create a drive mapping to the destination
        /// </summary>
        /// <param name="Destination">Full drive path</param>
        private int OpenDriveMap(string Destination, string driveLetter)
        {
            if ((Destination == null) || (Destination == ""))
                return -1;

            var netResource = new NETRESOURCEA[1];
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

        public static int Disconnect(string localpath)
        {
            int result = WNetCancelConnection2A(localpath, 1, 1);
            return result;
        }

        #endregion//------ Test Help ------------------------------------------

        #region properties

        public StringCollection Logging
        {
            get { return m_Logging; }
            //			set
            //			{
            //				m_Logging = value;
            //			}
        }

        #endregion properties
    }
}