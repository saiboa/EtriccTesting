using System;

namespace TestRuns
{
    /// <summary>
    /// Summary description for TestConstants.
    /// </summary>
    public class TestConstants
    {
        #region Enums/Constants

        public const string TEST_WORKER_STATUS_FILE = "TestWorker.txt";
        public const string TEST_LOG_PATH = @"C:\EtriccTests";

        //PROJECT ID
        public const int PROJECT_EUROBALTIC = 3000;
        public const int PROJECT_TESTOPSTELLING = 3001;
        public const int PROJECT_DEMO = 3002;

        // Test Process Phase 
        public const int TEST_NOT_STARTED = -1;
        public const int TEST_CLEARNED = 1;
        public const int TEST_INPUT_TEXT_ADDED = 2;
        public const int TEST_INITED = 3;
        public const int TEST_STARTING = 4;
        public const int TEST_RUNS = 5;
        public const int TEST_RUNS1 = 6;
        public const int TEST_RUNS2 = 7;
        public const int TEST_RUNS3 = 8;
        public const int TEST_RUNS4 = 9;
        public const int TEST_RUNS5 = 10;
        public const int TEST_RUNS6 = 11;
        public const int TEST_RUNS7 = 12;
        public const int TEST_RUNS8 = 13;
        public const int TEST_RUNS9 = 14;
        public const int TEST_AFTER_RUNS = 15;
        public const int TEST_STOP = 16;
        public const int TEST_AFTER = 17;
        public const int TEST_FINISHED = 18;
        public const int TEST_FINISHED1 = 19;
        public const int TEST_FINISHED2 = 20;
        public const int TEST_FINISHED3 = 21;
        public const int TEST_FINISHED4 = 22;
        public const int TEST_FINISHED5 = 23;


        public const int CONTINUE = 200;
        public const int CHECK_END = 201;

        // reset phases
        public const int CLEANUP_NOT_STARTED = 39;
        public const int CLEANUP_INIT_INPUT = 40;
        public const int CLEANUP_STARTING = 41;
        public const int CLEANUP_RUNS = 42;
        public const int CLEANUP_RUNS1 = 43;
        public const int CLEANUP_RUNS2 = 44;
        public const int CLEANUP_RUNS3 = 45;
        public const int CLEANUP_RUNS4 = 46;
        public const int CLEANUP_RUNS5 = 47;
        public const int CLEANUP_RUNS6 = 48;
        public const int CLEANUP_RUNS7 = 49;
        public const int CLEANUP_RUNS8 = 50;
        public const int CLEANUP_RUNS9 = 51;
        public const int CLEANUP_RUNS10 = 52;
        public const int CLEANUP_RUNS11 = 53;
        public const int CLEANUP_RUNS12 = 54;
        public const int CLEANUP_RUNS13 = 55;
        public const int CLEANUP_RUNS14 = 56;
        public const int CLEANUP_RUNS15 = 57;
        public const int CLEANUP_RUNS16 = 58;
        public const int CLEANUP_RUNS17 = 59;
        public const int CLEANUP_RUNS18 = 60;
        public const int CLEANUP_RUNS19 = 61;
        public const int CLEANUP_RUNS20 = 62;
        public const int CLEANUP_RUNS21 = 63;
        public const int CLEANUP_RUNS22 = 64;
        public const int CLEANUP_AFTER_RUNS = 65;
        public const int CLEANUP_STOP = 66;
        public const int CLEANUP_PASS = 67;
        public const int CLEANUP_FAIL = 68;
        public const int CLEANUP_CANCEL_JOBS = 69;
        public const int CLEANUP_FLUSH_LOADS = 70;
        public const int CLEANUP_AFTER = 71;
        public const int CLEANUP_FINISHED = 72;

        // job or transport
        public const int JOB_NOT_STARTED = 130;
        public const int JOB_CREATED = 131;
        public const int JOB_FINISHED = 132;
        public const int JOB_FINISHED1 = 133;
        public const int JOB_FINISHED2 = 134;
        // Excel color
        public const int EXCEL_BLACK = 1;
        public const int EXCEL_WHITE = 2;
        public const int EXCEL_RED = 3;
        public const int EXCEL_GREEN = 4;
        public const int EXCEL_BLUE = 5;
        public const int EXCEL_YELLOW = 6;
        public const int EXCEL_PINK = 7;
        public const int EXCEL_LIGHTBLUE = 8;
        public const int EXCEL_BRUNE = 9;

        // test result
        public const int TEST_UNDEFINED = 1000;
        public const int TEST_PASS = 1001;
        public const int TEST_FAIL = 1002;
        public const int TEST_EXCEPTION = 1003;

        public struct TESTINFO
        {
            public bool autoTestMode_15; //15
            public string buildApplication_4; //  4 Etricc 5
            public string buildInstallScriptDir_7; //  7
            public string buildType_3; //  3
            public string demo_9; // 9 
            public string epiaDeployPath_2; //  2
            public bool excelShow_11; //  11
            public string installedApp_1; //  1
            public string oS_12; //  12
            public string projectFile_6; // 6
            public string sendMail_10; //  10
            public string testAppWorkingDir_14; // 14
            public string testDirectory_13; //13
            public string testToolsVersion_5; //  5

            public override String ToString()
            {
                String str = "installedApp: " + installedApp_1 + Environment.NewLine
                             + "EtriccDeployPath: " + epiaDeployPath_2 + Environment.NewLine
                             + "Build Type: " + buildType_3 + Environment.NewLine
                             + "Build Application: " + buildApplication_4 + Environment.NewLine
                             + "Test Tools Version: " + testToolsVersion_5 + Environment.NewLine
                             + "Project Filename: " + projectFile_6 + Environment.NewLine
                             + "Build install File: " + buildInstallScriptDir_7 + Environment.NewLine
                             + "Demo test ?: " + demo_9 + Environment.NewLine
                             + "Send email after test?: " + sendMail_10 + Environment.NewLine
                             + "Excel Visible? " + excelShow_11 + Environment.NewLine
                             + "OS: " + oS_12 + Environment.NewLine
                             + "Directory used by test: " + testDirectory_13 + Environment.NewLine
                             + "TestTool Working Dir: " + testAppWorkingDir_14 + Environment.NewLine
                             + "Automation test?: " + autoTestMode_15 + Environment.NewLine;
                return (str);
            }
        }

        #endregion // —— Enums/Constants ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        #region Structs/Classes

        #endregion // —— Structs/Classes ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        #region Fields

        #endregion // —— Fields ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••?

        #region Constructors/Destructors/Cleanup

        #endregion // —— Constructors/Destructors/Cleanup ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••?

        #region Properties

        #endregion // —— Properties ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••?

        #region Methods

        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        // ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————

        // ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————

        // ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————

        // ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————

        // ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————

        // ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————

        // ——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
    }

    // class
}

// namespace