namespace TestTools
{
    public class ConstCommon
    {
        public const string SMTP_SERVERID = "outlook.egemin.be";
        public const string DRIVE_MAP_LETTER = "X:";

        // application names
        public const string EPIA = "Epia";
        public const string ETRICC = "Etricc";
        public const string ETRICC_UI = "Etricc UI";
        public const string ETRICCUI = "EtriccUI";
        public const string ETRICC_5 = "Etricc 5";
        public const string ETRICC_ETRICC5 = "Etricc+Etricc5";
        public const string KIMBERLY_CLARK = "Kimberly Clark";
        public const string ETRICCSTATISTICS = "EtriccStatistics";
        public const string KC = "KC";
        // Project names
        public const string EPIA_3 = "Epia 3";
        public const string EPIA_4 = "Epia 4";
        //public const string ETRICC_5 = "Etricc 5";    The project name is = Application name
        public const string EWCS_PROJECTS = "Ewcs - Projects";

        // test result
        public const int TEST_UNDEFINED = 1000;
        public const int TEST_PASS = 1001;
        public const int TEST_FAIL = 1002;
        public const int TEST_EXCEPTION = 1003;
        // 

        public const string PARSERCONFIGURATOR_ROOT = "\\Egemin\\Etricc Statistics ParserConfigurator";
        public const string PARSERCONFIGURATOR_EXE = "Egemin.Etricc.Statistics.ParserConfigurator.exe";
        //
        public const string EGEMIN_EPIA_SERVER = "Egemin.Epia.Server";
        public const string EGEMIN_ETRICC_SERVER = "Egemin.Etricc.Server";
        public const string EGEMIN_ETRICC_STATISTICS_PARSER = "Egemin.Etricc.Statistics.Parser";
        public const string EGEMIN_ETRICC_STATISTICS_PARSERCONFIGURATOR = "Egemin.Etricc.Statistics.ParserConfigurator";
        public const string EGEMIN_EWCS_SERVER = "Egemin.Ewcs.Server";
        public const string EGEMIN_EWCS_TOOLS_DATABASE_FILLER = "Egemin.Ewcs.Tools.DatabaseFiller";
        public const string EGEMIN_EPIA_SHELL = "Egemin.Epia.Shell";

        public const string EGEMIN_EPIA_SERVER_EXE = "Egemin.Epia.Server.Exe";
        public const string EGEMIN_ETRICC_SERVER_EXE = "Egemin.Etricc.Server.Exe";
        public const string EGEMIN_EPIA_SHELL_EXE = "Egemin.Epia.Shell.Exe";
        public const string EGEMIN_EWCS_SERVER_EXE = "Egemin.Ewcs.Server.Exe";
        public const string EGEMIN_ETRICC_EXPLORER_EXE = "EPIA.Explorer.Exe";

        // EPIA UI
        public const string MY_LAYOUT = "My layout";
        public const string MY_PLACE = "My Place";
        public const string MY_SETTINGS = "My settings";

        public const string DEFAULT_WINDOW_NAAM = "Egemin Shell";

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

        public const string ETRICC_TESTS_DIRECTORY = "C:\\EtriccTests";

        // Test Data Files
        public const string EUROBALTIC_PROJECT_NAME = "Eurobaltic";
        public const string EUROBALTIC_PROJECT_ZIP = "Eurobaltic.zip";
        public const string EUROBALTIC_PROJECT_XML = "Eurobaltic.xml";


        public const string TESTINFO_FILENAME = "TestInfo.txt";
        public const string TESTWORKING_FILENAME = "TestWorking.txt";
        public static string EPIA_SERVER_ROOT = OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Dematic\\Epia Server";
        public static string EPIA_CLIENT_ROOT = OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Dematic\\Epia Shell";

        public static string KIMBERLY_CLARK_SERVER_ROOT = OSVersionInfoClass.ProgramFilesx86FolderName() +
                                                          "\\Egemin\\Ewcs Server";

        public static string DEPLOY_TESTS_DIRECTORY = OSVersionInfoClass.ProgramFilesx86FolderName() +
                                                      @"\Dematic\AutomaticTesting";
    }
}