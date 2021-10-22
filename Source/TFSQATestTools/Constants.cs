using System.Configuration;
using TestTools;

namespace TFSQATestTools
{
    public static class Tfs
    {
        public static string ServerUrl = System.Configuration.ConfigurationManager.AppSettings.Get( "TFSServerUri" );
        public static string Server = System.Configuration.ConfigurationManager.AppSettings.Get( "TFSServer" );
        public static string UserName = System.Configuration.ConfigurationManager.AppSettings.Get( "TFSUsername" );
        public static string Password = System.Configuration.ConfigurationManager.AppSettings.Get( "TFSPassword" );
        public static string Domain = System.Configuration.ConfigurationManager.AppSettings.Get( "TFSDomain" );

        public static int ReconnectDelay = 1 * 60000; // 1 minute
    }

    public static class Constants
    {
        // Test Params
        public static string sDemonstration = ConfigurationManager.AppSettings.Get("Demonstration");
        public static string sDemoCaseCount = ConfigurationManager.AppSettings.Get("DemoCaseCount");
        public static string sTestResultFolderName = ConfigurationManager.AppSettings.Get("TestResultFolderName");
        public static string sMsgDebug = System.Configuration.ConfigurationManager.AppSettings.Get("MsgDebug");
        public static string sDeploymentLogFilename = System.Configuration.ConfigurationManager.AppSettings.Get("LogFilename");
        public static string sEpia4InstallerName = System.Configuration.ConfigurationManager.AppSettings.Get("Epia4InstallerName");
        public static string sRecompileDotnetVersion = System.Configuration.ConfigurationManager.AppSettings.Get("RecompileDotnetVersion");
        public static string STimerInterval = ConfigurationManager.AppSettings.Get("TimerInterval");

        // Test Apps
        //public const string EPIA4 = "Epia4";
        public const string ETRICCUI = "EtriccUI";
        public const string ETRICC5 = "Etricc5";
        public const string ETRICCSTATISTICS = "EtriccStatistics";
        public const string EWMS = "Ewms";

        public const string ETRICCSTATISTICS_UI = "EtriccStatisticsUI";
        public const string ETRICCSTATISTICS_PARSERCONFIGURATOR = "EtriccStatisticsParserConfigurator";
        public const string ETRICCSTATISTICS_PARSER_SETUP = "EtriccStatisticsParser";
       
        // Test Projects
        public const string EPIA_3 = "Epia 3";
        public const string EPIA_4 = "Epia 4";
        public const string ETRICC_5 = "Etricc 5";
        public const string EWCS_PROJECTS = "Ewcs - Projects";

        public const string ETRICC = "Etricc";
        public const string ETRICC_UI = "Etricc UI";
        public const string EPIA3_DEPLOYMENT_CONFIG_FILE = "Epia3Deployment.configXML";
        public const string KIMBERLY_CLARK = "Kimberly Clark";
        public const string KC = "KC";

        // Config File 
        public const string TfsSettingsSection = "TFS.Settings.Section";
        public const string TestConfigSection = "Test.Config.Section";

        // Msi files
        public const string EPIA_MSI = "Epia.msi";
        public const string EPIA_RESOURCEFILE_EDITOR_MSI = "Epia.ResourceFileEditor.msi";
        public const string ETRICC_CORE_MSI = "Etricc.msi";
        public const string ETRICC_SHELL_MSI = "Etricc Shell.msi";
        public const string ETRICC_PLAYBACK_MSI = "Etricc Playback.msi";
        public const string ETRICC_HOSTTEST_MSI = "Etricc HostTest.msi";
        public const string ETRICC_STATISTICS_PARSER_MSI = "Etricc.Statistics.Parser.msi";
        public const string ETRICC_STATISTICS_PARSER_CONFIGURATOR_MSI = "Etricc.Statistics.ParserConfigurator.msi";
        public const string ETRICC_STATISTICS_UI_MSI = "Etricc.Statistics.UI.msi";

        public const string HIDDEN_USERNAME = "Egemin";
        public const string HIDDEN_PASSWORD = "herculepoirot";
      
    }

    public static class EgeminApplication
    {
        public const string EPIA = "Epia";
        public const string EPIA_RESOURCEFILEEDITOR = "Epia.ResourceFileEditor";
        public const string ETRICC = "Etricc";
        public const string ETRICC_SHELL = "Etricc Shell";
        public const string ETRICC_HOSTTEST = "Etricc HostTest";
        public const string ETRICC_PLAYBACK = "Etricc Playback";

        public const string ETRICC_STATISTICS_PARSER = "Etricc.Statistics.Parser";
        public const string ETRICC_STATISTICS_PARSERCONFIGURATOR = "Etricc.Statistics.ParserConfigurator";
        public const string ETRICC_STATISTICS_UI = "Etricc.Statistics.UI";

        public const string EPIA_LAUNCHER = "EPIA.Launcher";
        public const string EPIA_EXPLORER = "EPIA.Explorer";

        public const string AUTOMATICTESTING = "AutomaticTesting";

        public enum SetupType
        {
            Default,
            EpiaServerOnly,
            EpiaShellOnly,
        }
    }

    public static class TestApp
    {
        public const string EPIA4 = "Epia4";
        public const string ETRICCUI = "EtriccUI";
        public const string ETRICCSTATISTICS = "EtriccStatistics";
        public const string EPIANET45 = "EpiaNet45";
        public const string ETRICCNET45 = "EtriccNet45";
    }

    public static class TestDefName
    {
        public const string EPIA4 = "Epia4TestDefinition.txt";
        public const string ETRICCUI = "EtriccUITestTypeDefinition.txt";
        public const string ETRICCSTATISTICS = "StatisticsTestTypeDefinition.txt";
        public const string EPIANET45 = "EpiaNet45TestDefinition.txt";
        public const string ETRICCNET45 = "EtriccNet45TestDefinition.txt";
    }



}
