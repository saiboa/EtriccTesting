using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using TestTools;

namespace TFS2010AutoDeploymentTool
{
    class Constants
    {
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Constants of Constants (4)
        // TFSServer
        public static string sTFSServerUrl = System.Configuration.ConfigurationManager.AppSettings.Get("TFSServerUri");
        public static string sTFSServer = System.Configuration.ConfigurationManager.AppSettings.Get("TFSServer");
        public static string sTFSUsername = System.Configuration.ConfigurationManager.AppSettings.Get("TFSUsername");
        public static string sTFSPassword = System.Configuration.ConfigurationManager.AppSettings.Get("TFSPassword");
        public static string sTFSDomain = System.Configuration.ConfigurationManager.AppSettings.Get("TFSDomain");
        // test params
        public static string sDemonstration = ConfigurationManager.AppSettings.Get("Demonstration");
        public static string sTestResultFolderName = ConfigurationManager.AppSettings.Get("TestResultFolderName");
        public static string sMsgDebug = System.Configuration.ConfigurationManager.AppSettings.Get("MsgDebug");
        public static string sDeploymentLogFilename = System.Configuration.ConfigurationManager.AppSettings.Get("LogFilename");
        public static string sEpia4InstallerName = System.Configuration.ConfigurationManager.AppSettings.Get("Epia4InstallerName");
        public static string sEtricc5InstallationFolder = OSVersionInfoClass.ProgramFilesx86FolderName() + 
                        System.Configuration.ConfigurationManager.AppSettings.Get("Etricc5InstallationFolder");
        public static string sRecompileDotnetVersion = System.Configuration.ConfigurationManager.AppSettings.Get("RecompileDotnetVersion");

        // testApp
        public const string EPIA4 = "Epia4";
        public const string ETRICCUI = "EtriccUI";
        public const string ETRICC5 = "Etricc5";
        public const string ETRICCSTATISTICS = "EtriccStatistics";
        public const string EWMS = "Ewms";

        public const string ETRICCSTATISTICS_UI = "EtriccStatisticsUI";
        public const string ETRICCSTATISTICS_PARSERCONFIGURATOR = "EtriccStatisticsParserConfigurator";
        public const string ETRICCSTATISTICS_PARSER_SETUP = "EtriccStatisticsParser";
       
        // test Project
        public const string EPIA_3 = "Epia 3";
        public const string EPIA_4 = "Epia 4";
        public const string ETRICC_5 = "Etricc 5";
        public const string EWCS_PROJECTS = "Ewcs - Projects";

        public const string ETRICC = "Etricc";
        public const string ETRICC_UI = "Etricc UI";
        public const string EPIA3_DEPLOYMENT_CONFIG_FILE = "Epia3Deployment.configXML";
        public const string KIMBERLY_CLARK = "Kimberly Clark";
        public const string KC = "KC";

        // config file 
        public const string TfsSettingsSection = "TFS.Settings.Section";
        public const string TestConfigSection = "Test.Config.Section";

        // msi files
        public const string EPIA_MSI = "Epia.msi";
        public const string EPIA_RESOURCEFILE_EDITOR_MSI = "Epia.ResourceFileEditor.msi";
        public const string ETRICC_CORE_MSI = "Etricc.msi";
        public const string ETRICC_SHELL_MSI = "Etricc Shell.msi";
        public const string ETRICC_PLAYBACK_MSI = "Etricc Playback.msi";
        public const string ETRICC_HOSTTEST_MSI = "Etricc HostTest.msi";
        public const string ETRICC_STATISTICS_PARSER_MSI = "Etricc.Statistics.Parser.msi";
        public const string ETRICC_STATISTICS_PARSER_CONFIGURATOR_MSI = "Etricc.Statistics.ParserConfigurator.msi";
        public const string ETRICC_STATISTICS_UI_MSI = "Etricc.Statistics.UI.msi";
        

        #endregion // —— Constants ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
    }
}
