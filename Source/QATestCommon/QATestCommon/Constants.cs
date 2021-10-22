using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QATestCommon
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
}
