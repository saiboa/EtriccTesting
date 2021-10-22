using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
//using Microsoft.TeamFoundation.Build.Proxy;
using Microsoft.TeamFoundation.Client;

namespace Epia3Deployment
{
    class BuildUtilities
    {
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Methods of BuildUtilities (5)
        
        public static string getBuildBasePath(string path)
        {
            int ib = path.LastIndexOf("\\");
            while (ib > 0)
            {
                string y = path.Substring(ib + 1);
                if ((y.StartsWith("Epia") || y.StartsWith("Etricc") || y.StartsWith("Release") || y.StartsWith("KC")) && y.IndexOf("-") > 0 && y.IndexOf(".") > 0)
                {
                    return path;
                }
                path = path.Substring(0, ib);
                ib = path.LastIndexOf("\\");
            }

            return string.Empty;
        }

        //  X:\Version\Epia 3\Epia\Epia - Version_20090414.2  -->  Epia - Version_20090414.2
        //  X:\Nightly\Etricc 5\Etricc - Nightly_20100223.1 --> Etricc - Nightly_20100223.1
        //  X:\Nightly\Ewcs - Projects\Kimberly Clark\KC - Nightly_20100127.1  --> KC - Nightly_20100127.1
        //  X:\Version\Etricc 5\Release 5.5.0 of Etricc - Version_20100129.1 --> Release 5.5.0 of Etricc - Version_20100129.1
        public static string getBuildnr(string path)
        {
            string nr = string.Empty;
            int ib = path.LastIndexOf("\\");
            while (ib > 0)
            {
                string y = path.Substring(ib + 1);
                //MessageBox.Show(path.Substring(ib + 1));
                if ((y.StartsWith("Epia") || y.StartsWith("Etricc") || y.StartsWith("Release")
                    || y.StartsWith("KC")) && y.IndexOf("_") > 0 && y.IndexOf(".") > 0 )
                    nr = y;

                path = path.Substring(0, ib);
                ib = path.LastIndexOf("\\");
            }

            return nr;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="validBuildDir">X:\Version\Epia 3\Epia\Epia - Version_20090414.2</param>
        /// <returns>Epia</returns>
        public static string getTestApplication(string validBuildDir)
        {
            string app = string.Empty;
            if (validBuildDir.IndexOf("Etricc 5") > 0)
                app = "Etricc 5";
            else if (validBuildDir.IndexOf("Etricc UI") > 0)
                app = "Etricc UI";
            else if (validBuildDir.IndexOf("Epia 4") > 0)
                app = "Epia";
            else if (validBuildDir.IndexOf("KC") > 0)
                app = "KC";
            else if (validBuildDir.IndexOf("Ewms") > 0)
                app = "Ewms";
            
            return app;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="validBuildDir"> X:\Version\Epia 3\Epia\Epia - Version_20090414.2 </param>
        /// <returns>Version </returns>
        public static string getTestDefinition(string validBuildDir)
        {
            string def = string.Empty;
            //int ia = validBuildDir.IndexOf(":");
            //def = validBuildDir.Substring(ia+2);
            //def = def.Substring(0, def.IndexOf("\\"));
             if (validBuildDir.IndexOf("Version") > 0)
                def = "Version";
            else if (validBuildDir.IndexOf("Nightly") > 0)
                def = "Nightly";
            else if (validBuildDir.IndexOf("CI") > 0)
                def = "CI";
            else if (validBuildDir.IndexOf("Weekly") > 0)
                def = "Weekly";
            else def = "No def error";
            return def;
        }
        
        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

    }
}
