using System;
using System.IO;
using System.Security.Principal;

namespace TestTools
{
    public class Epia3Common
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="installScriptsDir"></param>
        /// <param name="buildBaseDir"></param>
        /// <param name="buildNr"></param>
        /// <param name="testApp"></param>
        /// <param name="buildDef"></param>
        /// <param name="buildConfig"></param>
        public static void GetAllParameters(string installScriptsDir, ref string buildBaseDir,
                                            ref string buildNr, ref string testApp, ref string buildDef,
                                            ref string buildConfig)
        {
            buildBaseDir = getBuildBasePath(installScriptsDir);
            buildNr = getBuildnr(installScriptsDir);

            if (buildNr.StartsWith("Etricc"))
                testApp = "Etricc";
            else if (buildNr.StartsWith("Epia"))
                testApp = "Epia";
            else if (buildNr.StartsWith("KC"))
                testApp = "Kimberly Clark";

            int ib = buildNr.IndexOf("-");
            int ie = buildNr.IndexOf("_");
            buildDef = buildNr.Substring(ib + 1, ie - (ib + 1)).Trim();

            if (installScriptsDir.IndexOf("Debug") > 0)
                buildConfig = "Debug";
            else
                buildConfig = "Release";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path">the path that include buildnumber
        /// example: X:\Nightly\Epia 3\Epia\Epia - Nightly_20080528.1\Mixed Platforms\Debug\InstallScripts
        /// it return X:\Nightly\Epia 3\Epia\Epia - Nightly_20080528.1
        /// </param>
        /// <returns></returns>
        public static string getBuildBasePath(string path)
        {
            int ib = path.LastIndexOf("\\");
            while (ib > 0)
            {
                string y = path.Substring(ib + 1);
                //MessageBox.Show(path.Substring(ib + 1));
                if ((y.StartsWith("Epia") || y.StartsWith("Etricc") || y.StartsWith("KC")) && y.IndexOf("-") > 0 &&
                    y.IndexOf(".") > 0)
                {
                    return path;
                }
                path = path.Substring(0, ib);
                ib = path.LastIndexOf("\\");
            }

            return string.Empty;
        }

        /// <summary>
        /// return a buildnr
        /// example: X:\Nightly\Epia 3\Epia\Epia - Nightly_20080528.1\Mixed Platforms\Debug\InstallScripts
        /// return : Epia - Nightly_20080528.1
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string getBuildnr(string path)
        {
            string nr = string.Empty;
            int ib = path.LastIndexOf("\\");
            while (ib > 0)
            {
                string y = path.Substring(ib + 1);
                if ((y.StartsWith("Epia") || y.StartsWith("Etricc") || y.StartsWith("KC")) && y.IndexOf("-") > 0 &&
                    y.IndexOf(".") > 0)
                    nr = y;
                // one level up
                path = path.Substring(0, ib);
                ib = path.LastIndexOf("\\");
            }

            return nr;
        }

        public static void CreateTestLog(ref string slogFilePath, string outFilePath, string outFilename,
                                         ref StreamWriter Writer)
        {
            string logPath = string.Empty;
            if (WindowsIdentity.GetCurrent().Name.ToUpper().StartsWith("TEAMSYSTEMS\\JIEMINSHI"))
            {
                logPath = Path.Combine(Directory.GetCurrentDirectory(),
                                       outFilename + "-" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".log");
            }
            else
                logPath = Path.Combine(outFilePath, outFilename + ".log");


            slogFilePath = logPath;

            if (!File.Exists(logPath))
            {
                Writer = File.CreateText(logPath);
                Console.WriteLine("===" + logPath + "------  not exist");
            }
            else
            {
                Console.WriteLine("===" + logPath + "------ exist");
                Writer = File.AppendText(logPath);
                Writer.WriteLine("\n");
            }

            Writer.WriteLine(DateTime.Now + "\tTest Results for Test Run; ");
            Writer.WriteLine(DateTime.Now + "\t============================");
            Writer.Close();
        }


        public static void WriteTestLogTitle(string slogFilePath, string title, int count, bool test)
        {
            if (test)
                return;

            try
            {
                StreamWriter sw = null;
                try
                {
                    sw = File.AppendText(slogFilePath);
                    //Path.Combine( logFilePath, logFileName ));
                    string logLine = String.Format(
                        "{0:G}: {1}.", DateTime.Now, "(" + count + ") ==========" + title);
                    sw.WriteLine("");
                    sw.WriteLine(logLine);
                }
                finally
                {
                    sw.Close();
                }
            }
            catch (Exception)
            {
            }
        }

        public static void WriteTestLogMsg(string slogFilePath, string msg, bool test)
        {
            if (test)
                return;

            try
            {
                StreamWriter sw = null;
                try
                {
                    sw = File.AppendText(slogFilePath);
                    string logLine = String.Format("{0:G}: {1}.", DateTime.Now, "\t" + msg);
                    sw.WriteLine(logLine);
                }
                finally
                {
                    sw.Close();
                }
            }
            catch (Exception)
            {
            }
        }

        public static void WriteTestLogPass(string slogFilePath, string testcase, bool test)
        {
            if (test)
                return;

            try
            {
                StreamWriter sw = null;
                try
                {
                    sw = File.AppendText(slogFilePath);
                    string logLine = String.Format("{0:G}: {1}.", DateTime.Now, "\t" + testcase + ": Test Passed:");
                    sw.WriteLine(logLine);
                }
                finally
                {
                    sw.Close();
                }
            }
            catch (Exception)
            {
            }
        }

        public static void WriteTestLogFail(string slogFilePath, string testcase, bool test)
        {
            if (test)
                return;

            try
            {
                StreamWriter sw = null;
                try
                {
                    sw = File.AppendText(slogFilePath);
                    string logLine = String.Format("{0:G}: {1}.", DateTime.Now, "\t" + testcase + ": Test Failed:");
                    sw.WriteLine(logLine);
                }
                finally
                {
                    sw.Close();
                }
            }
            catch (Exception)
            {
            }
        }

        public static void CloseTestLog(string slogFilePath, bool test)
        {
            if (test)
                return;

            try
            {
                StreamWriter sw = null;
                try
                {
                    sw = File.AppendText(slogFilePath);
                    string logLine = String.Format("{0:G}: {1}.", DateTime.Now, "\tTest Completed for Test Run; ");
                    sw.WriteLine(logLine);
                }
                finally
                {
                    sw.Close();
                }
            }
            catch (Exception)
            {
            }
        }
    }
}