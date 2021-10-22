using System;
using System.IO;

namespace TestTools
{
    /// ¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤?
    /// <summary>
    /// Summary description for logger.
    /// </summary>
    public class Logger
    {
        #region Fields of Logger (1)

        private static string slogFilePath = @"C:\";

        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        #region Constructors/Destructors/Cleanup of Logger (2)

        /// <summary>
        /// Default constructor.
        /// </summary>
        public Logger()
        {
            // TODO: Add constructor logic here
        }

        public Logger(string logfilepath)
        {
            slogFilePath = logfilepath;
        }

        #endregion // —— Constructors/Destructors/Cleanup •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        #region Methods of Logger (2)

        public string GetTempPath()
        {
            string path = Environment.GetEnvironmentVariable("TEMP");
            if (!path.EndsWith("\\")) path += "\\";
            return path;
        }

        public void LogMessageToFile(string msg, int logCount, int interval)
        {
            if ((interval == 0) || logCount%(interval) == 0)
            {
                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                StreamWriter sw = null;
                try
                {
                    sw = File.AppendText(slogFilePath);
                    sw.WriteLine(logLine);
                }
                catch (Exception)
                {
                    System.Threading.Thread.Sleep(2000);
                }
                finally
                {
                    sw.Close();
                }

                //using (System.IO.StreamWriter sw = System.IO.File.AppendText(slogFilePath)) 
                //{
                //    sw.WriteLine(logLine); 
                //} 
            }
        }

        public void LogMessageToFile(string msg, int logCount, int interval, string exceptionToLocal)
        {
            if ((interval == 0) || logCount % (interval) == 0)
            {
                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                StreamWriter sw = null;
                bool writeOK = false;
                while (writeOK == false)
                {
                    try
                    {
                        sw = File.AppendText(slogFilePath);
                        sw.WriteLine(logLine);
                        writeOK = true;
                    }
                    catch (System.IO.IOException ex)
                    {
                        if (ex.Message.IndexOf("being used by another process") >= 0)
                        {
                            writeOK = false;
                            System.Threading.Thread.Sleep(10000);
                            Console.WriteLine("write log msg : " + msg);
                            Console.WriteLine("write log exception : " + ex.Message);
                        }
                        else
                        {
                            writeOK = false;
                            System.Threading.Thread.Sleep(10000);
                            Console.WriteLine("write log msg : " + msg);
                            Console.WriteLine("write log exception : " + ex.Message);
                            //throw;
                        }
                    }
                    finally
                    {
                        if (sw != null)
                            sw.Close();
                    }
                }

                //using (System.IO.StreamWriter sw = System.IO.File.AppendText(slogFilePath)) 
                //{
                //    sw.WriteLine(logLine); 
                //} 
            }
        }

        public string GetLogPath()
        {
            return slogFilePath;
        }

        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
    }
}