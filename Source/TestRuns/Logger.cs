using System;
using System.IO;

namespace TestRuns
{
    /// <summary>
    /// Summary description for Logger.
    /// </summary>
    public class Logger
    {
        #region Fields

        private static string slogFilePath = @"C:\";

        #endregion

        #region Constructors/Destructors/Cleanup

        /// <summary>
        /// Default constructor.
        /// </summary>
        public Logger()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public Logger(string logfilepath)
        {
            slogFilePath = logfilepath;
        }

        #endregion

        #region Methods

        public string GetTempPath()
        {
            string path = Environment.GetEnvironmentVariable("TEMP");
            if (!path.EndsWith("\\")) path += "\\";
            return path;
        }

        public void LogMessageToFile(string msg)
        {
            if (msg.Trim().Length == 0)
                msg = "\t\t ------ ";
            else
                msg = "\t\t ------ " + msg;

            StreamWriter sw = File.AppendText(slogFilePath);
            //Path.Combine( logFilePath, logFileName ));
            try
            {
                string logLine = String.Format(
                    "{0:G}: {1}.", DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }

        public void LogTestRunStartup(string testid)
        {
            string msg = " ------    Test: " + testid + "-------   start run  ------";
            StreamWriter sw = File.AppendText(slogFilePath);
            try
            {
                string logLine = String.Format(
                    "{0:G}: {1}.", DateTime.Now, msg);
                sw.WriteLine("");
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }

        public void LogCreatedJob(string id, string type, string locID, string agvID)
        {
            string msg = "\t\t ------    Created Job with ID: " + id + " type: " + type + " at " + locID + " by " +
                         agvID;
            ;
            StreamWriter sw = File.AppendText(slogFilePath);
            try
            {
                string logLine = String.Format(
                    "{0:G}: {1}.", DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }

        public void LogPassLine(string arg)
        {
            if (arg.Trim().Length == 0)
                arg = "\t\t ------ PASS:";
            else
                arg = "\t\t ------   PASS:\t\t " + arg;

            StreamWriter sw = File.AppendText(slogFilePath);
            try
            {
                string logLine = String.Format(
                    "{0:G}: {1}.", DateTime.Now, arg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }

        public void LogFailLine(string arg)
        {
            string msg = "\t\t ------   FAIL:\t\t " + arg;
            StreamWriter sw = File.AppendText(slogFilePath);
            try
            {
                string logLine = String.Format(
                    "{0:G}: {1}.", DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }


        public void LogCreatedTransport(string id, string type, string sourceID, string destID, string agvID)
        {
            string msg = "\t\t ------    Created Transport: ID " + id + " type: " + type + " at " + destID + " by " +
                         agvID + " with priority 5";
            if (type == "MOVE")
                msg = "\t\t ------    Created Transport: ID " + id + " type: " + type + " at " + sourceID + " to " +
                      destID + " by " + agvID + " with priority 5";

            StreamWriter sw = File.AppendText(slogFilePath);
            try
            {
                string logLine = String.Format(
                    "{0:G}: {1}.", DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }

        public void LogCreatedTransport(string id, string type, string sourceID, string destID, string agvID,
                                        int priority)
        {
            string msg = "\t\t ------    Created Transport: ID " + id + " type: " + type + " at " + destID + " by " +
                         agvID + " with priority " + priority;
            if (type.ToUpper() == "MOVE")
                msg = "\t\t ------    Created Transport: ID " + id + " type: " + type + " at " + sourceID + " to " +
                      destID + " by " + agvID + " with priority " + priority;
            else if (type.ToUpper() == "PICK")
                msg = "\t\t ------    Created Transport: ID " + id + " type: " + type + " at " + sourceID + " by " +
                      agvID + " with priority " + priority;
            else if (type.ToUpper() == "DROP")
                msg = "\t\t ------    Created Transport: ID " + id + " type: " + type + " at " + destID + " by " + agvID +
                      " with priority " + priority;
            else
                msg = "\t\t ------    Created Transport: ID " + id + " type: " + type;


            StreamWriter sw = File.AppendText(slogFilePath);
            try
            {
                string logLine = String.Format(
                    "{0:G}: {1}.", DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }

        public void LogTestException(string exMsg, string exStack)
        {
            string msg = "\t\t ------    Exception: " + exMsg + " \r\n " + exStack;
            //string logLine = System.String.Format("{0:G}: {1}.", System.DateTime.Now, msg);

            StreamWriter sw = File.AppendText(slogFilePath);
            try
            {
                string logLine = String.Format(
                    "{0:G}: {1}.", DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }

        #endregion
    }
}