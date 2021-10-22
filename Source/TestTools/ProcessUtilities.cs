using System;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Threading;
using System.Windows.Automation;

namespace TestTools
{
    public class ProcessUtilities
    {
        public static void CloseProcess(string processName)
        {
            Process[] ps = Process.GetProcessesByName(processName);
            try
            {
                for (int i = 0; i < ps.Length; i++)
                {
                    ps[i].Kill();
                    Console.WriteLine("Close " + processName + " at:" + DateTime.Now.ToString("HH:mm:ss"));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(" close " + processName + " exception:" + ex.Message, ex.StackTrace);
            }
        }

        public static Process StartProcessNoWait(string processDir, string procFilename, string args)
        {
            string path = Path.Combine(processDir, procFilename);
            var proc = new Process();
            proc.EnableRaisingEvents = false;
            proc.StartInfo.FileName = path;
            proc.StartInfo.Arguments = args;
            proc.StartInfo.WorkingDirectory = processDir;
            proc.Start();
            Thread.Sleep(5000);
            return proc;
            //proc5.WaitForExit();
        }

        public static void StartProcessWaitForExit(string processDir, string procFilename, string args)
        {
            string path = Path.Combine(processDir, procFilename);
            var proc = new Process();
            proc.EnableRaisingEvents = false;
            proc.StartInfo.FileName = path;
            proc.StartInfo.Arguments = args;
            proc.StartInfo.WorkingDirectory = processDir;
            proc.Start();
            proc.WaitForExit();
        }

        public static int GetApplicationProcessID(string processName, out Process proc)
        {
            int pID = 0;
            proc = null;
            //System.Diagnostics.Process[] pShell = System.Diagnostics.Process.GetProcessesByName("Egemin.Epia.Presentation.CompositeUI.Shell");
            Process[] procs = Process.GetProcessesByName(processName);
            Console.WriteLine(processName + " procs.Length:" + procs.Length + " at " + DateTime.Now.ToString("HH:mm:ss"));
            try
            {
                for (int i = 0; i < procs.Length; i++)
                {
                    if (procs[i].Responding)
                    {
                        pID = procs[i].Id;
                        proc = procs[i];
                        Console.WriteLine("Proc ID:" + pID);
                    }
                    else
                    {
                        procs[i].Kill();
                        Console.WriteLine("Kill Proc at:" + DateTime.Now.ToString("HH:mm:ss"));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(processName + " check proc exception:" + ex.Message, ex.StackTrace);
                pID = -1;
            }

            return pID;
        }

        public static bool StartProcessAndWaitUntilResponding(string processDir, string procFilename,
                                                              string args, string procName, int waitMinute)
        {
            bool status = false;
            try
            {
                var proc = new Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = procFilename;
                if (args != null)
                    proc.StartInfo.Arguments = args;
                proc.StartInfo.WorkingDirectory = processDir;
                proc.Start();
                //proc.WaitForExit();

                DateTime startTime = DateTime.Now;
                Console.WriteLine(procName + " has id:" + proc.Id);
                Console.WriteLine("Responding:" + proc.Responding);
                while (!proc.Responding)
                {
                    Thread.Sleep(1000);
                    TimeSpan mTime = DateTime.Now - startTime;
                    if (proc.Responding)
                    {
                        Console.WriteLine(procName + " is responding" + DateTime.Now.ToString("HH:mm:ss"));
                        status = true;
                        break;
                    }
                    else
                    {
                        if (mTime.TotalSeconds >= waitMinute*60)
                        {
                            Console.WriteLine("after " + waitMinute + " min " + procName + " no responding" +
                                              DateTime.Now.ToString("HH:mm:ss"));
                            status = false;
                            break;
                        }
                        else
                            Console.WriteLine(procName + " no responding" + DateTime.Now.ToString("HH:mm:ss"));
                    }
                }
                return status;
            }
            catch (Exception ex)
            {
                Console.WriteLine("proc not exist:" + ex.Message, ex.StackTrace);
                return false;
            }
        }

        public static bool StartProcessAndWaitUntilUIWindowFound(string processDir, string procFilename,
                                                                 string args, string procName, int waitMinute,
                                                                 ref AutomationElement testForm)
        {
            try
            {
                string path = Path.Combine(processDir, procFilename);
                Process proc = Process.Start(path);
                Console.WriteLine("*****" + proc.Id);
                Thread.Sleep(90000);

                IntPtr pt = proc.MainWindowHandle;
                while (pt == IntPtr.Zero)
                {
                    Console.WriteLine("intPtr is zero!");
                    Thread.Sleep(1000);
                }

                Console.WriteLine("intPtr is:" + pt);

                testForm = AutomationElement.FromHandle(proc.MainWindowHandle);

                if (testForm == null)
                {
                    throw new Exception("Failed to find Window:" + procName);
                }
                else
                {
                    Console.WriteLine("Found it!");
                    Console.WriteLine("Found Main window Handle . . . " + proc.MainWindowHandle.ToString());
                    Console.WriteLine("Found Main window Title . . . " + proc.MainWindowTitle);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("proc not exist:" + ex.Message + "  ----  " + ex.StackTrace);
                return false;
            }
        }

        public static string RunProcessAndGetOutput(string exePath, string param)
        {
            string outputText = string.Empty;
            StreamReader outputReader = null;
            StreamReader errorReader = null;

            try
            {
                //Create Process Start information
                var processStartInfo =
                    new ProcessStartInfo(exePath, param);

                processStartInfo.ErrorDialog = false;
                processStartInfo.UseShellExecute = false;
                processStartInfo.RedirectStandardError = true;
                processStartInfo.RedirectStandardInput = true;
                processStartInfo.RedirectStandardOutput = true;

                //Execute the process
                var process = new Process();
                process.StartInfo = processStartInfo;
                bool processStarted = process.Start();
                if (processStarted)
                {
                    //Get the output stream
                    outputReader = process.StandardOutput;
                    errorReader = process.StandardError;
                    process.WaitForExit();

                    //Display the result
                    outputText = "Output:" + Environment.NewLine + "\\t\\t==============" + Environment.NewLine;
                    outputText += outputReader.ReadToEnd();
                    outputText += Environment.NewLine + "\\t\\tErr:" + Environment.NewLine + "\\t\\t==============" +
                                  Environment.NewLine;
                    outputText += errorReader.ReadToEnd();
                    //MessageBox.Show(outputText);
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message + "-----" + ex.StackTrace);
                outputText = "RunProcessAndGetOutput error:" + ex.Message + "----" + ex.StackTrace;
            }
            finally
            {
                if (outputReader != null)
                {
                    outputReader.Close();
                }
                if (errorReader != null)
                {
                    errorReader.Close();
                }
            }
            return outputText;
        }

        public static void SendTestResultToDevelopers(
            string resultFile, string layout, string buildType, Logger logger, int failedCounter,
            string testOverview, string testInputData, string sendMail)
        {
            try
            {
                var oMsg = new MailMessage();
                var oAttch = new Attachment(resultFile); //, System.Web.Mail.MailEncoding.Base64);; 
                SendEmailTo(resultFile, ref oMsg, ref oAttch, layout, buildType, failedCounter, testOverview,
                            testInputData, sendMail);

                logger.LogMessageToFile("--------------------------------", 0, 0);
                logger.LogMessageToFile("SmtpServer: " + ConstCommon.SMTP_SERVERID, 0, 0);
                logger.LogMessageToFile("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx send mail ======: "
                                        + sendMail, 0, 0);

                var client = new SmtpClient();
                client.Host = ConstCommon.SMTP_SERVERID;

                try
                {
                    client.Send(oMsg);
                }
                catch (Exception ex)
                {
                    logger.LogMessageToFile("The following exception occurred: " + ex, 0, 0);
                    //check the InnerException
                    while (ex.InnerException != null)
                    {
                        logger.LogMessageToFile("--------------------------------", 0, 0);
                        logger.LogMessageToFile("The following InnerException reported: "
                                                + ex.InnerException, 0, 0);
                        ex = ex.InnerException;
                    }
                }
                logger.LogMessageToFile("email sent to developers ", 0, 0);
                oMsg = null;
                oAttch = null;
            }
            catch (Exception e)
            {
                logger.LogMessageToFile("send mail : " + e.Message + " --- " + e.StackTrace, 0, 0);
                Console.WriteLine("{0} Exception caught.", e);
            }
        }

        public static void SendEmailTo(string xPath, ref MailMessage oMsg,
                                       ref Attachment oAttch, string layout, string buildType, int failedCounter,
                                       string testOverview, string testInputData, string sendMail)
        {
            int sendHour = DateTime.Now.Hour;
            oMsg.From = new MailAddress("teamsystems@egemin.be");
            //msg.Subject = "Greetings";
            //msg.Body = "This is a  message.";

            // TODO: Replace with recipient e-mail address.
            if (sendMail.ToLower().StartsWith("false"))
            {
                oMsg.To.Add("jiemin.shi@egemin.be");
                oMsg.Subject = testOverview + "[" + Environment.MachineName + "]" +
                               DateTime.Now.ToString("ddMMM-HH:mm");
            }
            else
            {
                if (failedCounter > 0)
                {
                    string strAll =
                        "jiemin.shi@egemin.be;Wim.VanBetsbrugge@egemin.be;Dirk.Declercq@egemin.be;Gunther.Storme@egemin.be;Walter.DeFeyter@egemin.be";
                    oMsg.To.Add(strAll);
                    //oMsg.To = "jiemin.shi@egemin.be;jiemin.shi@egemin.be;";
                    oMsg.Subject = "E'pia Nightly Test Result (" + layout + ")[" + buildType + "]" +
                                   DateTime.Now.ToString("ddMMM-HH:mm")
                                   + "-[" + Environment.MachineName + "]-" + testOverview;
                }
                else
                {
                    oMsg.To.Add("jiemin.shi@egemin.be;");
                    oMsg.Subject = "E'pia Nightly Test OK (" + layout + ")[" + buildType + "]" +
                                   DateTime.Now.ToString("ddMMM-HH:mm")
                                   + "-[" + Environment.MachineName + "]-" + testOverview;
                }
            }

            oMsg.IsBodyHtml = true;
            // HTML Body (remove HTML tags for plain text).
            //oMsg.Body = "<HTML><BODY><B>Hello World!</B></BODY></HTML>";
            oMsg.Body = testOverview + testInputData;
            oMsg.Attachments.Add(oAttch);
        }
    }
}