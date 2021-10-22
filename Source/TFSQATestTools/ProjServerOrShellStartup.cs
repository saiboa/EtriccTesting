using System;
using System.ServiceProcess;
using System.Threading;
using TestTools;

namespace TFSQATestTools
{
    public class ProjServerOrShellStartup
    {
        public static bool CheckThisServiceIsStartedUp(string serviceName, ref string errorMsg, string slogFilePath,
                                                       bool sOnlyUITest)
        {
            bool startupOK = true;
            Console.WriteLine("CheckThisServiceIsStartup : " + serviceName);

            var controller = new ServiceController();
            controller.MachineName = Environment.MachineName;
            controller.ServiceName = serviceName; // "MSSQLSERVER"; or ReportServer

            if (controller.Status == ServiceControllerStatus.Running)
            {
                Epia3Common.WriteTestLogMsg(slogFilePath, serviceName + " is running: ", sOnlyUITest);
                startupOK = true;
            }
            else
            {
                Console.WriteLine(serviceName + " has status " + controller.Status.ToString());
                controller.Start();
                // wait until service status is Running
                DateTime sStartTime = DateTime.Now;
                TimeSpan sTime = DateTime.Now - sStartTime;
                int wat = 0;
                var svcEpia = new ServiceController(serviceName);
                string epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                Console.WriteLine(svcEpia.ServiceName + " has status " + svcEpia.Status.ToString());
                Console.WriteLine(" time is :" + sTime.TotalSeconds);
                while ((!epiaServiceStatus.StartsWith("running")) && wat < 600)
                {
                    Thread.Sleep(2000);
                    wat = wat + 2;
                    //serviceStatus = controller.Status.ToString().ToLower();
                    Console.WriteLine("wait " + serviceName + " service status running:  time is (sec) : " + wat +
                                      "  and status is:" + epiaServiceStatus);

                    svcEpia = new ServiceController(serviceName);
                    epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                    Console.WriteLine("--- "+svcEpia.ServiceName + " has status " + svcEpia.Status.ToString());
                    Console.WriteLine("--- " + " time is :" + sTime.TotalSeconds);
                }

                if (svcEpia.Status != ServiceControllerStatus.Running)
                {
                    errorMsg = serviceName + "Service Startup failed:" + epiaServiceStatus;
                    Epia3Common.WriteTestLogMsg(slogFilePath, errorMsg, sOnlyUITest);
                    startupOK = false;
                }
            }
            return startupOK;
        }

        public static bool ServerStartup(string serverFolderName, string sServerRunAs, ref string errorMsg,
                                         string slogFilePath, bool sOnlyUITest)
        {
            bool startupOK = true;
            string serviceName = "Egemin Epia Server";
            string serverPath = OSVersionInfoClass.ProgramFilesx86() + @"\Dematic\" + serverFolderName;
            string ApplicationFileName = ConstCommon.EGEMIN_EPIA_SERVER_EXE;
            if (serverFolderName.Equals("Etricc Server"))
            {
                serviceName = "Egemin Etricc Server";
                ApplicationFileName = ConstCommon.EGEMIN_ETRICC_SERVER_EXE;
            }

            Console.WriteLine("sServerRunAs : " + sServerRunAs);
            if (sServerRunAs.ToLower().IndexOf("service") >= 0)
            {
                // uninstall Egemin.Epia.server Service
                Console.WriteLine("UNINSTALL " + serviceName + " Service : ");
                Console.WriteLine("serverPath " + serviceName + " serverPath : ");
                Console.WriteLine("serviceName:  " + serviceName + " ApplicationFileName : "+ ApplicationFileName);
                ProcessUtilities.StartProcessWaitForExit(serverPath, ApplicationFileName, " /u");
                Thread.Sleep(2000);

                Console.WriteLine("INSTALL " + serviceName + " Service : ");
                ProcessUtilities.StartProcessWaitForExit(serverPath, ApplicationFileName, " /i");
                Thread.Sleep(2000);

                // wait until service installed
                int watt = 0;
                var svcServiceX = new ServiceController(serviceName);
                Console.WriteLine(" time is 1111 :" );
                Console.WriteLine("wait " + serviceName + " service status exist:  time is (sec) : " + watt +
                                      "  and status is:" + svcServiceX.Status.ToString().ToLower());

                Console.WriteLine(" time is 22222 :");
                string svcServiceXStatus = null;
                while (svcServiceXStatus == null && watt < 120)
                {
                    try
                    {
                        svcServiceXStatus = svcServiceX.Status.ToString().ToLower();
                    }
                    catch (Exception ex)
                    {
                        
                        Console.WriteLine(serviceName + " -- get service status exception : " + ex.Message +
                                       "  and ex.StackTrace:" + ex.StackTrace);

                        Thread.Sleep(5000);
                        svcServiceXStatus = null;

                    }
                }
                Console.WriteLine("Check " + serviceName + " status : "+ svcServiceXStatus);
                Console.WriteLine("Start " + serviceName + " Service : ");
                ProcessUtilities.StartProcessWaitForExit(serverPath, ApplicationFileName, " /start");
                Thread.Sleep(2000);

                var svcEpia = new ServiceController(serviceName);
                Console.WriteLine(svcEpia.ServiceName + " has status " + svcEpia.Status.ToString());
                // wait until epia service status is Running
                DateTime sStartTime = DateTime.Now;
                TimeSpan sTime = DateTime.Now - sStartTime;
                int wat = 0;
                string epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                Console.WriteLine(" time is :" + sTime.TotalSeconds);
                while (!epiaServiceStatus.StartsWith("running") && wat < 60)
                {
                    Thread.Sleep(2000);
                    wat = wat + 2;
                    epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                    Console.WriteLine("wait " + serviceName + " service status running:  time is (sec) : " + wat +
                                      "  and status is:" + epiaServiceStatus);
                }

                if (svcEpia.Status != ServiceControllerStatus.Running)
                {
                    errorMsg = serviceName + "Service Startup failed:";
                    Epia3Common.WriteTestLogMsg(slogFilePath,
                                                serviceName + " Service startup failed: " + epiaServiceStatus,
                                                sOnlyUITest);
                    startupOK = false;
                    //sEpiaServerStartupOK = false;
                    //throw new Exception("Epia service start up failed:"); //   get message from log file sErrorMessage//
                }
            }
            else if (sServerRunAs.ToLower().IndexOf("console") >= 0)
            {
                Console.WriteLine("Start " + serviceName + "  as console applications : ");
                // Start Epia SERVER as Console
                ProcessUtilities.StartProcessNoWait(serverPath, ApplicationFileName, string.Empty);
                Thread.Sleep(20000);
                Console.WriteLine(serverFolderName + " Started : ");
            }

            return startupOK;
        }

        public static bool ServerStartup(string branding, string serverFolderName, string sServerRunAs, ref string errorMsg,
                                         string slogFilePath, bool sOnlyUITest)
        {
            bool startupOK = true;
            string serviceName = "Egemin Epia Server";
            string serverPath = OSVersionInfoClass.ProgramFilesx86() + "\\" + branding + "\\" + serverFolderName;
            string ApplicationFileName = ConstCommon.EGEMIN_EPIA_SERVER_EXE;
            if (serverFolderName.Equals("Etricc Server"))
            {
                serviceName = "Egemin Etricc Server";
                ApplicationFileName = ConstCommon.EGEMIN_ETRICC_SERVER_EXE;
            }

            Console.WriteLine("sServerRunAs : " + sServerRunAs);
            if (sServerRunAs.ToLower().IndexOf("service") >= 0)
            {
                // uninstall Egemin.Epia.server Service
                Console.WriteLine("UNINSTALL " + serviceName + " Service : ");
                Console.WriteLine("serverPath " + serviceName + " serverPath : ");
                Console.WriteLine("serviceName:  " + serviceName + " ApplicationFileName : " + ApplicationFileName);
                ProcessUtilities.StartProcessWaitForExit(serverPath, ApplicationFileName, " /u");
                Thread.Sleep(2000);

                Console.WriteLine("INSTALL " + serviceName + " Service : ");
                ProcessUtilities.StartProcessWaitForExit(serverPath, ApplicationFileName, " /i");
                Thread.Sleep(2000);

                // wait until service installed
                int watt = 0;
                var svcServiceX = new ServiceController(serviceName);
                Console.WriteLine(" time is 1111 :");
                Console.WriteLine("wait " + serviceName + " service status exist:  time is (sec) : " + watt +
                                      "  and status is:" + svcServiceX.Status.ToString().ToLower());

                Console.WriteLine(" time is 22222 :");
                string svcServiceXStatus = null;
                while (svcServiceXStatus == null && watt < 120)
                {
                    try
                    {
                        svcServiceXStatus = svcServiceX.Status.ToString().ToLower();
                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine(serviceName + " -- get service status exception : " + ex.Message +
                                       "  and ex.StackTrace:" + ex.StackTrace);

                        Thread.Sleep(5000);
                        svcServiceXStatus = null;

                    }
                }
                Console.WriteLine("Check " + serviceName + " status : " + svcServiceXStatus);
                Console.WriteLine("Start " + serviceName + " Service : ");
                ProcessUtilities.StartProcessWaitForExit(serverPath, ApplicationFileName, " /start");
                Thread.Sleep(2000);

                var svcEpia = new ServiceController(serviceName);
                Console.WriteLine(svcEpia.ServiceName + " has status " + svcEpia.Status.ToString());
                // wait until epia service status is Running
                DateTime sStartTime = DateTime.Now;
                TimeSpan sTime = DateTime.Now - sStartTime;
                int wat = 0;
                string epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                Console.WriteLine(" time is :" + sTime.TotalSeconds);
                while (!epiaServiceStatus.StartsWith("running") && wat < 60)
                {
                    Thread.Sleep(2000);
                    wat = wat + 2;
                    epiaServiceStatus = svcEpia.Status.ToString().ToLower();
                    Console.WriteLine("wait " + serviceName + " service status running:  time is (sec) : " + wat +
                                      "  and status is:" + epiaServiceStatus);
                }

                if (svcEpia.Status != ServiceControllerStatus.Running)
                {
                    errorMsg = serviceName + "Service Startup failed:";
                    Epia3Common.WriteTestLogMsg(slogFilePath,
                                                serviceName + " Service startup failed: " + epiaServiceStatus,
                                                sOnlyUITest);
                    startupOK = false;
                    //sEpiaServerStartupOK = false;
                    //throw new Exception("Epia service start up failed:"); //   get message from log file sErrorMessage//
                }
            }
            else if (sServerRunAs.ToLower().IndexOf("console") >= 0)
            {
                Console.WriteLine("Start " + serviceName + "  as console applications : ");
                // Start Epia SERVER as Console
                ProcessUtilities.StartProcessNoWait(serverPath, ApplicationFileName, string.Empty);
                Thread.Sleep(20000);
                Console.WriteLine(serverFolderName + " Started : ");
            }

            return startupOK;
        }
    }
}