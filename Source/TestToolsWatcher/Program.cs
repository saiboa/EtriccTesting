using System;
using System.Collections.Generic;

using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.ServiceProcess;

namespace Egemin.Epia.Testing.TestToolsWatcher
{
    static class Program
    {
        static void Main(params string[] args)
        {
            var service = new Service1();

            if (!Environment.UserInteractive)
            {
                var servicesToRun = new ServiceBase[] { service };
                ServiceBase.Run(servicesToRun);
                return;
            }

            Console.WriteLine("Running as a Console Application");
            RunX();
            /*Console.WriteLine(" 1. Run Service");
            Console.WriteLine(" 2. Other Option");
            Console.WriteLine(" 3. Exit");
            Console.Write("Enter Option: ");

            var input = Console.ReadLine();

            switch (input)
            {
                case "1":
                    service.Start(args);
                    Console.WriteLine("Running Service - Press Enter To Exit");
                    Console.ReadLine();
                    break;
                case "2":
                    break;
            }*/
            Console.WriteLine("Closing");
        }
    

        static void RunX()
        {
            string remoteCommandFile = System.IO.Path.Combine(@"C:\EtriccTests", "RemoteCommand.txt");
            int ii = 0;
            while (true)
            {
                Console.WriteLine("start RUN ****************************\n");
                //System.Windows.Forms.MessageBox.Show("kk:: " + ii++);
                Thread.Sleep(10000);
                // Check test Working file
                System.IO.FileInfo workFile = new System.IO.FileInfo(remoteCommandFile);
                File.SetAttributes(workFile.FullName, FileAttributes.Normal);

                string info = string.Empty;
                bool ReadOK = false;
                while (ReadOK == false)
                {
                    try
                    {
                        StreamReader readerInfo = File.OpenText(remoteCommandFile);
                        info = readerInfo.ReadToEnd();
                        readerInfo.Close();
                        ReadOK = true;
                    }
                    catch (Exception ex)
                    {
                        ReadOK = false;
                        Thread.Sleep(5000);
                        Console.WriteLine("Read RemoteCommand.txt exception: " + ex.Message);
                    }
                }

                try
                {
                    #region process remote command
                    if (info.Length == 0)
                        Console.WriteLine(" --- Current command Empty -------------- ");
                    else
                    {
                        Console.WriteLine(" --- Current command Received -------------- " + info);

                        if (info.ToLower().StartsWith("forcestartup"))
                        {
                            Console.WriteLine("Start command Received -------------- " + info);
                            try
                            {
                                Console.WriteLine("number TFSQATestTools is running: = 000000000000000000000000000");
                                //System.Diagnostics.Process[] ps = System.Diagnostics.Process.GetProcessesByName("TFSQATestTools");
                                //int runTools = ps.Length;
                                //Console.WriteLine("number TFSQATestTools is running: = " + runTools);

                                string path = ProgramFilesx86() + @"\Egemin\AutomaticTesting";

                                string TotalPath = Path.Combine(path, "TFSQATestTools.exe");
                                var proc = new System.Diagnostics.Process();
                                proc.EnableRaisingEvents = false;
                                proc.StartInfo.FileName = "TFSQATestTools.bat";
                                proc.StartInfo.Arguments = string.Empty;
                                proc.StartInfo.WorkingDirectory = path;
                                proc.Start();
                                Thread.Sleep(5000);

                                
                          
                                if (System.IO.File.Exists(TotalPath))
                                    TotalPath = "exist";
                                else
                                    TotalPath = "NOTexist";
                                //StartProcessNoWait(path, "TFSQATestTools.exe", string.Empty);

                                //while (ps.Length < runTools + 1)
                                //{
                                //    Console.WriteLine("running: number is not OK  and wait 5 sec. -> num = " + ps.Length);
                                //    Thread.Sleep(2000);
                                //    ps = Process.GetProcessesByName("TFSQATestTools");
                                //}
                                //ps = System.Diagnostics.Process.GetProcessesByName("TFSQATestTools");
                                //Console.WriteLine("final running: number is OK  -> num = " + ps.Length);
                                WriteRemoteCommandReply("new testtools is startup:" + path);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                        else if (info.ToLower().StartsWith("forcekill"))
                        {
                            Console.WriteLine("Start command Received -------------- " + info);
                            try
                            {
                                System.Diagnostics.Process[] ps = System.Diagnostics.Process.GetProcessesByName("TFSQATestTools");
                                for (int i = 0; i < ps.Length; i++)
                                {
                                    Console.WriteLine(ps[i].Id + " is killed");
                                    ps[i].Kill();
                                }

                                System.Diagnostics.Process[] ps2 = System.Diagnostics.Process.GetProcessesByName("WerFault");
                                for (int i = 0; i < ps2.Length; i++)
                                {
                                    Console.WriteLine(ps2[i].Id + " is killed");
                                    ps2[i].Kill();
                                }
                                WriteRemoteCommandReply(" is forcekilled");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }
                        else
                        {
                            Console.WriteLine("other command  -------------- " + info);
                            //WriteRemoteCommandReply("hhhhhhhhhhhhhhhhhhh:"+ii++);
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    string remoteCommandExceptionFile = Path.Combine(@"C:\EtriccTests", "RemoteCommandException.txt");
                    StreamWriter writeExc = File.CreateText(remoteCommandExceptionFile);
                    writeExc.WriteLine("remoteCommandException:" + ex.Message + " --- " + ex.StackTrace);
                    writeExc.Close();
                }

            }
        }

        public static System.Diagnostics.Process StartProcessNoWait(string processDir, string procFilename, string args)
        {
            string path = Path.Combine(processDir, procFilename);
            var proc = new System.Diagnostics.Process();
            proc.EnableRaisingEvents = false;
            proc.StartInfo.FileName = path;
            proc.StartInfo.Arguments = args;
            proc.StartInfo.WorkingDirectory = processDir;
            proc.Start();
            Thread.Sleep(5000);
            return proc;
            //proc5.WaitForExit();
        }

        public static string ProgramFilesx86()
        {
            if (8 == IntPtr.Size ||
                (!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"))))
            {
                return Environment.GetEnvironmentVariable("ProgramFiles(x86)");
            }

            return Environment.GetEnvironmentVariable("ProgramFiles");
        }

        private static void WriteRemoteCommandReply(string message)
        {
            bool writeOK = false;
            while (writeOK == false)
            {
                try
                {
                    StreamWriter write2 = File.CreateText(Path.Combine(@"C:\EtriccTests", "RemoteCommand.txt"));
                    write2.WriteLine(message);
                    write2.Close();
                    writeOK = true;
                }
                catch (Exception ex)
                {
                    writeOK = false;
                    Thread.Sleep(5000);
                    Console.WriteLine("write " + message + " to RemoteCommand.txt exception: " + ex.Message);
                }
            }
        }
    }

    public partial class Service1 : ServiceBase
    {
        public Service1() 
        {
            this.ServiceName = "QATestsService";
            //ToolsWindowsServiceInstaller();/*InitializeComponent();*/ 
        }

        public void Start(string[] args) 
        {
            //System.Windows.Forms.MessageBox.Show("startzzzzzzzzzzzz\n");
            OnStart(args); 
        }
        
        protected override void OnStart(string[] args)
        {
            WriteRemoteCommandReply("hehe started");
            Task t = Task.Factory.StartNew(RunX);
            //RunX();
            
        }



        protected override void OnStop()
        {
            WriteRemoteCommandReply("hehe stopped");
        }

        static void RunX()
        {
            string remoteCommandFile = System.IO.Path.Combine(@"C:\EtriccTests", "RemoteCommand.txt");
            int ii = 0;
            while (true)
            {
                Console.WriteLine("start RUN ****************************\n");
                //System.Windows.Forms.MessageBox.Show("kk:: " + ii++);
                Thread.Sleep(10000);
                // Check test Working file
                System.IO.FileInfo workFile = new System.IO.FileInfo(remoteCommandFile);
                File.SetAttributes(workFile.FullName, FileAttributes.Normal);

                string info = string.Empty;
                bool ReadOK = false;
                while (ReadOK == false)
                {
                    try
                    {
                        StreamReader readerInfo = File.OpenText(remoteCommandFile);
                        info = readerInfo.ReadToEnd();
                        readerInfo.Close();
                        ReadOK = true;
                    }
                    catch (Exception ex)
                    {
                        ReadOK = false;
                        Thread.Sleep(5000);
                        Console.WriteLine("Read RemoteCommand.txt exception: " + ex.Message);
                    }
                }

                try
                {
                    #region process remote command
                    if (info.Length == 0)
                        Console.WriteLine(" --- Current command Empty -------------- ");
                    else
                    {
                        Console.WriteLine(" --- Current command Received -------------- " + info);

                        if (info.ToLower().StartsWith("forcestartup"))
                        {
                            Console.WriteLine("Start command Received -------------- " + info);
                            try
                            {
                                Console.WriteLine("number TFSQATestTools is running: = 000000000000000000000000000");
                                //System.Diagnostics.Process[] ps = System.Diagnostics.Process.GetProcessesByName("TFSQATestTools");
                                //int runTools = ps.Length;
                                //Console.WriteLine("number TFSQATestTools is running: = " + runTools);

                                string path = ProgramFilesx86() + @"\Egemin\AutomaticTesting";

                                string TotalPath = Path.Combine(path, "TFSQATestTools.exe");
                                var proc = new System.Diagnostics.Process();
                                /*proc.EnableRaisingEvents = false;
                                proc.StartInfo.FileName = "TFSQATestTools.bat";
                                proc.StartInfo.Arguments = string.Empty;
                                proc.StartInfo.WorkingDirectory = path;
                                proc.Start();
                                Thread.Sleep(5000);
                                */
                                String applicationName = TotalPath;

                                // launch the application
                                //Toolkit.ApplicationLoader.PROCESS_INFORMATION procInfo;
                                //Toolkit.ApplicationLoader.StartProcessAndBypassUAC(applicationName, out procInfo);
                          
                                if (System.IO.File.Exists(TotalPath))
                                    TotalPath = "exist";
                                else
                                    TotalPath = "NOTexist";
                                //StartProcessNoWait(path, "TFSQATestTools.exe", string.Empty);

                                //while (ps.Length < runTools + 1)
                                //{
                                //    Console.WriteLine("running: number is not OK  and wait 5 sec. -> num = " + ps.Length);
                                //    Thread.Sleep(2000);
                                //    ps = Process.GetProcessesByName("TFSQATestTools");
                                //}
                                //ps = System.Diagnostics.Process.GetProcessesByName("TFSQATestTools");
                                //Console.WriteLine("final running: number is OK  -> num = " + ps.Length);
                                WriteRemoteCommandReply("new testtools is startup:" + path);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                        else if (info.ToLower().StartsWith("forcekill"))
                        {
                            Console.WriteLine("Start command Received -------------- " + info);
                            try
                            {
                                System.Diagnostics.Process[] ps = System.Diagnostics.Process.GetProcessesByName("TFSQATestTools");
                                for (int i = 0; i < ps.Length; i++)
                                {
                                    Console.WriteLine(ps[i].Id + " is killed");
                                    ps[i].Kill();
                                }

                                System.Diagnostics.Process[] ps2 = System.Diagnostics.Process.GetProcessesByName("WerFault");
                                for (int i = 0; i < ps2.Length; i++)
                                {
                                    Console.WriteLine(ps2[i].Id + " is killed");
                                    ps2[i].Kill();
                                }
                                WriteRemoteCommandReply(" is forcekilled");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }
                        else
                        {
                            Console.WriteLine("other command  -------------- " + info);
                            //WriteRemoteCommandReply("hhhhhhhhhhhhhhhhhhh:"+ii++);
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    string remoteCommandExceptionFile = Path.Combine(@"C:\EtriccTests", "RemoteCommandException.txt");
                    StreamWriter writeExc = File.CreateText(remoteCommandExceptionFile);
                    writeExc.WriteLine("remoteCommandException:" + ex.Message + " --- " + ex.StackTrace);
                    writeExc.Close();
                }

            }
        }

        public static System.Diagnostics.Process StartProcessNoWait(string processDir, string procFilename, string args)
        {
            string path = Path.Combine(processDir, procFilename);
            var proc = new System.Diagnostics.Process();
            proc.EnableRaisingEvents = false;
            proc.StartInfo.FileName = path;
            proc.StartInfo.Arguments = args;
            proc.StartInfo.WorkingDirectory = processDir;
            proc.Start();
            Thread.Sleep(5000);
            return proc;
            //proc5.WaitForExit();
        }

        public static string ProgramFilesx86()
        {
            if (8 == IntPtr.Size ||
                (!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"))))
            {
                return Environment.GetEnvironmentVariable("ProgramFiles(x86)");
            }

            return Environment.GetEnvironmentVariable("ProgramFiles");
        }

        private static void WriteRemoteCommandReply(string message)
        {
            bool writeOK = false;
            while (writeOK == false)
            {
                try
                {
                    StreamWriter write2 = File.CreateText(Path.Combine(@"C:\EtriccTests", "RemoteCommand.txt"));
                    write2.WriteLine(message);
                    write2.Close();
                    writeOK = true;
                }
                catch (Exception ex)
                {
                    writeOK = false;
                    Thread.Sleep(5000);
                    Console.WriteLine("write " + message + " to RemoteCommand.txt exception: " + ex.Message);
                }
            }
        }
    }

}
