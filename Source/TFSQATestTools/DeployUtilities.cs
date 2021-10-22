using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using TestTools;
using Condition = System.Windows.Automation.Condition;

namespace TFSQATestTools
{
    public class DeployUtilities
    {
        [DllImport("mpr.dll")]
        public static extern int WNetAddConnection2A(
            [MarshalAs(UnmanagedType.LPArray)] NETRESOURCEA[] lpNetResource,
            [MarshalAs(UnmanagedType.LPStr)] string lpPassword,
            [MarshalAs(UnmanagedType.LPStr)] string UserName,
            int dwFlags);

        [DllImport("mpr.dll")]
        public static extern int WNetCancelConnection2A(string sharename, int dwFlags, int fForce);

        public static AutomationElement GetMainWindow(string mainFormId)
        {
            AutomationElement aeWindow = null;
            AutomationElementCollection aeAllWindows = null;
            // find main window
            Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            int k = 0;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while (aeWindow == null && mTime.TotalSeconds <= 120)
            {
                Console.WriteLine("aeWindow[k]=");
                k++;
                aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                Thread.Sleep(3000);
                for (int i = 0; i < aeAllWindows.Count; i++)
                {
                    if (aeAllWindows[i].Current.AutomationId.Equals(mainFormId))
                    {
                        aeWindow = aeAllWindows[i];
                        Console.WriteLine("aeWindow[" + i + "]=" + aeWindow.Current.Name);
                        break;
                    }
                }
                mTime = DateTime.Now - mStartTime;
            }

            return aeWindow;
        }

        public static string getThisPCOS()
        {
            string name = OSVersionInfoClass.OSVersionInfo.Name;
            var sb = new StringBuilder();
            for (int i = 0; i < name.Length; i++)
            {
                if (char.IsLetterOrDigit(name[i]))
                {
                    sb.Append(name[i]);
                }
            }

            string bit = OSVersionInfoClass.OSVersionInfo.OSBits.ToString();
            if (bit.IndexOf("32") >= 0)
                bit = "32";
            else
                bit = "64";
            return sb + "." + bit;
        }

        public static void StartExecution()
        {
            //System.Windows.Forms.MessageBox.Show("sss");
            Thread.Sleep(20000);
            AutomationElement aeWindow = GetMainWindow("ToolsForm");
            if (aeWindow != null)
            {
                string id = "btnStartAuto";
                AutomationElement aeAutoStartButton = AUIUtilities.FindElementByID(id, aeWindow);
                if (aeAutoStartButton != null)
                {
                    Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeAutoStartButton));
                }
            }
        }

        public static bool CopySetupFilesWithWildcards(string fromPath, string toPath, string filenameWithWildcards,
                                                       Logger logger /*, ref Tester.STATE m_State*/)
        {
            /*if (fromPath.StartsWith(@"\\"))
            {
                //if the first action fails try to logon to the server
                if (CreateDriveMap(fromPath) != 0)
                {
                    System.Windows.MessageBox.Show("CreateDriveMap2   failed:" + fromPath);
                    return false;
                }
            }*/

            if (!Directory.Exists(fromPath))
            {
                Directory.CreateDirectory(fromPath);
            }


            if (Directory.Exists(toPath))
            {
                var DirInfo = new DirectoryInfo(toPath);
                FileInfo[] FilesToDelete = DirInfo.GetFiles();

                foreach (FileInfo file in FilesToDelete)
                {
                    try
                    {
                        var attributes = FileAttributes.Normal;
                        File.SetAttributes(file.FullName, attributes);
                        file.Delete();
                    }
                    catch (Exception ex)
                    {
                        //if (m_Settings.EnableLog)
                        //{
                        //string logPath = Configuration.BuildInformationfilePath;
                        //string path = Path.Combine( logPath, Configuration.LogFilename );
                        //Logger logger = new Logger(path );
                        logger.LogMessageToFile("----------Setup Error  --------", 0, 0);
                        logger.LogMessageToFile("CopySetup Exception:" + ex, 0, 0);
                        //Log("CopySetup Exception:" + ex.ToString());
                        //}
                        MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
                        //m_State = Tester.STATE.EXCEPTION;
                        return false;
                    }
                }
            }
            else
                Directory.CreateDirectory(toPath);

            FileInfo[] FilesToCopy;
            try
            {
                var DirInfo = new DirectoryInfo(fromPath);
                FilesToCopy = DirInfo.GetFiles(filenameWithWildcards);

                foreach (FileInfo file in FilesToCopy)
                {
                    file.CopyTo(Path.Combine(toPath, file.Name));
                }
                //Log("Copied Setup from " + fromPath + " to " + toPath);
                try
                {
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    logger.LogMessageToFile("Copied Setup from " + fromPath + " to " + toPath, 0, 0);
                    //}
                }
                catch (Exception ex1)
                {
                    //if (m_Settings.EnableLog)
                    logger.LogMessageToFile("------ Test Exception : " + ex1.Message + "\r\n" + ex1.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                }
            }
            catch (Exception ex)
            {
                try
                {
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    logger.LogMessageToFile("----------Setup Error --------", 0, 0);
                    logger.LogMessageToFile("CopySetup Exception:" + ex, 0, 0);
                    //Log("CopySetup Exception:" + ex.ToString());
                    //}
                }
                catch (Exception ex2)
                {
                    //if (m_Settings.EnableLog)
                    logger.LogMessageToFile("------ Test Exception : " + ex2.Message + "\r\n" + ex2.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                }
                MessageBox.Show("FromPath=" + fromPath + "   " + ex + "\r\n" + ex.StackTrace);
                //m_State = Tester.STATE.EXCEPTION;
                return false;
            }
            return true;
        }

        /// <summary>
        /// Cancel a drive mapping to the destination
        /// </summary>
        /// <param name="Destination">Full drive path</param>
        public static int Disconnect(string localpath)
        {
            int result = WNetCancelConnection2A(localpath, 1, 1);
            return result;
        }

        /// <summary>
        /// Open a drive mapping to the destination
        /// </summary>
        /// <param name="Destination">Full drive path</param>
        public static int OpenDriveMap(string Destination, string driveLetter)
        {
            if ((Destination == null) || (Destination == ""))
                return -1;

            var netResource = new NETRESOURCEA[1];
            netResource[0] = new NETRESOURCEA();
            netResource[0].dwType = 1;
            netResource[0].lpLocalName = driveLetter;
            netResource[0].lpRemoteName = Destination;
            netResource[0].lpProvider = null;
            int dwFlags = 1; /*CONNECT_INTERACTIVE = 8|CONNECT_PROMPT = 16*/

            int result = WNetAddConnection2A(netResource, null, null, dwFlags);
            //int result = WNetAddConnection2A(netResource, null, @"teamsystems\tfstest", dwFlags);

            return result;
        }

        public static bool TestRunning()
        {
            Process proc = null;
            int ipidEpia = ProcessUtilities.GetApplicationProcessID("Egemin.Epia.Testing.UIAutoTest", out proc);
            int ipidEpiaResourceEditor =
                ProcessUtilities.GetApplicationProcessID("Egemin.Epia.Foundation.Globalization.ResourceFileEditor",
                                                         out proc);
            int ipidEtricc = ProcessUtilities.GetApplicationProcessID("Egemin.Epia.Testing.EtriccUIAutoTest", out proc);
            int ipidStatistics = ProcessUtilities.GetApplicationProcessID(
                "Egemin.Epia.Testing.EtriccStatisticsProgTest", out proc);
            int ipidEpiaProtected = ProcessUtilities.GetApplicationProcessID(
                "Egemin.Epia.Testing.Epia4AppTestProtected", out proc);
            int ipidEpiaServer = ProcessUtilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SERVER, out proc);
            int ipidExcel = ProcessUtilities.GetApplicationProcessID("EXCEL", out proc);
            int ipidMsiExec = ProcessUtilities.GetApplicationProcessID("msiexec", out proc);

            int run = ipidEtricc + ipidEpia + ipidStatistics + ipidEpiaProtected + ipidEpiaServer +
                      ipidEpiaResourceEditor + ipidExcel + ipidMsiExec;
            if (run > 0)
                return true;
            else
                return false;
        }

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        public static void CreateInitialTestInfoFile(string testApp, string testInfoTxtFile)
        {
            StreamWriter writeInfo = File.CreateText(testInfoTxtFile);

            writeInfo.WriteLine("TestInfo:" + "Initial time:" + DateTime.Now.ToLocalTime());
                // + System.Environment.NewLine);
            string defFileName = "Epia4TestDefinition.txt";
            string defPath = @"C:\EtriccTests\QA\TestDefinitions";

            if (testApp.Equals(TestApp.EPIA4))
            {
                defFileName = TestDefName.EPIA4;
            }
            else if (testApp.Equals(TestApp.ETRICCUI))
            {
                defFileName = TestDefName.ETRICCUI;
            }
            else if (testApp.Equals(TestApp.ETRICCSTATISTICS))
            {
                defFileName = TestDefName.ETRICCSTATISTICS;
            }
            else if (testApp.Equals(TestApp.EPIANET45))
            {
                defFileName = TestDefName.EPIANET45;
            }

            StreamReader reader = File.OpenText(Path.Combine(defPath, defFileName));
            string line = reader.ReadLine();
            while (line != null)
            {
                writeInfo.WriteLine(line);
                line = reader.ReadLine();
            }
            reader.Close();
            writeInfo.Close();
        }

        public static void AddUpdateStatusInTestInfoFile(string infoFile, string status, string message,
                                                         string infoFileKey, string infoFileKeyPC, Logger logger)
        {
            var AllLines = new StringCollection();
            var NewAllLines = new StringCollection();
            var FinalAllLines = new StringCollection();

            // Read all lines from test info file
            StreamReader reader = File.OpenText(infoFile);
            string infoline = reader.ReadLine();
            while (infoline != null)
            {
                AllLines.Add(infoline);
                infoline = reader.ReadLine();
            }
            reader.Close();

            // Update info file
            foreach (string line in AllLines)
            {
                if (line.StartsWith(infoFileKey))
                {
                    if (line.Length < (infoFileKey.Length + 1))
                    {
                        // PC Empty line removed here and new line with this infoFileKeyPC will be added in FinalAllLines if not exist
                        logger.LogMessageToFile(infoFileKey + "PC Empty line " + " exist, not add to file " + line, 0, 0);
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        NewAllLines.Add(line);
                    }
                }
                    //NewAllLines.Add(infoFileKey + "-" + status + ":" + message);
                else
                {
                    NewAllLines.Add(line);
                }
            }

            // Update info file
            bool hasRecord = false;
            foreach (string line in NewAllLines)
            {
                if (line.StartsWith(infoFileKeyPC))
                {
                    // not add this line to final list, already in testing, do nothing, status will be change later 
                    //FinalAllLines.Add(infoFileKey + "-" + status + ":" + message);
                    hasRecord = true;
                }
                else
                {
                    FinalAllLines.Add(line);
                }
            }

            // if not exist, add this record
            if (hasRecord == false)
                FinalAllLines.Add(infoFileKeyPC + "-" + status + ":" + message);

            StreamWriter write = File.CreateText(infoFile);
            foreach (string line in FinalAllLines)
            {
                write.WriteLine(line);
                logger.LogMessageToFile(" write testinfo file: content is " + line, 0, 0);
            }
            write.Close();
        }

        public static string[] GetTestPlatformArray(string testApp, string testInfoFile)
        {
            string[] x = null;
            string currentPlatform = string.Empty;
            var currentPlatformList = new List<string>();
            string key = getThisPCOS();
            StreamReader reader = File.OpenText(testInfoFile);
            string line = reader.ReadLine();
            while (line != null)
            {
                // Windows7.32.AnyCPU.Debug.ETRICCSTATAUTOT-GUI Tests Passed:Tests OK --> AnyCPU.Debug
                if (line.StartsWith(key))
                {
                    int beginPos = key.Length + 1;
                    int endPos = line.LastIndexOf(']');
                    currentPlatform = line.Substring(beginPos, endPos - beginPos);
                    currentPlatformList.Add(currentPlatform);
                    Console.WriteLine("---   currentPlatform in test info:" + currentPlatform);
                }

                line = reader.ReadLine();
            }
            reader.Close();

            x = currentPlatformList.ToArray();
            return x;
        }

        /// <summary>
        ///  get build nr of release version from hotfix version: 
        ///  example: hotfix : Hotfix 5.7.8.2 of Etricc.Production.Hotfix_20120822.1
        ///           return:  Release 5.7.8 of Etricc.Production.Release_20111116.1
        /// </summary>
        /// <param name="buildnr">hotfix build number</param>
        /// <returns></returns>
        public static string GetReleaseBuildNrFromThisHotfixBuild(string buildnr)
        {
            int ind = buildnr.IndexOf("of");
            string version = buildnr.Substring(7, ind - 8);

            int indLastComma = version.LastIndexOf('.');
            string releaseVersion = version.Substring(0, indLastComma);
            return releaseVersion;
        }

        public static DateTime GetDateCompletedOfThisBuild(string project, string buildDef, string thisBuildNr)
        {
            DateTime datetime = DateTime.Today;
            TfsTeamProjectCollection tfsProjectCollection;

            string sTFSServerUrl = "http://team2010App.teamSystems.egemin.be:8080/tfs/Development";
            string sTFSUsername = "TfsBuild";
            string sTFSPassword = "Egemin01";
            string sTFSDomain = "TeamSystems.Egemin.Be";
            var serverUri = new Uri(sTFSServerUrl);
            ICredentials tfsCredentials
                = new NetworkCredential(sTFSUsername, sTFSPassword, sTFSDomain);

            tfsProjectCollection
                = new TfsTeamProjectCollection(serverUri, tfsCredentials);

            IBuildServer buildServer;
            tfsProjectCollection.EnsureAuthenticated();
            buildServer = (IBuildServer) tfsProjectCollection.GetService(typeof (IBuildServer));

            IBuildDetailSpec buildDetailSpec = buildServer.CreateBuildDetailSpec(project, buildDef);
            //buildDetailSpec.MaxBuildsPerDefinition = 1; 
            buildDetailSpec.QueryOrder = BuildQueryOrder.FinishTimeDescending;
            buildDetailSpec.Status = BuildStatus.Succeeded; //Only get succeeded builds  
            //buildDetailSpec.MinFinishTime = timeFrom;
            IBuildQueryResult results = buildServer.QueryBuilds(buildDetailSpec);
            //if (results.Failures.Length == 0 ) 
            //{ 
            //IBuildDetail buildDetail = results.Builds[0]; 
            //Console.WriteLine("Build: " + buildDetail.BuildNumber); 
            //Console.WriteLine("Account requesting build “ + 
            //“(build service user for triggered builds): " + buildDetail.RequestedBy); 
            //   Console.WriteLine("Build triggered by: " + buildDetail.RequestedFor); 
            //}

            IBuildDetail[] buildnrs = results.Builds;
            //IBuildDetail[] buildnrs = buildServer.QueryBuilds(selectedProject, s);
            string bnrs = string.Empty;
            string quality = string.Empty;
            //BuildObject thisBuild = new BuildObject();

            for (int i = 0; i < buildnrs.Length; i++)
            {
                if (buildnrs[i].BuildNumber.IndexOf(thisBuildNr) >= 0)
                {
                    datetime = buildnrs[i].FinishTime;
                    //System.Windows.Forms.MessageBox.Show("release version: " + buildnrs[i].BuildNumber, "count:" + buildnrs.Length);
                    //System.Windows.Forms.MessageBox.Show("date: " + buildnrs[i].FinishTime, "count:" + buildnrs.Length);
                    break;
                }
            }

            return datetime;
        }

        static public DateTime getCurrentBuildCompleteDate(string sBuildNr, ref DateTime releaseVersionDate, string teamProject, string buildDef, ref string sErrorMessage)
        {
            DateTime thisBuildDate = DateTime.MinValue;
            releaseVersionDate = DateTime.MinValue;
            thisBuildDate = DeployUtilities.GetDateCompletedOfThisBuild(teamProject, buildDef, sBuildNr);
            if (sBuildNr.IndexOf("Hotfix") >= 0) //string buildnr = "Hotfix 4.3.2.1 of Epia.Production.Hotfix_20120405.1";
            {
                //Release 4.4.4 of Epia.Production.Release_20120731.1
                string releaseBuildNr = DeployUtilities.GetReleaseBuildNrFromThisHotfixBuild(sBuildNr);
                releaseVersionDate = DeployUtilities.GetDateCompletedOfThisBuild(teamProject, buildDef, releaseBuildNr);
            }
            else if (sBuildNr.IndexOf("Production.Release") >= 0)
            {
                releaseVersionDate = DeployUtilities.GetDateCompletedOfThisBuild(teamProject, buildDef, sBuildNr);
            }
            return thisBuildDate;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="stepMsg"></param>
        /// <param name="App">Epia4, EtriccShell</param>
        /// <param name="logger"></param>
        /// <returns></returns>
        public static bool InstallApplicationSetupByStep(string App, int step, string errorMsg, Logger logger)
        {
            // example of window screen text: Epia
            //string[] SetupStepDescriptions = new string[100];
            //SetupStepDescriptions[0] = "Welcome";
            //SetupStepDescriptions[1] = "Welcome to the E'pia Framework 2010.05.11.1 Setup Wizard";
            //SetupStepDescriptions[2] = "Components";
            //SetupStepDescriptions[3] = "Installation Folders";
            //SetupStepDescriptions[4] = "Confirm Installation";
            //SetupStepDescriptions[5] = "Installing E'pia Framework ...";
            //SetupStepDescriptions[6] = "E'pia Framework 2010.12.22.* Information";
            //SetupStepDescriptions[7] = "Installation Complete";
            bool clickCloseButton = false;
            AutomationElement aeWindow = null;
            AutomationElementCollection aeAllWindows = null;
            AutomationElement aeClickButton = null;

            Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            DateTime StartTime = DateTime.Now;
            TimeSpan Time = DateTime.Now - StartTime;
            logger.LogMessageToFile("<-----> Start Install : " + App + "  Step by step, :" + step, 0, 0);
            try
            {
                // //find install application Window
                while (aeWindow == null && Time.TotalMinutes <= 2)
                {
                    //find all Windows
                    aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                    logger.LogMessageToFile(
                        "<-----> Start2 Install : " + App + "  Step by step, aeAllWindows.Count: " + aeAllWindows.Count,
                        0, 0);
                    Thread.Sleep(500);
                    try
                    {
                        for (int i = 0; i < aeAllWindows.Count; i++)
                        {
                            logger.LogMessageToFile(
                                "<----->Window title(" + i + ")  : " + aeAllWindows[i].Current.Name, 0, 0);
                            if (aeAllWindows[i].Current.Name.StartsWith("E'pia Framework")
                                || aeAllWindows[i].Current.Name.StartsWith("E'tricc Shell"))
                            {
                                aeWindow = aeAllWindows[i];
                                logger.LogMessageToFile("<----->Window title(" + i + ")  : " + aeWindow.Current.Name, 0,
                                                        0);
                                // Make sure our window is usable.
                                // WaitForInputIdle will return before the specified time 
                                // if the window is ready.
                                //WindowPattern windowPattern = aeWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                                //if (false == windowPattern.WaitForInputIdle(30000))
                                //{
                                //    System.Windows.Forms.MessageBox.Show("Object not responding in a timely manner, click OK continue", stepMsg);
                                //}
                                break;
                            }
                        }

                        if (aeWindow != null)
                        {
                            var wp = aeWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                            if (wp.Current.WindowVisualState == WindowVisualState.Minimized)
                            {
                                //System.Windows.Forms.MessageBox.Show("wp.Current.WindowVisualState == WindowVisualState.Minimized", stepMsg);
                                wp.SetWindowVisualState(WindowVisualState.Normal);
                                Thread.Sleep(1000);
                            }

                            #region // process Text

                            logger.LogMessageToFile("<-----> find aeWindow name is:" + aeWindow.Current.Name, 0, 0);
                            //find all install Window screen texts
                            Condition cText = new PropertyCondition(
                                AutomationElement.ControlTypeProperty, ControlType.Text);

                            AutomationElementCollection aeAllTexts = aeWindow.FindAll(TreeScope.Children, cText);
                            Thread.Sleep(3000);
                            for (int i = 0; i < aeAllTexts.Count; i++)
                            {
                                logger.LogMessageToFile(
                                    "<----->Window text(" + i + ")  : " + aeAllTexts[i].Current.Name, 0, 0);
                                if (aeAllTexts[i].Current.Name.StartsWith("Components"))
                                {
                                    switch (App)
                                    {
                                        case EgeminApplication.EPIA:

                                            #region

                                            AutomationElement aeIAgreeRadioButton
                                                = AUIUtilities.FindElementByName("E'pia Server", aeWindow);
                                            var tg =
                                                aeIAgreeRadioButton.GetCurrentPattern(TogglePattern.Pattern) as
                                                TogglePattern;
                                            ToggleState tgState = tg.Current.ToggleState;
                                            if (tgState == ToggleState.Off)
                                                tg.Toggle();

                                            AutomationElement aeShellckb
                                                = AUIUtilities.FindElementByName("E'pia Shell", aeWindow);
                                            var tg2 =
                                                aeShellckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                            ToggleState tg2State = tg2.Current.ToggleState;
                                            if (tg2State == ToggleState.Off)
                                                tg2.Toggle();

                                            logger.LogMessageToFile(
                                                "<---> Components Sever and shell checkbox state set on  ... ", 0, 0);
                                            Thread.Sleep(3000);

                                            #endregion

                                            break;
                                        case EgeminApplication.ETRICC_SHELL:

                                            #region

                                            AutomationElement aeServerckb
                                                =
                                                AUIUtilities.FindElementByName(
                                                    "E'pia Server Extensions (Resource & Security Files)", aeWindow);
                                            var tgShell =
                                                aeServerckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                            ToggleState tgShellState = tgShell.Current.ToggleState;
                                            if (tgShellState == ToggleState.Off)
                                                tgShell.Toggle();

                                            AutomationElement aeEtriccShellckb
                                                =
                                                AUIUtilities.FindElementByName(
                                                    "E'pia Shell Extensions (Shell Module & Config)", aeWindow);
                                            var tgEtricc =
                                                aeEtriccShellckb.GetCurrentPattern(TogglePattern.Pattern) as
                                                TogglePattern;
                                            ToggleState tg2ShellState = tgEtricc.Current.ToggleState;
                                            if (tg2ShellState == ToggleState.Off)
                                                tgEtricc.Toggle();

                                            AutomationElement aeEtricccCorekb
                                                = AUIUtilities.FindElementByName("E'tricc Core Extensions (Wrappers)",
                                                                                 aeWindow);
                                            var tg3 =
                                                aeEtricccCorekb.GetCurrentPattern(TogglePattern.Pattern) as
                                                TogglePattern;
                                            ToggleState tg3State = tg3.Current.ToggleState;
                                            if (tg3State == ToggleState.Off)
                                                tg3.Toggle();

                                            logger.LogMessageToFile(
                                                "<---> Components Sever and shell checkbox state set on  ... ", 0, 0);
                                            Thread.Sleep(3000);

                                            #endregion

                                            break;
                                    }
                                }
                            }

                            #endregion

                            Condition cButton = new AndCondition(
                                new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                                new PropertyCondition(AutomationElement.IsContentElementProperty, true),
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                                );

                            AutomationElementCollection aeButtons = aeWindow.FindAll(TreeScope.Children, cButton);
                            Thread.Sleep(1000);
                            bool clickButtonFound = false;
                            for (int i = 0; i < aeButtons.Count; i++)
                            {
                                logger.LogMessageToFile(
                                    "<----->Window enabled button(" + i + ")  : " + aeButtons[i].Current.Name, 0, 0);
                                if (aeButtons[i].Current.Name.StartsWith("Next"))
                                {
                                    aeClickButton = aeButtons[i];
                                    clickButtonFound = true;
                                    break;
                                }
                                else if (aeButtons[i].Current.Name.StartsWith("Close"))
                                {
                                    aeClickButton = aeButtons[i];
                                    clickButtonFound = true;
                                    clickCloseButton = true;
                                    break;
                                }
                            }

                            if (clickButtonFound)
                            {
                                logger.LogMessageToFile("<---> clickButtonFound... " + aeClickButton.Current.Name, 0, 0);
                                Point OptionPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                Input.MoveTo(OptionPt);
                                Thread.Sleep(1000);
                                logger.LogMessageToFile(
                                    "<---> " + aeClickButton.Current.Name + " button clicking ... ", 0, 0);
                                Input.ClickAtPoint(OptionPt);
                                Thread.Sleep(500);
                            }
                            else
                                Thread.Sleep(2000);
                        }
                        else
                        {
                            errorMsg = "Error: install window not found:" + App;
                        }
                    }
                    catch (ElementNotAvailableException ex)
                    {
                        string msg = ex + "----" + Environment.NewLine + ex.StackTrace;
                        logger.LogMessageToFile("<---> " + msg, 0, 0);
                        aeWindow = null;
                    }
                    Time = DateTime.Now - StartTime;
                }
            }
            catch (ElementNotAvailableException ex)
            {
                string msg = ex + "----" + Environment.NewLine + ex.StackTrace;
                logger.LogMessageToFile("<---> " + msg, 0, 0);
                clickCloseButton = false;
            }

            return clickCloseButton;
        }

        public static string getFullPathFromFilename(string filename, string topDirectory, string errorMsg)
        {
            string path = null;
            ;
            string[] FileList = Directory.GetFiles(topDirectory, filename, SearchOption.AllDirectories);

            if (FileList.Length > 0)
                path = FileList[0];

            return path;
        }

        public static bool IsThisOSPlatformTested(string testInfoTxtFile, Logger logger, string testApp, string platform)
        {
            bool tested = false;
            string testKey = string.Empty;
            string infoline = string.Empty;
            try
            {
                string thisPCOS = getThisPCOS();
                testKey = thisPCOS + "[" + platform + "]";
                Console.WriteLine("IsThisOSPlatformTested: OSPlatformKey is-->" + testKey);

                StreamReader readerInfo = File.OpenText(testInfoTxtFile);
                infoline = readerInfo.ReadLine();
                while (infoline != null)
                {
                    Console.WriteLine("IsThisOSPlatformTested: infoline-->" + infoline);
                    if (infoline.StartsWith(testKey))
                    {
                        Console.WriteLine("OK: infoline-->" + infoline);
                        if (infoline.Length > testKey.Length + 5 && infoline.IndexOf('-') > 0)  // Windows7.32[x86Debug]EPIAAUTOTEST1-GUI... 
                        {
                            //Log("but "+testPC + " is already in test info file");
                            logger.LogMessageToFile(testKey + " is already in test info file", 0, 0);
                            Console.WriteLine("OK: Tested-->");
                            tested = true;
                            break;
                        }
                    }
                    infoline = readerInfo.ReadLine();
                }

                readerInfo.Close();

                if (tested == false)
                {
                    //Log(testKey + " ---------->> is not in test info file");
                    logger.LogMessageToFile(testKey + " ----------> is not in test info file", 0, 0);
                }
                Console.WriteLine("return tested-->" + tested);
                return tested;
            }
            catch (Exception ex)
            {
                //Log("IsTestWorking exception:" + testKey + " - message:" + ex.Message + " --- " + ex.StackTrace);
                logger.LogMessageToFile(
                    "IsThisOSPlatformTested exception:" + testKey + "+" + testApp + "+" + platform + "=" + " - message:" +
                    ex.Message + " --- " + ex.StackTrace,
                    0, 0);
                return tested;
            }
        }

        public static string GetTestReportContentString(int sTotalCounter, int sTotalPassed, int sTotalFailed, int sTotalException, int sTotalUntested, 
            string sCurrentPlatform, string sInstallMsiDir)
        {
            string PCName = System.Environment.MachineName;

            string strHostName = Dns.GetHostName();
            IPHostEntry ipEntry = Dns.GetHostEntry(strHostName);
            System.Net.IPAddress[] addr = ipEntry.AddressList;
            string IPs = string.Empty;
            for (int i = 0; i < addr.Length; i++)
            {
                if (addr[i].AddressFamily.ToString() == System.Net.Sockets.AddressFamily.InterNetwork.ToString())
                {
                    IPs = addr[i].ToString() + " " + IPs; 
                    Console.WriteLine("IP Address {0}: {1} ", i, addr[i].ToString());
                }
            }
            
            string config = sInstallMsiDir.Substring(sInstallMsiDir.LastIndexOf(@"\") );
            string str = "<html><body><b><center>Test Overview</center></b><br>" + "<br>" +
                        "<table col=" +
                        '"' + "5" + '"' + " > <tr><th></th><th>Total Tests:</th><th>&nbsp;</th><th>"
                        + sTotalCounter
                        + "</th>  <th></th> </tr>"
                        + "<tr><td><br></td><td></td><td></td><td></td><td></td></tr>"
                        + "                            <tr><td></td>	<td>Pass:     </td> <td></td>       <td>"
                        + sTotalPassed
                        + "</td><td></td><td></td></tr><tr><td></td>	<td>Fail:     </td>  <td></td>	    <td>"
                        + sTotalFailed
                        + "</td><td></td><td></td></tr><tr><td></td>	<td>Exception:</td>  <td></td>	    <td>"
                        + sTotalException
                        + "</td><td></td><td></td></tr><tr><td></td>	<td>Untested:</td>   <td></td>	    <td>"
                        + sTotalUntested
                        + "</td><td></td>	<td></td></tr></table><br>"
                        + DeployUtilities.getThisPCOS() + "<br>"
                        + PCName +  "  --- IP address:  "+ IPs+"<br>"
                        + sCurrentPlatform + "<br>"
                        + config + "<br></body></html>";


            return str;
        }


        #region Nested type: NETRESOURCEA

        /// <summary>
        /// network related struct
        /// </summary>
        public struct NETRESOURCEA
        {
            public int dwType;
            public int dwUsage;
            [MarshalAs(UnmanagedType.LPStr)] public string lpComment;
            [MarshalAs(UnmanagedType.LPStr)] public string lpLocalName;
            [MarshalAs(UnmanagedType.LPStr)] public string lpProvider;
            [MarshalAs(UnmanagedType.LPStr)] public string lpRemoteName;

            public override String ToString()
            {
                String str = "LocalName: " + lpLocalName + " RemoteName: " + lpRemoteName
                             + " Comment: " + lpComment + " lpProvider: " + lpProvider;
                return (str);
            }
        }

        #endregion
    }
}