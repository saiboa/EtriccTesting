using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Xml;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using System.Collections.Specialized;
using System.Reflection;
using TestTools;

namespace TFSQATestTools
{
    public partial class ToolsForm : Form
    {
        List<string> xBuildDefs = new List<string>();
        #region Fields of main form
        //public static DispatcherTimer dispatcherTimer = new DispatcherTimer();
        public static DateTime sStartUpTime;
        public Tester tester = new Tester();
        public string TFSQAtestToolVersion = "4.12.11.30";
        internal static TestTools.Logger logger = null;
        internal static TimeSpan sTickTime = new TimeSpan(0, 0, 15);
        internal static TimeSpan sTickTime2Min = new TimeSpan(0, 2, 59);
        public static int sTickCount = 0;
        public static int sTickInterval = 60000;
        public static  List<string> sSelectedBuildDefs = null;
        public static string sDateFilter = "Today";
        public static bool autoStartup = false;
        public static bool autoPressStartButton = false;

        public static bool sVMSwitchMode = false;
        public static int sCounter = 0;

        public static string[] sEpiaTestDefinitions;
        public static string[] sEtriccUITestDefinitions;
        public static string[] sStatisticsTestDefinitions;
        public static string sCommandReplyFile = Path.Combine(@"C:\EtriccTests", "RemoteReply.txt");
        public static string sTestDefinitionFilesPath = System.Configuration.ConfigurationManager.AppSettings.Get("TestDefinitionFilePath");
        public static string sQAFolder = string.Empty;
        public static Task MyTask = null;

        static List<DataObject> list = new List<DataObject>();
        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        public ToolsForm()
        {
            try
            {
                #region
                InitializeComponent();

                cmbTestApp1DefFileName.Text = TestDefName.EPIA4;
                cmbTestApp2DefFileName.Text = TestDefName.ETRICCUI;
                cmbTestApp3DefFileName.Text = TestDefName.ETRICCSTATISTICS;

                sStartUpTime = System.DateTime.Now;
                sTickInterval = Convert.ToInt32(Constants.STimerInterval);
                lbStartTime.Text = sStartUpTime.ToString("yyyyMMdd HH:mm:ss");
                tester.sTestStartUpTime = sStartUpTime;
                //tester.m_TestWorkingDirectory = System.IO.Directory.GetCurrentDirectory();

                FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                TFSQAtestToolVersion = fileVersionInfo.FileVersion.ToString();

                tester.TESTTOOL_VERSION = TFSQAtestToolVersion;
                tester.sTestDefinitionFilesPath = sTestDefinitionFilesPath;
  
                lbSelectedConfigFile.Text = "Start config test settings .....";
              
                string sLogFilePath = tester.getLogPath();
                logger = new TestTools.Logger(sLogFilePath);
                logger.LogMessageToFile("Start Log: " + sLogFilePath, 0, 0);

                //dispatcherTimer.Enabled = true;
                tester.OnLoggingChanged += new EventHandler(tester_OnLoggingChanged);

           
                // Timer for WPF Window
                //DispatcherTimer dispatcherTimer = new DispatcherTimer();
                //dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
                //dispatcherTimer.Interval = sTickTime;
                //dispatcherTimer.Start();
                /*
                int ret = DeployUtilities.Disconnect(Constants.TEST_DEFINITION_DRIVE_MAP_LETTER);
                if (ret == 0)
                {
                    logger.LogMessageToFile("Disconnect MAP DRIVE OK:", 0, 0);
                }
                else if (ret == 2250)
                    logger.LogMessageToFile(
                        "Disconnnet: MAP DRIVE The Network connection could not be found :" + ret,
                        0, 0);
                else
                    System.Windows.MessageBox.Show("Disconnect  DriveMap failed with error code:" + ret);

                Thread.Sleep(3000);

                ret = DeployUtilities.OpenDriveMap(@"\\Teamsystem.Teamsystems.egemin.be\Team Systems Builds", Constants.TEST_DEFINITION_DRIVE_MAP_LETTER);
                if (ret == 0)
                {
                    logger.LogMessageToFile("Open MAP DRIVE OK:", 0, 0);
                }
                else if (ret == 85)
                    logger.LogMessageToFile("Open MAP DRIVE not connected due to existing connection:", 0, 0);
                else
                    System.Windows.MessageBox.Show("OpenDriveMap failed with error code:" + ret);
                */
                //sTestDefinitionFilesPath = @"C:\EtriccTests\QA\TestDefinitions";
                sQAFolder = sTestDefinitionFilesPath.Substring(0, sTestDefinitionFilesPath.LastIndexOf('\\'));
                TfsTeamProjectCollection tfsProjectCollection = null;
                DirectoryInfo dirInfo = new DirectoryInfo(sTestDefinitionFilesPath);
                while (System.IO.Directory.Exists(sTestDefinitionFilesPath))
                {
                    FileManipulation.DeleteRecursiveFolder(dirInfo);
                    Thread.Sleep(2000);
                }

                if (System.IO.Directory.Exists(sTestDefinitionFilesPath))
                {
                    string sErrorMessage = "QA Definition folder not deleted ";
                    Console.WriteLine(sErrorMessage);
                    System.Windows.Forms.MessageBox.Show(sErrorMessage, "start message ");
                }

                string msgX = "Get QA Definition Files";
                bool TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX, ref tfsProjectCollection);
                while (TFSConnected == false)
                {
                    DialogResult dr = MessageBox.Show("Are you want continue to connect again?", "TFS connection failed", MessageBoxButtons.YesNo);
                    if (dr == DialogResult.Yes)
                    {
                        TestTools.MessageBoxEx.Show(msgX + "\nWill try to reconnect the Server ...",
                                "Get QA Definition Files", (uint)Tfs.ReconnectDelay);
                        System.Threading.Thread.Sleep(Tfs.ReconnectDelay);
                        TFSConnected = TfsUtilities.CheckTFSConnection(ref msgX, ref tfsProjectCollection);
                    }
                    else
                    {
                        break;
                    }
                }

                if (TFSConnected)
                {
                    if (TfsUtilities.GetTestProjectQA(tfsProjectCollection, sQAFolder, ref msgX) == false)
                    {
                        System.Windows.Forms.MessageBox.Show("" + msgX, "Get test definition files from TFS failed");
                        Console.WriteLine(msgX);
                        // create local test definition files
                        #region
                        try
                        {
                            string Epia4TestDefinitionFile = Path.Combine(sTestDefinitionFilesPath, TestDefName.EPIA4);
                            StreamWriter writeInfo = File.CreateText(Epia4TestDefinitionFile);
                            writeInfo.WriteLine("Windows7.32.AnyCPU.Debug");
                            writeInfo.WriteLine("Windows7.32.AnyCPU.Release");
                            writeInfo.WriteLine("Windows7.32.AnyCPU.Protected");
                            writeInfo.WriteLine("Windows7.32.x86.Debug");
                            writeInfo.WriteLine("Windows7.32.x86.Release");
                            writeInfo.WriteLine("Windows7.32.x86.Protected");
                            writeInfo.WriteLine("Windows7.64.AnyCPU.Debug");
                            writeInfo.WriteLine("Windows7.64.AnyCPU.Release");
                            writeInfo.WriteLine("Windows7.64.AnyCPU.Protected");
                            writeInfo.WriteLine("WindowsServer2008R2.64.AnyCPU.Debug");
                            writeInfo.WriteLine("WindowsServer2008R2.64.AnyCPU.Release");
                            writeInfo.Close();

                            string EtriccUITestTypeDefinitionFile = Path.Combine(sTestDefinitionFilesPath, TestDefName.ETRICCUI);
                            writeInfo = File.CreateText(EtriccUITestTypeDefinitionFile);
                            writeInfo.WriteLine("Windows7.32.x86.Debug");
                            writeInfo.WriteLine("Windows7.32.AnyCPU.Debug");
                            writeInfo.WriteLine("Windows7.32.AnyCPU.Release");
                            writeInfo.WriteLine("Windows7.64.AnyCPU.Debug");
                            writeInfo.WriteLine("WindowsServer2008R2.64.AnyCPU.Debug");
                            writeInfo.Close();

                            string StatisticsTestTypeDefinitionFile = Path.Combine(sTestDefinitionFilesPath, TestDefName.ETRICCSTATISTICS);
                            writeInfo = File.CreateText(StatisticsTestTypeDefinitionFile);
                            writeInfo.WriteLine("Windows7.32.x86.Debug");
                            writeInfo.WriteLine("Windows7.32.AnyCPU.Debug");
                            writeInfo.Close();
                            //string zipFile = EticcTests.zip;
                            //string zipFile = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "QA.zip");
                            //FastZip fz = new FastZip();
                            //fz.ExtractZip(zipFile, OSVersionInfoClass.ProgramFilesx86() + @"\Dematic\AutomaticTesting", "");
                        }
                        catch (Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show("QA.zip unzip exception: " + ex.Message + "--------------" + ex.StackTrace);
                        }
                        Thread.Sleep(5000);
                        #endregion
                    }
                }
                else
                {
                    // create local test definition files
                    #region
                    try
                    {
                        string Epia4TestDefinitionFile = Path.Combine(sTestDefinitionFilesPath, "Epia4TestDefinition.txt");
                        if (!System.IO.Directory.Exists(sTestDefinitionFilesPath))
                            System.IO.Directory.CreateDirectory(sTestDefinitionFilesPath);

                        StreamWriter writeInfo = File.CreateText(Epia4TestDefinitionFile);
                        writeInfo.WriteLine("Windows7.32.AnyCPU.Debug");
                        writeInfo.WriteLine("Windows7.32.AnyCPU.Release");
                        writeInfo.WriteLine("Windows7.32.AnyCPU.Protected");
                        writeInfo.WriteLine("Windows7.32.x86.Debug");
                        writeInfo.WriteLine("Windows7.32.x86.Release");
                        writeInfo.WriteLine("Windows7.32.x86.Protected");
                        writeInfo.WriteLine("Windows7.64.AnyCPU.Debug");
                        writeInfo.WriteLine("Windows7.64.AnyCPU.Release");
                        writeInfo.WriteLine("Windows7.64.AnyCPU.Protected");
                        writeInfo.WriteLine("WindowsServer2008R2.64.AnyCPU.Debug");
                        writeInfo.WriteLine("WindowsServer2008R2.64.AnyCPU.Release");
                        writeInfo.Close();

                        string EtriccUITestTypeDefinitionFile = Path.Combine(sTestDefinitionFilesPath, "EtriccUITestTypeDefinition.txt");
                        writeInfo = File.CreateText(EtriccUITestTypeDefinitionFile);
                        writeInfo.WriteLine("Windows7.32.x86.Debug");
                        writeInfo.WriteLine("Windows7.32.AnyCPU.Debug");
                        writeInfo.WriteLine("Windows7.32.AnyCPU.Release");
                        writeInfo.WriteLine("Windows7.64.AnyCPU.Debug");
                        writeInfo.WriteLine("WindowsServer2008R2.64.AnyCPU.Debug");
                        writeInfo.Close();

                        string StatisticsTestTypeDefinitionFile = Path.Combine(sTestDefinitionFilesPath, "StatisticsTestTypeDefinition.txt");
                        writeInfo = File.CreateText(StatisticsTestTypeDefinitionFile);
                        writeInfo.WriteLine("Windows7.32.x86.Debug");
                        writeInfo.WriteLine("Windows7.32.AnyCPU.Debug");
                        writeInfo.Close();
                        //string zipFile = EticcTests.zip;
                        //string zipFile = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "QA.zip");
                        //FastZip fz = new FastZip();
                        //fz.ExtractZip(zipFile, OSVersionInfoClass.ProgramFilesx86() + @"\Dematic\AutomaticTesting", "");
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show("Create default test definition files exception: " + ex.Message + "--------------" + ex.StackTrace);
                    }
                    Thread.Sleep(5000);
                    #endregion
                }

                //DirectoryInfo DirInfo = new DirectoryInfo(@"M:\QA\TestDefinitions");
                DirectoryInfo DirInfo = new DirectoryInfo(sTestDefinitionFilesPath);
                FileInfo[] FilesOfTestDefinition = DirInfo.GetFiles("Epia*");
                sEpiaTestDefinitions = new string[FilesOfTestDefinition.Length]; 
                int i = 0;
                foreach (FileInfo file in FilesOfTestDefinition)
                    sEpiaTestDefinitions[i++] = file.Name;

                FilesOfTestDefinition = DirInfo.GetFiles("EtriccUI*");
                sEtriccUITestDefinitions = new string[FilesOfTestDefinition.Length]; 
                i = 0;
                foreach (FileInfo file in FilesOfTestDefinition)
                    sEtriccUITestDefinitions[i++] = file.Name;

                FilesOfTestDefinition = DirInfo.GetFiles("Statistics*");
                sStatisticsTestDefinitions = new string[FilesOfTestDefinition.Length]; 
                i = 0;
                foreach (FileInfo file in FilesOfTestDefinition)
                    sStatisticsTestDefinitions[i++] = file.Name;
            
                // initial Tfs settings
                if (System.Configuration.ConfigurationManager.AppSettings.Get("AutoPressStartButton").ToLower().Equals("true"))
                    autoPressStartButton = true;
            
                if (System.Configuration.ConfigurationManager.AppSettings.Get("AutoStartup").ToLower().Equals("true"))
                    autoStartup = true;

                SaveRunStartup(autoStartup);  // this will add register to auto startup this App after next system restart

                if (autoStartup)
                { 
                     System.Configuration.Configuration config =
                       System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                    string sectionName = Constants.TfsSettingsSection;
                    TfsSettingsSection customSection = (TfsSettingsSection)config.Sections[sectionName];
                    if (customSection != null)
                    {   // fill UI from TestSection 
                        sDateFilter = customSection.Element.DateFilter;
                        // new fields -------------------------------------------------
                        cmbTestApp1.Text = customSection.Element.TestApp1;
                        cmbTestApp1DefFileName.Text = customSection.Element.App1TestDefName;
                        cmbTestApp2.Text = customSection.Element.TestApp2;
                        cmbTestApp2DefFileName.Text = customSection.Element.App2TestDefName;
                        cmbTestApp3.Text = customSection.Element.TestApp3;

                        //System.Windows.Forms.MessageBox.Show("cmbTestApp3.Text:" + cmbTestApp3.Text, "Length:" + cmbTestApp3.Text.Length);

                        cmbTestApp3DefFileName.Text = customSection.Element.App3TestDefName;
                        if (cmbTestApp2.Text.Length == 0)
                            cmbTestApp2DefFileName.Text = "";
                        if (cmbTestApp3.Text.Length == 0)
                            cmbTestApp3DefFileName.Text = "";
                        // end new fields ----------------------------------------------

                        string allBuildDefsString = customSection.Element.BuildDefinitions;
                        string[] strArray = allBuildDefsString.Split(';');
                        ArrayList arList = new ArrayList();
                        for (int ix = 0; ix < strArray.Length; ix++)
                        {
                            if (strArray[ix].Length > 0)
                                lstBoxBuildDefinitions.Items.Add(strArray[ix]);
                        }

                        if (cmbTestApp1.Text.Length > 4)
                            cmbTestApp1DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp1.Text);

                        if (cmbTestApp2.Text.Length > 4)
                            cmbTestApp2DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp2.Text);

                        if (cmbTestApp3.Text.Length > 4)
                            cmbTestApp3DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp3.Text);
                    }
                    else
                    {
                        if (cmbTestApp1.Text.Length > 4)
                            cmbTestApp1DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp1.Text);

                        if (cmbTestApp2.Text.Length > 4)
                            cmbTestApp2DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp2.Text);

                        if (cmbTestApp3.Text.Length > 4)
                            cmbTestApp3DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp3.Text);
                
                    }
                }

                VMSwitchModeTimer.Enabled = true;
                tester.Log("VMSwitchModeTimer.Enabled...");
                VMSwitchModeTimer.Start();
                tester.Log("VMSwitchModeTimer.Started...");

                try
                {
                    StreamWriter writeInfo = File.CreateText(Path.Combine(@"C:\EtriccTests", "RemoteCommand.txt"));
                    writeInfo.WriteLine(string.Empty);
                    writeInfo.Close();

                    writeInfo = File.CreateText(Path.Combine(@"C:\EtriccTests", "RemoteReply.txt"));
                    writeInfo.WriteLine("TestRunning");
                    writeInfo.WriteLine(string.Empty);
                    writeInfo.Close();
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Create RemoteCommand.txt exception: " + ex.Message);
                }


                #endregion
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Start exception: " + ex.Message);
            }

            KeyPreview = true;
            KeyDown += new KeyEventHandler(Form1_KeyDown); 
        }

        void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1 && e.Control && e.Shift) // works for F1      
            {
                System.Windows.Forms.MessageBox.Show("xxxxxxxxxxxxxxxxxx");
            }

            //System.Windows.Forms.MessageBox.Show("xxxxxxxxxxxxxxxxxx");
            //System.Diagnostics.Debug.Write(e.KeyCode); 
        } 

        /// <summary>
        /// Runs the Program on Startup.
        /// </summary>
        /// <param name="RunOnStartup">True to Run on Startup, False to NOT Run on Startup.</param>
        static private void SaveRunStartup(Boolean RunOnStartup)
        {
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            if (RunOnStartup == true)
            {
                key.SetValue("TFS2010AutoDeploymentTool", System.Windows.Forms.Application.ExecutablePath.ToString());
            }
            else
            {
                key.DeleteValue("TFS2010AutoDeploymentTool", false);
            }
        }

        private delegate void tester_OnLoggingChangedDelegate(object sender, System.EventArgs e);

        private void tester_OnLoggingChanged(object sender, System.EventArgs e)
        {
            // make threadsafe
            try
            {
                if (InvokeRequired)
                {
                    BeginInvoke(new tester_OnLoggingChangedDelegate(tester_OnLoggingChanged), new object[] { sender, e });
                    return;
                }
                //update the multiline textBox
                string newLog = string.Empty;

                foreach (string logLine in tester.Logging)
                    newLog += logLine + Environment.NewLine;

                txtResult.Text = newLog;

                txtResult.SelectionLength = 0;
                txtResult.SelectionStart = txtResult.Text.Length;
                txtResult.ScrollToCaret();
            }
            catch (Exception ex)
            {
                string msg = "tester_OnLoggingChanged: " + ex.Message + " --- " + ex.StackTrace;
                //logger.LogMessageToFile(msg, 0, 0);
                //System.Windows.Forms.MessageBox.Show(msg, "tester_OnLoggingChanged: ");
                //throw;
            }
           
        }

        /*public static void AddUpdateStatusInTestInfoFile(string path, string status, string message, string infoFileKey, string infoFileKeyPC)
        {
            StringCollection AllLines = new StringCollection();
            StringCollection NewAllLines = new StringCollection();
            StringCollection FinalAllLines = new StringCollection();

            // Read all lines from test info file
            StreamReader reader = File.OpenText(Path.Combine(path, ConstCommon.TESTINFO_FILENAME));
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
                if (line.StartsWith(infoFileKey + "." + "MMMMM"))
                    System.Threading.Thread.Sleep(1000);
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
                    FinalAllLines.Add(infoFileKey + "-" + status + ":" + message);
                    hasRecord = true;
                }
                else
                {
                    FinalAllLines.Add(line);
                }
            }

            if ( hasRecord == false)
                FinalAllLines.Add(infoFileKeyPC + "-" + status + ":" + message);

            StreamWriter write = File.CreateText(Path.Combine(path, ConstCommon.TESTINFO_FILENAME));
            foreach (string line in FinalAllLines)
            {
                write.WriteLine(line);
            }
            write.Close();
        }*/


        public static bool IsAllTestStatusPassed(string[] mTestDefinitionTypes, string sTestResultFolder,
                                                 ref string sErrorMessage, Tester tester)
        {
            string path = Path.Combine(sTestResultFolder, ConstCommon.TESTINFO_FILENAME);

            tester.Log("path= " + path);   
            string[] testinfos = File.ReadAllLines(path);


            tester.Log("testinfos length= " + testinfos.Length);   
            var teststatus = new string[testinfos.Length];
            int passCnt = 0;
            Console.WriteLine("<<<>>>> testinfos.Length : " + testinfos.Length);
            try
            {
                for (int i = 1; i < testinfos.Length; i++)
                {
                    Console.WriteLine("<<< testinfos[" + i + "] : " + testinfos[i]);
                    tester.Log("<<< testinfos[" + i + "] : " + testinfos[i]);   
                    //Console.WriteLine(i + " length : " + (testinfos[i].IndexOf(":") - testinfos[i].IndexOf("-")));
                    if (testinfos[i].Trim().Length > 10) // some time by manual edit info file, info line can be empty
                    {
                        if (testinfos[i].IndexOf("-") >= 0)
                        {
                            // Windows7.32[x86.Debug]EPIAAUTOTEST1-GUI Tests Passed:Tests OK
                            teststatus[i] = testinfos[i].Substring(testinfos[i].IndexOf("-") + 1);
                            // become : GUI Tests Passed:Tests OK     
                            Console.WriteLine("<<< teststatus[" + i + "] : " + teststatus[i]);
                            tester.Log("<<< teststatus[" + i + "] : " + teststatus[i]);  
                            if (teststatus[i].Contains("GUI Tests Passed"))
                            {
                                passCnt++;
                                Console.WriteLine("GUI Tests Passed :passCnt= " + passCnt);
                            }
                        }
                        else //Windows7.64[AnyCPU.Protected] --> no '-' not tested --> not passed
                        {
                            teststatus[i] = testinfos[i];
                        }
                    }
                    else
                    {
                        passCnt++; // for empty or corrupt line, also consider as pass line
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(" exception  " + ex.Message + "---" + ex.StackTrace,
                                "IsAllTestStatusPassed:" + "testinfos.Length: " + testinfos.Length);
            }


            tester.Log("passCnt= " + passCnt);
            if (passCnt == (testinfos.Length - 1))
            {
                StreamWriter sw = File.AppendText(path);
                try
                {
                    sw.WriteLine("All Tests Complete: " + DateTime.Now.ToString());
                }
                finally
                {
                    sw.Close();
                }

                return true;
            }
            else
                return false;
        }


        private void btnConn_Click(object sender, System.EventArgs e)
        {
            try
            {
                string value22 = "192.168.253.0";

                if (value22.Length > 1)
                {
                    int indexk = value22.IndexOf('.');
                    string value = "162" + value22.Substring(indexk);
                    System.Windows.Forms.MessageBox.Show("1:"+ value);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("2");
                }
              
                //string configuration = configuration.Substring(0, configuration.IndexOf('.'));

                //List<string> xBuildDefs = new List<string>();
                /*xBuildDefs.Clear();   // List<string>
                xBuildDefs.Insert(0, "Epia.Development.Dev08.Nightly");
               
                List<BuildObject> allBuilds = null;
                try
                {
                    allBuilds = TfsUtilities.GetAllBuildObjects(xBuildDefs, "Today", tester, logger);
                }
                catch (Exception ex)
                {
                    string msg = "BuildUtilities.GetAllBuildObjects: " + ex.Message + "---" + ex.StackTrace;
                    return;
                }
                System.Windows.MessageBox.Show("mTestDefinitionTypes "  + " Length " );
               */

                string sErrorMessage = "Init";
                string[] mTestDefinitionTypes = System.IO.File.ReadAllLines("C:\\EtriccTests\\QA\\TestDefinitions\\StatisticsTestTypeDefinition.txt");
                string sTestResultFolder = @"X:\Nightly\Etricc 5 - Statistics Programs\Etricc Stat Prog.Main.Nightly\Etricc Stat Prog.Main.Nightly_20130927.1\TestResults";
                string xTestDefinitionTypes = "";
                for (int i = 0; i < mTestDefinitionTypes.Length; i++)
                {
                    xTestDefinitionTypes = xTestDefinitionTypes + "\n " + mTestDefinitionTypes[i];
                }
                System.Windows.MessageBox.Show("mTestDefinitionTypes " + xTestDefinitionTypes + " Length " + mTestDefinitionTypes.Length);



                if (IsAllTestStatusPassed(mTestDefinitionTypes, sTestResultFolder, ref sErrorMessage, tester) == true)
                {
                    // update quality to GUI Tests Passed
                    Console.WriteLine("update quality to true -----  ");
                    Thread.Sleep(1000);
                }
                else
                    System.Windows.MessageBox.Show("not passed  sErrorMessage:" + sErrorMessage);

                /*System.Windows.Forms.MessageBox.Show("DeployUtilities.getThisPCOS():" + DeployUtilities.getThisPCOS());
                if (DeployUtilities.getThisPCOS().StartsWith("Windows8.64") || DeployUtilities.getThisPCOS().StartsWith("WindowsServer2012.64"))
                {
                    string errorMsg = string.Empty;
                    ProjAppInstall.UninstallApplication(EgeminApplication.EPIA_RESOURCEFILEEDITOR, ref errorMsg);
                    //ProjAppInstall.InstallApplicationNet45(@"C:\EtriccTests\", EgeminApplication.EPIA_RESOURCEFILEEDITOR, EgeminApplication.SetupType.Default, ref errorMsg, null);
                    //ProjAppInstall.InstallApplicationNet45(@"C:\EtriccTests\", EgeminApplication.EPIA, EgeminApplication.SetupType.Default, ref errorMsg, null);
                    return;
                }*/
                
                TfsTeamProjectCollection tfsProjectCollection = null;
                Uri serverUri = new Uri(Tfs.ServerUrl);
                System.Net.ICredentials tfsCredentials
                    = new System.Net.NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

                bool reconnect = false;
                //TfsTeamProjectCollection tfsProjectCollection = null;
                while (reconnect == false)
                {
                    reconnect = true;
                    try
                    {
                        tfsProjectCollection = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                    }
                    catch (Exception ex)
                    {
                        DialogResult intResultaat = System.Windows.Forms.MessageBox.Show(ex.Message + " - " + ex.StackTrace, "reconnect TFSServer",
                                System.Windows.Forms.MessageBoxButtons.RetryCancel);
                        if (DialogResult == DialogResult.Retry)
                            reconnect = false;
                        else
                        {
                            reconnect = true;
                            return;
                        }
                    }
                }

                VersionControlServer versionControlServer = (VersionControlServer)tfsProjectCollection.GetService(typeof(VersionControlServer));
                // Check if source branch exists
                //versionControlServer.CreateBranch("$/Test/Application/Main", "$/Test/Application/Production/Test9", VersionSpec.Latest);                                
                if (tfsProjectCollection.Name != null)
                {
                    TeamProject[] projects = versionControlServer.GetAllTeamProjects(true);
                    string prjs = string.Empty;
                    for (int i = 0; i < projects.Length; i++)
                    {
                        prjs = prjs + "\n " + projects[i].Name;
                    }
                    System.Windows.MessageBox.Show("Connection OK \n Server Name is " + tfsProjectCollection.Name + " \nwith projects:" + prjs);
                }
                else
                    System.Windows.MessageBox.Show("Connection Failed ");              
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message + "" + ex.StackTrace, ex.ToString());
            }
        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            string sectionName = Constants.TestConfigSection;
            System.Configuration.Configuration config =
                  System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            TestsConfigSection customSection = (TestsConfigSection)config.Sections[sectionName];
            Configuration configScreen = new Configuration(customSection);
            configScreen.ShowDialog();
        }

        private void btnConnTFS_Click(object sender, EventArgs e)
        {
            // keep selected build def
            sSelectedBuildDefs = new List<string>();
            if (lstBoxBuildDefinitions.Items.Count > 0)
            {
                for (int i = 0; i < lstBoxBuildDefinitions.Items.Count; i++)
                    sSelectedBuildDefs.Add(lstBoxBuildDefinitions.Items[i].ToString());
            }

            int appCount = 3;
            if ( cmbTestApp2.Text.Equals(string.Empty))
                appCount = 1;
            else if( cmbTestApp3.Text.Equals(string.Empty))
                appCount = 2;

            string[] testApps = new string[appCount];
            if (appCount == 3)
            {
                testApps[0] = cmbTestApp1.Text;
                testApps[1] = cmbTestApp2.Text;
                testApps[2] = cmbTestApp3.Text;
            }

            if (appCount == 2)
            {
                testApps[0] = cmbTestApp1.Text;
                testApps[1] = cmbTestApp2.Text;
            }

            if (appCount == 1)
            {
                testApps[0] = cmbTestApp1.Text;
            }
           
          
            string[] buildDef;
            GetBuildDefinitionsForm connTFSForm = new GetBuildDefinitionsForm(testApps, ref sSelectedBuildDefs, ref sDateFilter);
            while (connTFSForm.ShowDialog() == DialogResult.OK)
            {
                sDateFilter = connTFSForm.getDateFilter();
                lstBoxBuildDefinitions.Items.Clear();
                buildDef = null;
                // extract the data from the dialog
                buildDef = connTFSForm.getBuildDefinition();
                //System.Windows.Forms.MessageBox.Show(connTFSForm.Name + " was entered into the database===leng:" + buildDef.Length  );
                //if (buildDef == null)
                //    System.Windows.Forms.MessageBox.Show(connTFSForm.Name + " was entered into the database=== null:");
                if (buildDef.Length == 0)
                    continue;
                else
                {
                    for (int i = 0; i < buildDef.Length; i++)
                    {
                        lstBoxBuildDefinitions.Items.Add(buildDef[i]);
                    }
                    break;
                }
            }
        }

        private void btnStartAuto_Click(object sender, EventArgs e)
        {
            if (MyTask != null)
            {
                tester.Log("MyTask.Status.ToString() : " + MyTask.Status.ToString());
                MyTask.Dispose();
            }
            if (lstBoxBuildDefinitions.Items.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("Please get build definitions first ", "Egemin AutoDeployment & Test Tool");
                return;
            }
            SaveTfsSettingsSection();
            Startup();
        }

        private void SaveTfsSettingsSection()
        {
            string sectionName = Constants.TfsSettingsSection;
            System.Configuration.Configuration config =
                  System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            try
            {
                // Create a custom configuration section having the same name used in the roaming configuration file.
                // This is possible because the configuration section can be overridden by other configuration files. 
                TfsSettingsSection customSection = new TfsSettingsSection();
                if (config.Sections[sectionName] == null)
                {
                    // Store console settings.
                    customSection.Element.DateFilter = sDateFilter;
                    
                    string allBuildDefsString = string.Empty;
                    // Populate build definition     
                    foreach (String strThisBuildDef in lstBoxBuildDefinitions.Items)         
                    {
                        if (strThisBuildDef.Length > 0)
                        {
                            if (allBuildDefsString.Equals(string.Empty))
                                allBuildDefsString = strThisBuildDef;
                            else
                                allBuildDefsString = allBuildDefsString + ";" + strThisBuildDef;
                        
                        }
                    } 
                    customSection.Element.BuildDefinitions = allBuildDefsString;
                    
                    //System.Windows.Forms.MessageBox.Show("cmbTestDefinitions.SelectedItem.ToString():" + cmbTestDefinitions.SelectedItem.ToString());
                    // new fields
                    customSection.Element.TestApp1 = cmbTestApp1.Text;
                    customSection.Element.App1TestDefName = cmbTestApp1DefFileName.Text;
                    customSection.Element.TestApp2 = cmbTestApp2.Text;
                    customSection.Element.App2TestDefName = cmbTestApp2DefFileName.Text;
                    customSection.Element.TestApp3 = cmbTestApp3.Text;
                    customSection.Element.App3TestDefName = cmbTestApp3DefFileName.Text;
                    // Add configuration information to the configuration file.
                    config.Sections.Add(sectionName, customSection);
                    config.Save(ConfigurationSaveMode.Modified);
                    // Force a reload of the changed section, This makes the new values available for reading.
                    ConfigurationManager.RefreshSection(sectionName);
                }
                else
                {
                    // update configuration information to the configuration file.
                    //customSection.ConsoleElement.TestProject = cmbProject.SelectedItem.ToString();
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
                    
                    foreach (XmlElement element in xmlDoc.DocumentElement)
                    {
                        if (element.Name.Equals(sectionName))
                        {
                            foreach (XmlNode node in element.ChildNodes)
                            {
                                if (node.Name.Equals("settingsElement"))
                                {
                                    XmlNode pNode = node.ParentNode;
                                    XmlNode oldNode = node;
                                    XmlNode newNode = node;
                                    newNode.Attributes["dateFilter"].Value = sDateFilter;
                                    
                                    string allBuildDefsString = string.Empty;
                                    for (int i = 0; i < lstBoxBuildDefinitions.Items.Count; i++)
                                    {
                                        if (allBuildDefsString.Equals(string.Empty))
                                            allBuildDefsString = lstBoxBuildDefinitions.Items[i].ToString();
                                        else
                                            allBuildDefsString = allBuildDefsString + ";" + lstBoxBuildDefinitions.Items[i].ToString();
                                    }
                                    newNode.Attributes["buildDefinitions"].Value = allBuildDefsString;
                                    // new fields
                                    newNode.Attributes["testApplication1"].Value = cmbTestApp1.Text;
                                    newNode.Attributes["application1TestDefName"].Value = cmbTestApp1DefFileName.Text;
                                    newNode.Attributes["testApplication2"].Value = cmbTestApp2.Text;
                                    newNode.Attributes["application2TestDefName"].Value = cmbTestApp2DefFileName.Text;
                                    newNode.Attributes["testApplication3"].Value = cmbTestApp3.Text;
                                    newNode.Attributes["application3TestDefName"].Value = cmbTestApp3DefFileName.Text;
                                    pNode.ReplaceChild(newNode, oldNode);
                                }
                            }
                        }
                    }

                    xmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
                    ConfigurationManager.RefreshSection(sectionName);
                }
            }
            catch (ConfigurationErrorsException e)
            {
                Console.WriteLine("[Error exception: {0}]",
                    e.ToString());
            }
        }

        private void Startup()
        {
            tester.LoadTestConfigSectionSettings();
            tester.LoadTfsSettingsSection();

            string sectionName = Constants.TestConfigSection;
            System.Configuration.Configuration config =
                  System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            TestsConfigSection m_TestConfigSettings = (TestsConfigSection)config.Sections[sectionName];
            if (m_TestConfigSettings == null)
                throw new Exception("Load Test Config Section Settings failed");

            sVMSwitchMode = m_TestConfigSettings.Element.VMSwitchMode;
          
            //string logPath = tester.Configuration.BuildInformationfilePath;
            //string path = Path.Combine( logPath, tester.Configuration.LogFilename );
            //logger = new Logger(@"C:\Epia3Log.txt");
            //if (tester.Configuration.EnableLog)
            //{
            logger.LogMessageToFile("--------------------------TFS2010AutoDeploymentTools----------------------", 0, 0);
            logger.LogMessageToFile("------    " + " Tools Version " + TFSQAtestToolVersion + " ------- start up --->", 0, 0);
            logger.LogMessageToFile("------    TEST MACHINE: " + System.Environment.MachineName, 0, 0);
            logger.LogMessageToFile("------    OS Version: " + System.Environment.OSVersion, 0, 0);
            //logger.LogMessageToFile("------    Configuration file: " + CONFIGPATH);
            logger.LogMessageToFile("--------------------------------------------------------------------------", 0, 0);
            //}

            //disable start button
            btnStartAuto.Enabled = false;
            btnStopAuto.Enabled = !btnStartAuto.Enabled;

            ProcessUtilities.CloseProcess( "EXCEL" );
            ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SHELL );
            ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SERVER );
            ProcessUtilities.CloseProcess( "Egemin.EPIAExplorer" );
            //Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);

            timStart.Enabled = true;
            tester.Log("Searching for available build...");
            lbSelectedConfigFile.Text = "Searching for available build...";
            timStart.Interval = 1000;
            timStart.Start();
        }

        private void btnStopAuto_Click(object sender, EventArgs e)
        {
            timStart.Stop();
            timStart.Enabled = false;
            //toggle enabled state of buttons
            btnStopAuto.Enabled = false;
            btnStartAuto.Enabled = !btnStopAuto.Enabled;
            tester.State = Tester.STATE.PENDING;
        }

        private void timStart_Tick(object sender, EventArgs e)
        {
            bool isAutoTestRunning = false;
            try
            {
                lbSelectedConfigFile.Text = tester.GetCurrentBuildInTesting();

                if ( sDateFilter.Equals("<Any Time>") )
                {
                    timStart.Interval = sTickInterval;
                    timerLog("Date Filter  =" + sDateFilter + " and timStart.Interval = " + timStart.Interval);
                    //tester.Log("Date Filter  =" + sDateFilter + " and timStart.Interval = " + timStart.Interval);
                }
                else
                    timStart.Interval = sTickInterval;

                System.Diagnostics.Process proc = null;
                int ipid = 9999;
                if (isAutoTestRunning == false)
                {
                    ipid = ProcessUtilities.GetApplicationProcessID("Egemin.Epia.Testing.QATestEpiaUI", out proc);
                    if (ipid > 0)  // check Epia testing is running
                    {
                        timerLog("Egemin.Epia.Testing.QATestEpiaUI Test is running");
                        isAutoTestRunning = true;
                    }
                }

                if (isAutoTestRunning == false)
                {
                    ipid = ProcessUtilities.GetApplicationProcessID("Egemin.Epia.Testing.QATestEtriccUI", out proc);
                    if (ipid > 0)  // check Etricc testing is running
                    {
                        timerLog("Egemin.Epia.Testing.QATestEtriccUI Test is running");
                        isAutoTestRunning = true;
                    }
                }

                if (isAutoTestRunning == false)
                {
                    ipid = ProcessUtilities.GetApplicationProcessID("Egemin.Epia.Testing.QATestEtriccStatistics", out proc);
                    if (ipid > 0)  // check KC testing is running
                    {
                        timerLog("Egemin.Epia.Testing.QATestEtriccStatistics Test is running");
                        isAutoTestRunning = true;
                    }
                }

                if (isAutoTestRunning == false)
                {
                    ipid = ProcessUtilities.GetApplicationProcessID("Egemin.Epia.Testing.QATestEpiaProtected", out proc);
                    if (ipid > 0)  // check KC testing is running
                    {
                        timerLog("Egemin.Epia.Testing.QATestEpiaProtected Test is running");
                        isAutoTestRunning = true;
                    }
                }

                if (isAutoTestRunning == false)
                {
                    ipid = ProcessUtilities.GetApplicationProcessID("EPIA.Explorer", out proc);
                    if (ipid > 0)  // check Etricc 5 testing is running
                    {
                        timerLog("Etricc 5 Test is running");
                        //tester.ClickUiScreenActionToAvoidScreenStandBy();
                        isAutoTestRunning = true;
                    }
                }

                if (isAutoTestRunning == false)
                {
                    ipid = ProcessUtilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SERVER, out proc);
                    if (ipid > 0)
                    {
                        timerLog("Test " + ConstCommon.EGEMIN_EPIA_SERVER + " is running");
                        isAutoTestRunning = true;
                    }
                }

                if (isAutoTestRunning == false)
                {
                    ipid = ProcessUtilities.GetApplicationProcessID("Egemin.Epia.Testing.KimberlyClarkGUIAutoTest", out proc);
                    if (ipid > 0)  // check KC testing is running
                    {
                        timerLog("Egemin.Epia.Testing.KImberlyClarkGUIAutoTest Test is running");
                        isAutoTestRunning = true;
                    }
                }

                if (isAutoTestRunning == false)
                {
                    ipid = ProcessUtilities.GetApplicationProcessID("Egemin.Epia.Testing.QATestEpiaNet45UI", out proc);
                    if (ipid > 0)  // check KC testing is running
                    {
                        timerLog("Egemin.Epia.Testing.QATestEpiaNet45UI Test is running");
                        isAutoTestRunning = true;
                    }
                }

                //tester.Log("timStart_Tick...;  btnStartAuto.Enabled" + btnStartAuto.Enabled);
                if (isAutoTestRunning == false)
                {
                    if (btnStartAuto.Enabled)
                    {
                        isAutoTestRunning = true;
                    }
                }

                if (isAutoTestRunning == false)
                {
                    if (tester.GetDeploymentStatus() == true)
                    {
                        TimeSpan time = DateTime.Now - tester.GetDeploymentEndTime();
                        if (time.TotalSeconds < 120)
                        {
                            timerLog("--------- after deploymenttime:" + time.TotalSeconds);
                            isAutoTestRunning = true;
                        }
                    }
                }

                if (isAutoTestRunning == false)
                {
                    string status = tester.State.ToString();
                    if (tester.State == Tester.STATE.PENDING)
                    {
                        tester.SetDeploymentStatus(false);
                        List<string> defs = new List<string>();
                        for (int i = 0; i < lstBoxBuildDefinitions.Items.Count; i++)
                        {
                            defs.Add(lstBoxBuildDefinitions.Items[i].ToString());
                        }

                        //tester.Log("TimStart_Tick mTime.7:tester.defs count: " + defs.Count);
                        //tester.Start(cmbProject.SelectedItem.ToString(), cmbTestApp.SelectedItem.ToString(), cmbTargetPlatform.SelectedItem.ToString(),
                        //    defs, sDateFilter, ref sStartUpTime, chkProtected.Checked);
                        
                        //tester.Start(ref sStartUpTime);
                        // user task
                        tester.Log("Task.Factory.StartNew(() => { tester.Start(ref sStartUpTime); } ");
                        MyTask = Task.Factory.StartNew(() => { tester.Start(ref sStartUpTime); });

                        lbStartTime.Text = sStartUpTime.ToString("yyyyMMdd HH:mm:ss");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("timStart_Tick exception: " + ex.Message + " --- " + ex.StackTrace);
                tester.Log("timStart_Tick exception: " + ex.Message + " --- " + ex.StackTrace);
                //logger.LogMessageToFile(ex.Message+" ----- "+ex.StackTrace, 0, 0);
            }
        }

        private void timerLog(string message)
        {
            try
            {
                logger.LogMessageToFile(message, 0, 0);
            }
            catch (Exception ex)
            {
               logger.LogMessageToFile("check " + message + " exception:" + ex.Message + " --- " + ex.StackTrace, 0, 0);
               Console.WriteLine("check "+message + " exception:"+ ex.Message +" --- "+ ex.StackTrace);
            }
        }

        static public AutomationElement ElementFromCursor()
        {
            // Convert mouse position from System.Drawing.Point to System.Windows.Point.
            System.Windows.Point point = new System.Windows.Point(System.Windows.Forms.Cursor.Position.X, System.Windows.Forms.Cursor.Position.Y);
            AutomationElement element = AutomationElement.FromPoint(point);
            return element;
        }

        private void mnuAbout_Click(object sender, EventArgs e)
        {
            /*string version = string.Empty;
            Assembly asm = Assembly.GetExecutingAssembly();
            System.Diagnostics.FileVersionInfo oFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(asm.Location);

            string[] splits = asm.FullName.ToString().Split(',');
            string name = splits[0];
            string AssemblyVersion = splits[1];

            object[] attrs = asm.GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
            version = version + "\n" + "Title: " + ((AssemblyTitleAttribute)attrs[0]).Title;
            version = version + "\n\n" + "Description: " + oFileVersionInfo.FileDescription;

            object[] attrs1 = asm.GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
            version = version + "\n\n" + "Company: " + ((AssemblyCompanyAttribute)attrs1[0]).Company;

            object[] attri2 = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
            version = version + "\n\n" + "Product: " + ((AssemblyProductAttribute)attri2[0]).Product;
            version = version + "\n\n" + "Assembly" + AssemblyVersion;
            version = version + "\n\n" + "AssemblyFileVersion: " + oFileVersionInfo.FileVersion.ToString();
            version = version + "\n\n" + "" + oFileVersionInfo.LegalTrademarks;
            version = version + "\n\n" + "" + oFileVersionInfo.LegalCopyright;       
            MessageBox.Show(version, name, MessageBoxButtons.OK, MessageBoxIcon.Information);

            */
            //System.Windows.MessageBox.Show("Egemin E'pia Software Application\n\n E'pia Applications Deployment and Testing Tool\n\n Version:"
            //    + TFSQAtestToolVersion);
            
            AutomationElement aeDetectedUIElement;
            bool search = true;

            while (search)
            {
                System.Threading.Thread.Sleep(10000);
                aeDetectedUIElement = ElementFromCursor();
                if (aeDetectedUIElement != null)
                {
                    tester.Log(" ---------- aeDetectedUIElement ---------- OK");
                    tester.Log(" ---------- aeDetectedUIElement.Current.Name:" + aeDetectedUIElement.Current.Name);
                    tester.Log(" ---------- aeDetectedUIElement.Current.ClassName:" + aeDetectedUIElement.Current.ClassName);
                    tester.Log(" ---------- aeDetectedUIElement.Current.ControlType.LocalizedControlType:" + aeDetectedUIElement.Current.ControlType.LocalizedControlType);
                }
                //MessageBox.Show("aeDetectedUIElement OK ", "aeDetectedUIElement ", );
                else
                {
                    MessageBox.Show("aeDetectedUIElement ", "aeDetectedUIElement == null ");
                }

                TreeWalker walker = TreeWalker.ControlViewWalker;
                if (walker != null)
                {
                    tester.Log(" ---------- walker ---------- OK");
                }
                AutomationElement elementParent = null;
                AutomationElement elementParent2 = null;
                AutomationElement elementParent3 = null;
                AutomationElement elementParent4 = null;
                AutomationElement elementParent5 = null;
                AutomationElement node = aeDetectedUIElement;
                if (node != null)
                {
                    tester.Log(" ---------- node ---------- OK");
                }
                //if (node == elementRoot) return node;
                //do
                //{
                elementParent = walker.GetParent(node);
                if (elementParent != null)
                {
                    tester.Log(" ---------- elementParent ---------- OK");
                    elementParent2 = walker.GetParent(elementParent);
                    if (elementParent2 != null)
                    {
                        tester.Log(" ---------- elementParent2 ---------- OK");
                        elementParent3 = walker.GetParent(elementParent2);
                    }
                    else
                    {
                        MessageBox.Show("elementParent2 ", "elementParent2 == null ");
                    }

                    if (elementParent3 != null)
                    {
                        tester.Log(" ---------- elementParent3 ---------- OK");
                        elementParent4 = walker.GetParent(elementParent3);
                    }
                    else
                    {
                        tester.Log(" ---------- elementParent3 == null");
                    }

                    if (elementParent4 != null)
                    {
                        tester.Log(" ---------- elementParent4 ---------- OK");
                        elementParent5 = walker.GetParent(elementParent4);
                    }
                    else
                    {
                        tester.Log(" ---------- elementParent4 == null");
                    }
                }
                else
                {
                    MessageBox.Show("elementParent ", "elementParent == null ");
                }

                tester.Log(" ---------- Start text ---------- OK");
                string text = "ele name: " + aeDetectedUIElement.Current.Name + Environment.NewLine;
                tester.Log(" ---------- text1 ----------:" + text);
                text = text + "ele automationID: " + aeDetectedUIElement.Current.AutomationId + Environment.NewLine;
                tester.Log(" ---------- text2 ----------:" + text);
                text = text + "ele ClassName: " + aeDetectedUIElement.Current.ClassName + Environment.NewLine;
                tester.Log(" ---------- text3 ----------:" + text);
                text = text + "ele controlType: " + aeDetectedUIElement.Current.ControlType.ProgrammaticName + Environment.NewLine;
                tester.Log(" ---------- text4 ----------:" + text);
                if (aeDetectedUIElement.GetSupportedPatterns().Length > 0)
                {
                    for (int i = 0; i < aeDetectedUIElement.GetSupportedPatterns().Length; i++)
                    {
                        text = text + "ele CurrentPattern: " + aeDetectedUIElement.GetSupportedPatterns()[i].ProgrammaticName + Environment.NewLine;
                    }
                }

                if (elementParent != null)
                {
                    text = text + "-----------ele Parent name: " + elementParent.Current.Name + Environment.NewLine;
                    //text = text + "ele Parent name: " + elementParent.Current.Name + Environment.NewLine;
                    text = text + "ele Parent ControlType: " + elementParent.Current.ControlType.ProgrammaticName + Environment.NewLine;
                    text = text + "ele Parent AutomationId: " + elementParent.Current.AutomationId + Environment.NewLine;
                    text = text + "ele ClassName: " + elementParent.Current.ClassName + Environment.NewLine;
                }

                if (elementParent2 != null)
                {
                    text = text + "----------ele Parent2 name: " + elementParent2.Current.Name + Environment.NewLine;
                    text = text + "ele elementParent2 name: " + elementParent2.Current.Name + Environment.NewLine;
                    text = text + "ele elementParent2 ControlType: " + elementParent2.Current.ControlType.ProgrammaticName + Environment.NewLine;
                    text = text + "ele elementParent2 AutomationId: " + elementParent2.Current.AutomationId + Environment.NewLine;
                    text = text + "ele ClassName: " + elementParent2.Current.ClassName + Environment.NewLine;
                }

                if (elementParent3 != null)
                {
                    text = text + "----------ele Parent3 name: " + elementParent3.Current.Name + Environment.NewLine;
                    text = text + "ele elementParent3 name: " + elementParent3.Current.Name + Environment.NewLine;
                    text = text + "ele elementParent3 ControlType: " + elementParent3.Current.ControlType.ProgrammaticName + Environment.NewLine;
                    text = text + "ele elementParent3 AutomationId: " + elementParent3.Current.AutomationId + Environment.NewLine;
                    text = text + "ele elementParent3 ClassName: " + elementParent3.Current.ClassName + Environment.NewLine;
                }

                if (elementParent4 != null)
                {
                    text = text + "----------ele Parent4 name: " + elementParent4.Current.Name + Environment.NewLine;
                    text = text + "ele elementParent4 name: " + elementParent4.Current.Name + Environment.NewLine;
                    text = text + "ele elementParent4 ControlType: " + elementParent4.Current.ControlType.ProgrammaticName + Environment.NewLine;
                    text = text + "ele elementParent4 AutomationId: " + elementParent4.Current.AutomationId + Environment.NewLine;
                    text = text + "ele elementParent4 ClassName: " + elementParent4.Current.ClassName + Environment.NewLine;
                }

                if (elementParent5 != null)
                {
                    text = text + "----------ele Parent5 name: " + elementParent5.Current.Name + Environment.NewLine;
                    text = text + "ele elementParent5 name: " + elementParent5.Current.Name + Environment.NewLine;
                    text = text + "ele elementParent5 ControlType: " + elementParent5.Current.ControlType.ProgrammaticName + Environment.NewLine;
                    text = text + "ele elementParent5 AutomationId: " + elementParent5.Current.AutomationId + Environment.NewLine;
                    text = text + "ele elementParent5 ClassName: " + elementParent5.Current.ClassName + Environment.NewLine;
                }


                text = text + "--------XX--ele root name: " + AutomationElement.RootElement.Current.Name + Environment.NewLine;
                text = text + "ele root name: " + AutomationElement.RootElement.Current.Name + Environment.NewLine;
                text = text + "ele root ControlType: " + AutomationElement.RootElement.Current.ControlType.ProgrammaticName + Environment.NewLine;
                text = text + "ele root AutomationId: " + AutomationElement.RootElement.Current.AutomationId + Environment.NewLine;

                //text = text + "ele Parent name: " + elementParent.Current.Name + Environment.NewLine;   WILL CHECK LATER

                //text = "ele name: " + aeDetectedUIElement.Current. + Environment.NewLine;
                //text = "ele name: " + aeDetectedUIElement.Current.Name + Environment.NewLine;
                //text = "ele name: " + aeDetectedUIElement.Current.Name + Environment.NewLine; 

                DialogResult dr = MessageBox.Show(text, "detected ui element", MessageBoxButtons.RetryCancel, MessageBoxIcon.Information);
                if (dr == DialogResult.Retry)
                {
                    search = true;
                }
                else
                    search = false;

            }
        }

        private void mnuExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button_buildSelector_Click(object sender, EventArgs e)
        {
            this.openFileDialog.Filter = "Setup files (*.msi)|*.msi";
            this.openFileDialog.Title = "Select setup to test";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtBuildPath.Text = System.IO.Path.GetDirectoryName(openFileDialog.FileName);
                //txtFilePath.Text = openFileDialog.FileName;
            }
        }

        private void lstBoxBuildDefinitions_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtBuildPath.Text = string.Empty;
        }

        private void cmbTargetPlatform_SelectedIndexChanged(object sender, EventArgs e)
        {
            //sDateFilter = "<Any Time>";
        }

        private void txtBuildPath_DoubleClick(object sender, System.EventArgs e)
        {
            this.openFileDialog.Filter = "Setup files (*.msi)|*.msi";
            this.openFileDialog.Title = "Select setup to test";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtBuildPath.Text = System.IO.Path.GetDirectoryName(openFileDialog.FileName);
                //txtFilePath.Text = openFileDialog.FileName;
            }
        }


        private void btnStartManual_Click(object sender, EventArgs e)
        {
            if (!chkContinueAuto.Checked)
                btnStopAuto_Click(sender, e);

            // Start manual test here
            if (txtBuildPath.Text == string.Empty)
                System.Windows.MessageBox.Show("Please select a valid Build");
            else
            {
                //tester.LoadConfiguration(lbSelectedConfigFile.Text);
                tester.LoadTestConfigSectionSettings();
                tester.LoadTfsSettingsSection();
                tester.SetTestAutoMode(false);

                ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SHELL );
                ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_EPIA_SERVER );
                ProcessUtilities.CloseProcess( ConstCommon.EGEMIN_ETRICC_SERVER );

                //tester.Start2(this.txtBuildPath.Text, cmbProject.SelectedItem.ToString(), cmbTestApp.SelectedItem.ToString(), cmbTargetPlatform.SelectedItem.ToString(),
                //    null, sDateFilter, ref sStartUpTime, chkProtected.Checked);
                tester.Start2(this.txtBuildPath.Text, ref sStartUpTime);
            }
        }

        private void computerInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string pcos = DeployUtilities.getThisPCOS();
            string pa = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE");
            string systemBits =  ""+((String.IsNullOrEmpty(pa) || String.Compare(pa, 0, "x86", 0, 3, true) == 0) ? 32 : 64);
            System.Windows.MessageBox.Show("PROCESSOR_ARCHITECTURE:"+pa + System.Environment.NewLine +"system (bit): " + systemBits
                + Environment.NewLine + "PC OS:"+pcos);
        }

        /*protected override void OnLoad(EventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show("Enter Load");
            base.OnLoad(e);
            //System.Windows.Forms.MessageBox.Show("Exit Load");
        }*/

        protected override void OnShown(EventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show("Enter Show");
            base.OnShown(e);
            //System.Windows.Forms.MessageBox.Show("Exit Show");
            //Thread executableThread = new Thread(new ThreadStart(DeployUtilities.StartExecution));
            //executableThread.Start();
            //Thread.Sleep(5000);

            if (autoStartup)
            {
                System.Configuration.Configuration config =
                  System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                string sectionName = Constants.TfsSettingsSection;
                TfsSettingsSection customSection = (TfsSettingsSection)config.Sections[sectionName];
                if (customSection != null)
                {   // fill UI from TestSection 
                    sDateFilter = customSection.Element.DateFilter;
                    // new fields -------------------------------------------------
                    cmbTestApp1.Text = customSection.Element.TestApp1;
                    cmbTestApp1DefFileName.Text = customSection.Element.App1TestDefName;
                    cmbTestApp2.Text = customSection.Element.TestApp2;
                    cmbTestApp2DefFileName.Text = customSection.Element.App2TestDefName;
                    cmbTestApp3.Text = customSection.Element.TestApp3;

                    //System.Windows.Forms.MessageBox.Show("cmbTestApp3.Text:" + cmbTestApp3.Text, "Length:" + cmbTestApp3.Text.Length);

                    cmbTestApp3DefFileName.Text = customSection.Element.App3TestDefName;
                    if (cmbTestApp2.Text.Length == 0)
                        cmbTestApp2DefFileName.Text = "";
                    if (cmbTestApp3.Text.Length == 0)
                        cmbTestApp3DefFileName.Text = "";
                    // end new fields ----------------------------------------------

                    // already loaded during onLoad
                    /*string allBuildDefsString = customSection.Element.BuildDefinitions;
                    string[] strArray = allBuildDefsString.Split(';');
                    ArrayList arList = new ArrayList();
                    for (int ix = 0; ix < strArray.Length; ix++)
                    {
                        if (strArray[ix].Length > 0)
                            lstBoxBuildDefinitions.Items.Add(strArray[ix]);
                    }*/

                    if (cmbTestApp1.Text.Length > 4)
                        cmbTestApp1DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp1.Text);

                    if (cmbTestApp2.Text.Length > 4)
                        cmbTestApp2DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp2.Text);

                    if (cmbTestApp3.Text.Length > 4)
                        cmbTestApp3DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp3.Text);
                }
                else
                {
                    if (cmbTestApp1.Text.Length > 4)
                        cmbTestApp1DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp1.Text);

                    if (cmbTestApp2.Text.Length > 4)
                        cmbTestApp2DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp2.Text);

                    if (cmbTestApp3.Text.Length > 4)
                        cmbTestApp3DefFileName.Text = TfsUtilities.GetTestDefNameFromTestApp(cmbTestApp3.Text);
                }
            }

            if (autoPressStartButton)
            {
                Thread.Sleep(2000);
                AutomationElement aeWindow = DeployUtilities.GetMainWindow("ToolsForm");
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
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show("Enter OnFormClosing");
            base.OnFormClosing(e);
            
            SaveTfsSettingsSection();
            //System.Windows.Forms.MessageBox.Show("Exit OnFormClosing");
        }

        private void VMSwitchModeTimer_Tick(object sender, EventArgs e)
        {
            string remoteCommandFile = Path.Combine(@"C:\EtriccTests", "RemoteCommand.txt");
            // Check test Working file
            System.IO.FileInfo workFile = new FileInfo(remoteCommandFile);
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
           
            sCounter++;
            sCounter = sCounter % 15000;
            try
            {
                #region process remote command
                if (info.Length == 0)
                    Console.WriteLine(sCounter + "Current command Empty -------------- ");
                else
                {
                    Console.WriteLine(sCounter + " Current command Received -------------- " + info);

                    if (info.ToLower().StartsWith("stop"))
                    {
                        Console.WriteLine("current command Received -------------- Stop ");
                        // empty RemoteCommand.txt file, otherwise it will run Startup again and kill server and shell
                        StreamWriter write = File.CreateText(Path.Combine(@"C:\EtriccTests", "RemoteCommand.txt"));
                        write.Close();
                        //if (sRemoteVMstarted == true)
                        //{
                        timStart.Stop();
                        timStart.Enabled = false;
                        //toggle enabled state of buttons
                        btnStopAuto.Enabled = false;
                        btnStartAuto.Enabled = !btnStopAuto.Enabled;
                        tester.State = Tester.STATE.PENDING;

                        btnStartAuto.Enabled = true;
                        btnStopAuto.Enabled = false;

                        // change Start to IsStart  RemoteCommand.txt file, otherwise it will run Startup again and kill server and shell
                        WriteRemoteCommandReply("IsStopped");
                        //}
                    }
                    else if (info.ToLower().StartsWith("get tfsqatesttools pid"))
                    {
                        Console.WriteLine("Start command Received -------------- ");
                     
                        btnStartAuto.Enabled = false;
                        btnStopAuto.Enabled = true;
                        int ipid = 9999;
                        System.Diagnostics.Process proc = null;
                        ipid = TestTools.ProcessUtilities.GetApplicationProcessID("TFSQATestTools", out proc);
                        if (ipid > 0)  // check Epia testing is running
                        {
                            timerLog("Egemin.Epia.Testing.TFSQATestTools is running");
                        }
                        WriteRemoteCommandReply("TFSQATestTools PID:" + ipid);
                    }
                    else if (info.ToLower().StartsWith("startup new and kill old pid"))
                    {
                        Console.WriteLine("Start command Received -------------- ");
                        try
                        {
                            Console.WriteLine("number TFSQATestTools is running: = 000000000000000000000000000" );
                            Process[] ps = Process.GetProcessesByName("TFSQATestTools");
                            int runTools = ps.Length;
                            Console.WriteLine("number TFSQATestTools is running: = " + runTools);
                            //Startup();

                            TestTools.ProcessUtilities.StartProcessNoWait(TestTools.OSVersionInfoClass.ProgramFilesx86() +
                                    @"\Dematic\AutomaticTesting",
                                    "TFSQATestTools.exe", string.Empty);


                            while (ps.Length < runTools + 1)
                            {
                                Console.WriteLine("running: number is not OK  and wait 5 sec. -> num = " + ps.Length);
                                Thread.Sleep(2000);
                                ps = Process.GetProcessesByName("TFSQATestTools");
                            }

                            Console.WriteLine("final running: number is OK  -> num = " + ps.Length);
                            WriteRemoteCommandReply("Is startup new and kill old pid OK:");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                    else if (info.ToLower().StartsWith("kill old pid"))
                    {
                        Console.WriteLine("Start command Received -------------- ");
                        int ipid = 0;
                        try
                        {
                            int parmIndex = info.IndexOf("pid");
                            if (parmIndex > 0)
                            {
                               string retParm = info.Substring(parmIndex + 4);
                               ipid = int.Parse(retParm);
                               Console.WriteLine("old pid = " + ipid);

                               Process p = Process.GetProcessById(ipid);
                               WriteRemoteCommandReply(ipid + " is killed");
                               p.Kill();
                            }
                            //WriteRemoteCommandReply(ipid + " is killed222222");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                    }
                    else if (info.ToLower().StartsWith("gettoolsrunningstatus"))
                    {
                        Console.WriteLine("Start command Received -------------- gettoolsrunningstatus");
                        Thread executableThread = new Thread(new ThreadStart(GetToolsRunningStatus));
                        executableThread.Name = "gettoolsrunningstatus";
                        executableThread.Start();



                       
                    }
                    else if (info.ToLower().StartsWith("start"))
                    {
                        Console.WriteLine("Start command Received -------------- ");
                        //if (sRemoteVMstarted == false)
                        //{
                        if (lstBoxBuildDefinitions.Items.Count == 0)
                        {
                            System.Windows.Forms.MessageBox.Show("Please get build definitions first ", "Egemin AutoDeployment & Test Tool");
                            return;
                        }
                        btnStartAuto.Enabled = false;
                        btnStopAuto.Enabled = true;

                        SaveTfsSettingsSection();
                        Startup();
                        // change Start to IsStart  RemoteCommand.txt file, otherwise it will run Startup again and kill server and shell
                        WriteRemoteCommandReply("IsStarted");
                        //}
                    }
                    else if (info.ToLower().StartsWith("vmscreen"))
                    {
                        string machineName = System.Environment.MachineName;
                        AutomationElement root = AutomationElement.RootElement;
                        if (root == null )
                            WriteRemoteCommandReply(machineName + " screen is off");
                        else
                            WriteRemoteCommandReply(machineName + " screen is open");
                    }
                    else
                    {
                        Console.WriteLine("other command  -------------- " + info);
                    }

                    if (sCounter % (5) == 0)
                    {
                        if (DeployUtilities.TestRunning())
                        {

                            Console.WriteLine("<<<>>>> TestRunning : ");
                            bool wReply = false;
                            while (wReply == false)
                            {
                                try
                                {
                                    StreamWriter writeReply = File.CreateText(sCommandReplyFile);
                                    writeReply.WriteLine("TestRunning");
                                    writeReply.Close();
                                    wReply = true;
                                }
                                catch (Exception ex)
                                {
                                    wReply = false;
                                    Thread.Sleep(5000);
                                    Console.WriteLine("Write TestRunning exception: " + ex.Message);
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("NNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNNN Test NOT Running : ");
                            bool wReply = false;
                            while (wReply == false)
                            {
                                try
                                {
                                    StreamWriter writeReply = File.CreateText(sCommandReplyFile);
                                    writeReply.WriteLine("TestNotRunning");
                                    writeReply.Close();
                                    wReply = true;
                                }
                                catch (Exception ex)
                                {
                                    wReply = false;
                                    Thread.Sleep(5000);
                                    Console.WriteLine("Write TestRunning exception: " + ex.Message);
                                }
                            }
                        }
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                string remoteCommandExceptionFile = Path.Combine(@"C:\EtriccTests", "RemoteCommandException.txt");
                StreamWriter writeExc = File.CreateText(remoteCommandExceptionFile);
                writeExc.WriteLine("remoteCommandException:"+ex.Message +" --- "+ ex.StackTrace);
                writeExc.Close();
            }
        }

        private void GetToolsRunningStatus()
        {
            string retMsg = "try to get status";
                        try
                        {
                            AutomationElement aeToolsWindow = null;
                            DateTime sStartTime = DateTime.Now;
                            TimeSpan sTime = DateTime.Now - sStartTime;
                            Console.WriteLine(" time is :" + sTime.TotalSeconds);
                            while (aeToolsWindow == null && sTime.TotalSeconds < 300)
                            {
                                Console.WriteLine("get windows -------------- ");
                                AutomationElementCollection aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children,
                                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
                                    sTime = DateTime.Now - sStartTime;
                                    Thread.Sleep(3000);
                                    Console.WriteLine("get windows --------------aeAllWindows.Count =  " + aeAllWindows.Count);
                                    for (int i = 0; i < aeAllWindows.Count; i++)
                                    {
                                        Console.WriteLine("Found: aeWindow[" + i + "]=" + aeAllWindows[i].Current.Name);
                                        if (aeAllWindows[i].Current.Name.Equals("E'pia QA Test Tool"))
                                        {
                                            aeToolsWindow = aeAllWindows[i];
                                            Console.WriteLine("xxxFound: aeWindow[" + i + "]=" + aeToolsWindow.Current.Name);
                                            retMsg = "tools is running";
                                            break;
                                        }
                                    }
                            }

                            Console.WriteLine("222222     get windows -------------- ");
                            if (aeToolsWindow == null)
                            {
                                retMsg = "tools is NOT running";
                                Console.WriteLine(retMsg);
                            }
                            else
                            {
                                Console.WriteLine("33333     get windows -------------- ");
                                AutomationElement aeCrashWindow = null;
                                sStartTime = DateTime.Now;
                                sTime = DateTime.Now - sStartTime;
                                while (aeCrashWindow == null && sTime.TotalSeconds < 8)
                                {
                                    AutomationElementCollection aeAllWindows = aeToolsWindow.FindAll(TreeScope.Descendants,
                                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
                                    Console.WriteLine("4444     get windows -------------- " + aeAllWindows.Count);
                                    Console.WriteLine("4444     sTime.TotalSeconds " + sTime.TotalSeconds);
                                    Thread.Sleep(3000);
                                    sTime = DateTime.Now - sStartTime;
                                    for (int i = 0; i < aeAllWindows.Count; i++)
                                    {
                                        Console.WriteLine("ZZZZZ: aeWindow[" + i + "]=" + aeAllWindows[i].Current.Name);
                                        if (aeAllWindows[i].Current.Name.Equals("TFSQATestTools"))
                                        {
                                            aeCrashWindow = aeAllWindows[i];
                                            Console.WriteLine("Found: aeWindow[" + i + "]=" + aeToolsWindow.Current.Name);
                                            retMsg = "tools is crashed";
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            retMsg = ex.Message;
                            Console.WriteLine(ex.Message);
                        }


                        WriteRemoteCommandReply(retMsg);

            
        }


        private void WriteRemoteCommandReply(string message)
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

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtResult.Clear();
            tester.Logging.Clear();
        }

        private void cmbTestApp1_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstBoxBuildDefinitions.Items.Clear();
            cmbTestApp2.Text = string.Empty;
            cmbTestApp3.Text = string.Empty;
            if (cmbTestApp1.Text.Equals(TestApp.EPIA4))
            {
                cmbTestApp2.Items.Clear();
                cmbTestApp2.Items.Add(TestApp.ETRICCUI);
                cmbTestApp2.Items.Add(TestApp.ETRICCSTATISTICS);
                cmbTestApp2.Items.Add(string.Empty);

                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(TestApp.ETRICCUI);
                cmbTestApp3.Items.Add(TestApp.ETRICCSTATISTICS);
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp1DefFileName.Text = TestDefName.EPIA4; 
                cmbTestApp2DefFileName.Text = "";
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.ETRICCUI))
            {
                cmbTestApp2.Items.Clear();
                cmbTestApp2.Items.Add(TestApp.EPIA4);
                cmbTestApp2.Items.Add(TestApp.ETRICCSTATISTICS);
                cmbTestApp2.Items.Add(string.Empty);

                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(TestApp.EPIA4);
                cmbTestApp3.Items.Add(TestApp.ETRICCSTATISTICS);
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp1DefFileName.Text = TestDefName.ETRICCUI;
                cmbTestApp2DefFileName.Text = "";
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.ETRICCSTATISTICS))
            {
                cmbTestApp2.Items.Clear();
                cmbTestApp2.Items.Add(TestApp.EPIA4);
                cmbTestApp2.Items.Add(TestApp.ETRICCUI);
                cmbTestApp2.Items.Add(string.Empty);

                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(TestApp.EPIA4);
                cmbTestApp3.Items.Add(TestApp.ETRICCUI);
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp1DefFileName.Text = TestDefName.ETRICCSTATISTICS;
                cmbTestApp2DefFileName.Text = "";
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.EPIANET45))
            {
                cmbTestApp2.Items.Clear();
                cmbTestApp2.Items.Add(TestApp.ETRICCNET45);
                cmbTestApp3.Items.Clear();

                cmbTestApp1DefFileName.Text = TestDefName.EPIANET45;
                cmbTestApp2DefFileName.Text = "";
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.ETRICCNET45))
            {
                cmbTestApp2.Items.Clear();
                cmbTestApp2.Items.Add(TestApp.EPIANET45);
                cmbTestApp3.Items.Clear();

                cmbTestApp1DefFileName.Text = TestDefName.ETRICCNET45;
                cmbTestApp2DefFileName.Text = "";
                cmbTestApp3DefFileName.Text = "";
            }
        }

        private void cmbTestApp2_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstBoxBuildDefinitions.Items.Clear();
            cmbTestApp3.Text = string.Empty;
            if (cmbTestApp2.Text.Equals(string.Empty) )
            {
                cmbTestApp2DefFileName.Text = "";
                cmbTestApp3.Items.Clear();
            }
            else if (cmbTestApp1.Text.Equals(TestApp.EPIA4) && (cmbTestApp2.Text.Equals(TestApp.ETRICCUI)))
            {
                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(TestApp.ETRICCSTATISTICS);
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp2DefFileName.Text = TestDefName.ETRICCUI;
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.EPIA4) && (cmbTestApp2.Text.Equals(TestApp.ETRICCSTATISTICS)))
            {
                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(TestApp.ETRICCUI);
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp2DefFileName.Text = TestDefName.ETRICCSTATISTICS;
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.ETRICCUI) && (cmbTestApp2.Text.Equals(TestApp.EPIA4)))
            {
                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(TestApp.ETRICCSTATISTICS);
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp2DefFileName.Text = TestDefName.EPIA4; 
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.ETRICCUI) && (cmbTestApp2.Text.Equals(TestApp.ETRICCSTATISTICS)))
            {
                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(TestApp.EPIA4);
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp2DefFileName.Text = TestDefName.ETRICCSTATISTICS;
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.ETRICCSTATISTICS) && (cmbTestApp2.Text.Equals(TestApp.EPIA4)))
            {
                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(TestApp.ETRICCUI);
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp2DefFileName.Text = TestDefName.EPIA4; 
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.ETRICCSTATISTICS) && (cmbTestApp2.Text.Equals(TestApp.ETRICCUI)))
            {
                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(TestApp.EPIA4);
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp2DefFileName.Text = TestDefName.ETRICCUI;
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.EPIANET45) && (cmbTestApp2.Text.Equals(TestApp.ETRICCNET45)))
            {
                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp2DefFileName.Text = TestDefName.ETRICCNET45;
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp1.Text.Equals(TestApp.ETRICCNET45) && (cmbTestApp2.Text.Equals(TestApp.EPIANET45)))
            {
                cmbTestApp3.Items.Clear();
                cmbTestApp3.Items.Add(string.Empty);

                cmbTestApp2DefFileName.Text = TestDefName.EPIANET45;
                cmbTestApp3DefFileName.Text = "";
            }
        }

        private void cmbTestApp3_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstBoxBuildDefinitions.Items.Clear();
            if (cmbTestApp3.Text.Equals(string.Empty))
            {
                cmbTestApp3DefFileName.Text = "";
            }
            else if (cmbTestApp3.Text.Equals(TestApp.EPIA4))
            {
                cmbTestApp3DefFileName.Text = TestDefName.EPIA4; 
            }
            else if (cmbTestApp3.Text.Equals(TestApp.ETRICCUI))
            {
                cmbTestApp3DefFileName.Text = TestDefName.ETRICCUI;
            }
            else if (cmbTestApp3.Text.Equals(TestApp.ETRICCSTATISTICS))
            {
                cmbTestApp3DefFileName.Text = TestDefName.ETRICCSTATISTICS;
            }

        }

        private void HostTestMenuItem_Click(object sender, EventArgs e)
        {
            HostTestForm hostTestForm = new HostTestForm();
            hostTestForm.ShowDialog();
        }

        private void txtBuildPath_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
