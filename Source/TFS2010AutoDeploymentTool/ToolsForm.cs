#region Using directives
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows;
using System.Configuration;
using System.Xml;
using System.IO;

using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.VersionControl.Client;

using System.Threading;
using System.Windows.Automation;

using TestTools;
#endregion

namespace TFS2010AutoDeploymentTool
{
    public partial class ToolsForm : Form
    {
        #region Fields of main form
        //public static DispatcherTimer dispatcherTimer = new DispatcherTimer();
        public static DateTime sStartUpTime;
        public Tester tester = new Tester();
        public string TFS2010testToolVersion = "4.12.05.20";
        internal static TestTools.Logger logger = null;
        internal static TimeSpan sTickTime = new TimeSpan(0, 0, 15);
        internal static TimeSpan sTickTime2Min = new TimeSpan(0, 2, 59);
        public static int sTickCount = 0;
        public static int sTickInterval = 0;
        public static  List<string> sSelectedBuildDefs = null;
        public static string sDateFilter = "Today";
        public static bool autoStartup = false;
        public static bool autoPressStartButton = false;

        public static bool sVMSwitchMode = false;
        public static bool sRemoteVMstarted = false;

        public static string sCommandReplyFile = Path.Combine(@"C:\EtriccTests", "RemoteReply.txt");

        //private BackgroundWorker worker; // ToDo: try to use this thread to run testing later
        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        public ToolsForm()
        {
            InitializeComponent();
            // initial Tfs settings
            if (cmbProject.SelectedItem.ToString().Equals(Constants.EPIA_4)
                || cmbProject.SelectedItem.ToString().Equals(Constants.ETRICC_5))
            {
                cmbTestApp.Items.Clear();
                cmbTestApp.Items.Add((Constants.EPIA4));
                cmbTestApp.SelectedIndex = 0;

                cmbTargetPlatform.Items.Clear();
                cmbTargetPlatform.Items.Add("AnyCPU");
                cmbTargetPlatform.Items.Add("x86");
                cmbTargetPlatform.Items.Add("AnyCPU+x86");
                cmbTargetPlatform.SelectedIndex = 0;
            }

            sStartUpTime = System.DateTime.Now;
            lbStartTime.Text = sStartUpTime.ToString("yyyyMMdd HH:mm:ss");
            tester.sTestStartUpTime = sStartUpTime;
            tester.m_TestWorkingDirectory = System.IO.Directory.GetCurrentDirectory();
            tester.TESTTOOL_VERSION = TFS2010testToolVersion;
  
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
                    cmbProject.SelectedItem = customSection.Element.TestProject;
                    cmbTestApp.SelectedItem = customSection.Element.TestApp;
                    cmbTargetPlatform.SelectedItem = customSection.Element.TargetPlatform;
                    sDateFilter = customSection.Element.DateFilter;

                    string allBuildDefsString = customSection.Element.BuildDefinitions;
                    string[] strArray = allBuildDefsString.Split(';');
                    ArrayList arList = new ArrayList();
                    for (int i = 0; i < strArray.Length; i++)
                    {
                        if (strArray[i].Length > 0)
                            lstBoxBuildDefinitions.Items.Add(strArray[i]);
                    }

                    chkProtected.Checked = customSection.Element.BuildProtected;
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

                writeInfo = File.CreateText(Path.Combine(@"C:\EtriccTests", "CommandReply.txt"));
                writeInfo.WriteLine(string.Empty);
                writeInfo.Close();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Create RemoteCommand.txt exception: " + ex.Message);
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
        }

        private void btnConn_Click(object sender, System.EventArgs e)
        {
            try
            {
                //string frompath = @"\\TeamSystem.TeamSystems.Egemin.Be\Team Systems Builds\Nightly\Epia 4\Epia.Development.Dev02.Nightly\Epia.Development.Dev02.Nightly_20120425.1\Installation\Any CPU\Debug";
                //DeployUtilities.CopySetupFilesWithWildcards(frompath, @"C:\Setups", "*.msi", logger);
                //System.IO.File.Copy(frompath, @"C:\Setups\Epiaxxx.mis", true);
                //System.IO.File.Copy(@"C:\Setups\test.txt"
                //            , @"\\TeamSystem.TeamSystems.Egemin.Be\Team Systems Builds\Nightly\Epia 4\Epia.Development.Dev02.Nightly\Epia.Development.Dev02.Nightly_20120425.1\test.txt", true); 

                Uri serverUri = new Uri(Constants.sTFSServerUrl);
                System.Net.ICredentials tfsCredentials
                    = new System.Net.NetworkCredential(Constants.sTFSUsername, Constants.sTFSPassword, Constants.sTFSDomain);

                bool reconnect = false;
                TfsTeamProjectCollection tfsProjectCollection = null;
                while (reconnect == false)
                {
                    reconnect = true;
                    try
                    {
                       tfsProjectCollection = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                    }
                    catch (Exception ex)
                    {
                        DialogResult intResultaat = System.Windows.Forms.MessageBox.Show(ex.Message + " - "+ ex.StackTrace, "reconnect TFSServer", 
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

                //tfsProjectCollection.Connect(Microsoft.TeamFoundation.Framework.Common.ConnectOptions.IncludeServices);
                //TfsConfigurationServer tfsConfigurationServer = tfsProjectCollection.ConfigurationServer;

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

                /*TeamFoundationServer TFS = TeamFoundationServerFactory.GetServer(tester.GetTFSServerName());
                TFS.EnsureAuthenticated();

                if (TFS.Name != null)
                {
                    ICommonStructureService structureService = (ICommonStructureService)TFS.GetService(typeof(ICommonStructureService));
                    ProjectInfo[] projects = structureService.ListAllProjects();

                    string prjs = string.Empty;
                    foreach (ProjectInfo project in projects)
                    {
                        prjs = prjs + "\n " + project.Name;
                    }
                    System.Windows.MessageBox.Show("Connection OK \n Server Name is " + TFS.Name + " \nwith projects:" + prjs, 
                        tester.GetTFSServerName());
                }
                else
                    System.Windows.MessageBox.Show("Connection Failed: ");
                 * */
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message + "" + ex.StackTrace);
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

            string proj = cmbProject.SelectedItem.ToString();
            string testApp = cmbTestApp.SelectedItem.ToString();
            string[] buildDef;
            GetBuildDefinitionsForm connTFSForm = new GetBuildDefinitionsForm(proj, testApp, ref sSelectedBuildDefs, ref sDateFilter);
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
                    customSection.Element.TestProject = cmbProject.SelectedItem.ToString();
                    customSection.Element.TestApp = cmbTestApp.SelectedItem.ToString();
                    customSection.Element.TargetPlatform = cmbTargetPlatform.SelectedItem.ToString();
                    customSection.Element.DateFilter = sDateFilter;
                    
                    string allBuildDefsString = string.Empty;
                    // Populate build definition     
                    foreach (String strThisBuildDef in lstBoxBuildDefinitions.Items)         
                    {
                        if (strThisBuildDef.Length > 0 )
                            allBuildDefsString = strThisBuildDef + ";" + allBuildDefsString;
                    } 
                    customSection.Element.BuildDefinitions = allBuildDefsString;
                    customSection.Element.BuildProtected = chkProtected.Checked;
                      
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
                                    newNode.Attributes["project"].Value = cmbProject.SelectedItem.ToString();
                                    newNode.Attributes["testApp"].Value = cmbTestApp.SelectedItem.ToString();
                                    newNode.Attributes["targetPlatform"].Value = cmbTargetPlatform.SelectedItem.ToString();
                                    newNode.Attributes["dateFilter"].Value = sDateFilter;
                                    newNode.Attributes["buildProtected"].Value = chkProtected.Checked.ToString();

                                    string allBuildDefsString = string.Empty;
                                    for (int i = 0; i < lstBoxBuildDefinitions.Items.Count; i++)
                                    {
                                        allBuildDefsString = lstBoxBuildDefinitions.Items[i].ToString() + ";" + allBuildDefsString;
                                    }
                                    newNode.Attributes["buildDefinitions"].Value = allBuildDefsString; 
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
            logger.LogMessageToFile("------    " + " Tools Version " + TFS2010testToolVersion + " ------- start up --->", 0, 0);
            logger.LogMessageToFile("------    TEST MACHINE: " + System.Environment.MachineName, 0, 0);
            logger.LogMessageToFile("------    OS Version: " + System.Environment.OSVersion, 0, 0);
            //logger.LogMessageToFile("------    Configuration file: " + CONFIGPATH);
            logger.LogMessageToFile("--------------------------------------------------------------------------", 0, 0);
            //}

            //disable start button
            btnStartAuto.Enabled = false;
            btnStopAuto.Enabled = !btnStartAuto.Enabled;

            Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
            Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
            //Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);

            timStart.Enabled = true;
            tester.Log("Searching for available build...");
            lbSelectedConfigFile.Text = "Searching for available build...";
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
            lbSelectedConfigFile.Text = tester.GetCurrentBuildInTesting();

            //tester.Log("timStart_Tick...;  btnStartAuto.Enabled" + btnStartAuto.Enabled);
            if (btnStartAuto.Enabled)
            {
                return;
            }

            TimeSpan mTime = DateTime.Now - sStartUpTime;
            if (mTime.Hours > 1)
                timStart.Interval = 120000;

            //tester.Log("TimStart_Tick mTime.Hours: " + mTime.Hours);
            //logger.LogMessageToFile("TimStart_Tick mTime.Hours: " + mTime.Hours, 0, 0);
               
            System.Diagnostics.Process proc = null;
            int ipid = Utilities.GetApplicationProcessID("Egemin.Epia.Testing.EtriccUIAutoTest", out proc);
            if (ipid > 0)  // check Etricc testing is running
            {
                logger.LogMessageToFile("Egemin.Epia.Testing.EtriccUIAutoTest Test is running", 0, 0);
                if (sVMSwitchMode)
                {
                    StreamWriter write = File.CreateText(sCommandReplyFile);
                    write.WriteLine("TestRunning: Egemin.Epia.Testing.EtriccUIAutoTest Test is running");
                    write.Close();
                }
                return;
            }

            ipid = Utilities.GetApplicationProcessID("Egemin.Epia.Testing.UIAutoTest", out proc);
            if (ipid > 0)  // check Epia testing is running
            {
                logger.LogMessageToFile("Egemin.Epia.Testing.UIAutoTest Test is running", 0, 0);
                return;
            }
            
            ipid = Utilities.GetApplicationProcessID("Egemin.Epia.Testing.KimberlyClarkGUIAutoTest", out proc);
            if (ipid > 0)  // check KC testing is running
            {
                logger.LogMessageToFile("Egemin.Epia.Testing.KImberlyClarkGUIAutoTest Test is running", 0, 0);
                return;
            }

            ipid = Utilities.GetApplicationProcessID("Egemin.Epia.Testing.EtriccStatisticsProgTest", out proc);
            if (ipid > 0)  // check KC testing is running
            {
                logger.LogMessageToFile("Egemin.Epia.Testing.EtriccStatisticsProgTest Test is running", 0, 0);
                return;
            }

            ipid = Utilities.GetApplicationProcessID("Egemin.Epia.Testing.Epia4AppTestProtected", out proc);
            if (ipid > 0)  // check KC testing is running
            {
                logger.LogMessageToFile("Egemin.Epia.Testing.Epia4AppTestProtected Test is running", 0, 0);
                return;
            }
            
            ipid = Utilities.GetApplicationProcessID("EPIA.Explorer", out proc);
            if (ipid > 0)  // check Etricc 5 testing is running
            {
                logger.LogMessageToFile("Etricc 5 Test is running", 0, 0);
                tester.ClickUiScreenActionToAvoidScreenStandBy();
                return;
            }
            
            ipid = Utilities.GetApplicationProcessID(ConstCommon.EGEMIN_EPIA_SERVER, out proc);
            if (ipid > 0)
            {
                logger.LogMessageToFile("Test " + ConstCommon.EGEMIN_EPIA_SERVER + " is running", 0, 0);
                return;
            }
            
            if (tester.GetDeploymentStatus() == true)
            {
                TimeSpan time = DateTime.Now - tester.GetDeploymentEndTime();
                if (time.TotalMilliseconds < 120000)
                {
                    logger.LogMessageToFile("--------- after deploymenttime:" + time.TotalMilliseconds, 0, 0);
                    return;
                }
            }
           
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
                tester.Start(ref sStartUpTime);

                lbStartTime.Text = sStartUpTime.ToString("yyyyMMdd HH:mm:ss");
            }
        }

        private void cmbProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstBoxBuildDefinitions.Items.Clear();
            txtBuildPath.Text = string.Empty;
            if (cmbProject.SelectedItem.ToString().StartsWith("Epia 4"))
            {
                cmbTestApp.Items.Clear();
                cmbTestApp.Items.Add("Epia4");
            }
            else if (cmbProject.SelectedItem.ToString().StartsWith("Etricc 5"))
            {
                cmbTestApp.Items.Clear();
                cmbTestApp.Items.Add("EtriccUI");
                cmbTestApp.Items.Add("Etricc5");
                cmbTestApp.Items.Add("EtriccStatistics");
            }
            else if (cmbProject.SelectedItem.ToString().StartsWith("Epia 3"))
            {
                cmbTestApp.Items.Clear();
                cmbTestApp.Items.Add("Ewms");
            }
            cmbTestApp.SelectedIndex = 0;
            //cmbTestApp.Invalidate();
            sDateFilter = "<Any Time>";
        }

        private void cmbTestApp_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstBoxBuildDefinitions.Items.Clear();
            if (cmbTestApp.SelectedItem.ToString().StartsWith(Constants.EPIA4)
                || cmbTestApp.SelectedItem.ToString().StartsWith(Constants.ETRICCUI))
            {
                cmbTargetPlatform.Items.Clear();
                cmbTargetPlatform.Items.Add("AnyCPU");
                cmbTargetPlatform.Items.Add("x86");
                cmbTargetPlatform.Items.Add("AnyCPU+x86");
            }
            else if ( 
                cmbTestApp.SelectedItem.ToString().StartsWith(Constants.ETRICC5))
            {
                cmbTargetPlatform.Items.Clear();
                cmbTargetPlatform.Items.Add("x86");
            }
            else if (cmbTestApp.SelectedItem.ToString().StartsWith(Constants.ETRICCSTATISTICS))
            {
                cmbTargetPlatform.Items.Clear();
                cmbTargetPlatform.Items.Add("AnyCPU");
                cmbTargetPlatform.Items.Add("x86");
                cmbTargetPlatform.Items.Add("AnyCPU+x86");
            }
            else if (cmbTestApp.SelectedItem.ToString().StartsWith(Constants.EWMS))
            {
                cmbTargetPlatform.Items.Clear();
                cmbTargetPlatform.Items.Add("Undefined");
            }
            cmbTargetPlatform.SelectedIndex = 0;
            //cmbTargetPlatform.Invalidate();
            sDateFilter = "<Any Time>";
        }

        private void mnuAbout_Click(object sender, EventArgs e)
        {
            System.Windows.MessageBox.Show("Egemin E'pia Software Application\n\n E'pia Applications Deployment and Testing Tool\n\n Version:"
                + TFS2010testToolVersion);
        }

        private void mnuExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button_buildSelector_Click(object sender, EventArgs e)
        {
            if (lstBoxBuildDefinitions.Items.Count == 0)
            {
                System.Windows.MessageBox.Show("Please get build definitions first ");
                return;
            }

            string proj = cmbProject.SelectedItem.ToString();
            List<string> selectedBuildDefsList = new List<string>();
            foreach (String strCol in lstBoxBuildDefinitions.Items)
            {
                if (strCol.Length > 0)
                    selectedBuildDefsList.Add(strCol); 
            } 

            string[] buildDef;
            GetBuildNumbersForm getBuildNumbersSForm = new GetBuildNumbersForm(proj, selectedBuildDefsList, sDateFilter, txtBuildPath.Text);
            while (getBuildNumbersSForm.ShowDialog() == DialogResult.OK)
            {
                //lstBoxBuildDefinitions.Items.Clear();
                buildDef = null;
                // extract the data from the dialog
                buildDef = getBuildNumbersSForm.getBuildDefinition();
                //System.Windows.Forms.MessageBox.Show(connTFSForm.Name + " was entered into the database===leng:" + buildDef.Length  );
                //if (buildDef == null)
                //    System.Windows.Forms.MessageBox.Show(connTFSForm.Name + " was entered into the database=== null:");
                if (buildDef.Length == 0)
                    continue;
                else
                {
                    for (int i = 0; i < buildDef.Length; i++)
                    {
                        txtBuildPath.Text = buildDef[i];
                        //lstBoxBuildDefinitions.Items.Add(buildDef[i]);
                    }
                    break;
                }
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
                
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);

                //tester.Start2(this.txtBuildPath.Text, cmbProject.SelectedItem.ToString(), cmbTestApp.SelectedItem.ToString(), cmbTargetPlatform.SelectedItem.ToString(),
                //    null, sDateFilter, ref sStartUpTime, chkProtected.Checked);
                tester.Start2(this.txtBuildPath.Text, ref sStartUpTime);
            }
        }

        private void computerInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string pa = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE");
            string systemBits =  ""+((String.IsNullOrEmpty(pa) || String.Compare(pa, 0, "x86", 0, 3, true) == 0) ? 32 : 64);
            System.Windows.MessageBox.Show("PROCESSOR_ARCHITECTURE:"+pa + System.Environment.NewLine +"system (bit): " + systemBits);
        }

        protected override void OnLoad(EventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show("Enter Load");
            base.OnLoad(e);
            //System.Windows.Forms.MessageBox.Show("Exit Load");
        }

        protected override void OnShown(EventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show("Enter Show");
            base.OnShown(e);
            //System.Windows.Forms.MessageBox.Show("Exit Show");
            //Thread executableThread = new Thread(new ThreadStart(DeployUtilities.StartExecution));
            //executableThread.Start();
            //Thread.Sleep(5000);
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

            StreamReader readerInfo = File.OpenText(remoteCommandFile);
            string info = readerInfo.ReadToEnd();
            readerInfo.Close();

            //if (sVMSwitchMode)
            //{
                //System.Windows.Forms.MessageBox.Show("RemoteCommand.txt is " + info);
                if (info.ToLower().StartsWith("stop"))
                {
                    if (sRemoteVMstarted == true)
                    {
                        timStart.Stop();
                        timStart.Enabled = false;
                        //toggle enabled state of buttons
                        btnStopAuto.Enabled = false;
                        btnStartAuto.Enabled = !btnStopAuto.Enabled;
                        tester.State = Tester.STATE.PENDING;

                        btnStartAuto.Enabled = true;
                        btnStopAuto.Enabled = false;

                        sRemoteVMstarted = false;

                    }
                }
                else if (info.ToLower().StartsWith("start"))
                {
                    //if (sRemoteVMstarted == false)
                    //{
                        if (lstBoxBuildDefinitions.Items.Count == 0)
                        {
                            System.Windows.Forms.MessageBox.Show("Please get build definitions first ", "Egemin AutoDeployment & Test Tool");
                            return;
                        }
                        btnStartAuto.Enabled = false;
                        btnStopAuto.Enabled = true;

                        sRemoteVMstarted = true;
                        SaveTfsSettingsSection();
                        Startup();
                    //}
                }
                else
                {

                }
            //}
        }
    }
}
