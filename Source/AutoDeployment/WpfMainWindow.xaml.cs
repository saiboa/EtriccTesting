using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using System.Windows.Threading;
using System.Windows.Forms;

using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.VersionControl.Client;

using TestTools;

namespace Epia3Deployment
{
    /// <summary>
    /// Interaction logic for WpfMainWindow.xaml
    /// </summary>
    public partial class WpfMainWindow : Window
    {
        #region Fields of WpfMainWindow
        public static DispatcherTimer dispatcherTimer = new DispatcherTimer();
        public static string CONFIGPATH = string.Empty;
        public static DateTime sStartUpTime;
        public Tester tester = new Tester();
        public string testToolVersion = "2.11.08.08";
        internal static TestTools.Logger logger = null;
        internal static TimeSpan sTickTime = new TimeSpan(0, 0, 15);
        internal static TimeSpan sTickTime2Min = new TimeSpan(0, 2, 59);

        public static int sTickCount    = 0;
        public static int sTickInterval = 0;


        //private BackgroundWorker worker; // ToDo: try to use this thread to run testing later
        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        public WpfMainWindow()
        {
            InitializeComponent();
            sStartUpTime = System.DateTime.Now;
            lbStartTime.Content = sStartUpTime.ToString("yyyyMMdd HH:mm:ss");
            tester.sTestStartUpTime = sStartUpTime;
            tester.m_TestWorkingDirectory = System.IO.Directory.GetCurrentDirectory();
            tester.TESTTOOL_VERSION = testToolVersion;
            // --> Load 
            lbSelectedConfigFile.Content =
                System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                Constants.EPIA3_DEPLOYMENT_CONFIG_FILE);

            string sLogFilePath = tester.getLogPath();
            logger = new TestTools.Logger(sLogFilePath);
            logger.LogMessageToFile("Start Log: " + sLogFilePath, 0, 0);

            CONFIGPATH = lbSelectedConfigFile.Content.ToString();
            //dispatcherTimer.Enabled = true;
            tester.OnLoggingChanged += new EventHandler(tester_OnLoggingChanged);

            //DispatcherTimer dispatcherTimer = new DispatcherTimer();
            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            dispatcherTimer.Interval = sTickTime;
            dispatcherTimer.Start();

        }

        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            sTickCount++;
            if (btnStartAuto.IsEnabled)
                return;

            TimeSpan mTime = DateTime.Now - sStartUpTime;
            if (mTime.Minutes > 5)
                sTickInterval = 4;  // 1 min
            else if (mTime.Hours > 1)
                sTickInterval = 12; // 3 min
            else if (mTime.Hours > 5)
                sTickInterval = 20; // 5 min
            
            logger.LogMessageToFile("dispatcherTimer_Tick: " + mTime.Hours, 0, 0);

            System.Diagnostics.Process proc = null;
            int ipid = Utilities.GetApplicationProcessID("Egemin.Epia.Testing.EtriccUIAutoTest", out proc);
            if (ipid > 0)  // check Etricc testing is running
            {
                logger.LogMessageToFile("Egemin.Epia.Testing.EtriccUIAutoTest Test is running", 0, 0);
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
                if ((sTickInterval == 0) || sTickCount % (sTickInterval) == 0)
                {
                    logger.LogMessageToFile("--------- tick time:", 0, 0);
                    
                    tester.SetDeploymentStatus(false);
                    tester.Start(ref sStartUpTime);
                    lbStartTime.Content = sStartUpTime.ToString("yyyyMMdd HH:mm:ss");
                }
            }
        }

        private void btnStartAuto_Click(object sender, RoutedEventArgs e)
        {
            Startup();
        }

        private void btnStopAuto_Click(object sender, RoutedEventArgs e)
        {
            dispatcherTimer.Stop();
            dispatcherTimer.IsEnabled = false;
            //toggle enabled state of buttons
            btnStopAuto.IsEnabled = false;
            btnStartAuto.IsEnabled = !btnStopAuto.IsEnabled;
            //save the configuration
            tester.SaveConfiguration();
            tester.State = Tester.STATE.PENDING;
        }

        private void btnConfig_Click(object sender, RoutedEventArgs e)
        {
            string path = lbSelectedConfigFile.Content.ToString();
            Configuration configScreen = new Configuration(Settings.GetSettings(path));
            configScreen.ShowDialog();
        }

        private void btnStartManual_Click(object sender, RoutedEventArgs e)
        {
            if (!chkContinueAuto.IsChecked.Value)
                btnStopAuto_Click(sender, e);

            // Start manual test here
            if (txtBuildPath.Text == string.Empty)
                System.Windows.MessageBox.Show("Please fill in the Build Path");
            else
            {
                tester.LoadConfiguration(lbSelectedConfigFile.Content.ToString());
                tester.m_TestAutoMode = false;

                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);

                tester.Start(this.txtBuildPath.Text, ref sStartUpTime);
            }
        }

        

        private void mnuExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnConn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                /*
                string sInstallScriptsDir = "X:\\Nightly\\Epia 4\\Epia.Development.Dev03.Nightly\\Epia.Development.Dev03.Nightly_20110418.1\\Installation\\Any CPU\\Debug";
                //string sInstallScriptsDir = "X:\\Version\\Etricc 5\\Etricc.Version\\Release 5.6.19 of Etricc.Version_20110729.2\\Installation\\x86\\Debug";
                string sBuildBaseDir = string.Empty;
                string sBuildNr = string.Empty;
                string sTestApp = string.Empty;
                string sBuildDef = string.Empty;
                string sBuildConfig = string.Empty;
                
                GetAllParameters(sInstallScriptsDir,
                            ref sBuildBaseDir, ref sBuildNr, ref sTestApp, ref sBuildDef, ref sBuildConfig);

                System.Windows.MessageBox.Show("sBuildBaseDir: " + sBuildBaseDir, "1Show build nr");

                System.Windows.MessageBox.Show("sBuildNr: " + sBuildNr, "2Show build nr");

                System.Windows.MessageBox.Show("sTestApp: " + sTestApp, "3Show build nr");

                System.Windows.MessageBox.Show("sBuildDef: " + sBuildDef, "4Show build nr");

                System.Windows.MessageBox.Show("sBuildConfig: " + sBuildConfig, "5Show build nr");
                
                */
                /*
                string p = @"C:\Program Files\Egemin\Epia Shell";
                string filenameList = string.Empty;

                String[] files = System.IO.Directory.GetFiles(p, "*.dll", System.IO.SearchOption.AllDirectories);
                String[] files2 = System.IO.Directory.GetFiles(p, "*.exe", System.IO.SearchOption.AllDirectories);


                System.IO.StreamWriter writeWork = System.IO.File.CreateText(@"C:\SigcheckShell.bat");
               
                

                foreach (string filename in files)
                {
                    filenameList = filenameList 

                          + "C:\\Certif\\Sigcheck.exe" + '"'
                    
                        +" -e -s -v "
                        
                        
                        + '"'+ filename + '"' +  System.Environment.NewLine;
                }

                foreach (string filename in files2)
                {
                    filenameList = filenameList

                          + "C:\\Certif\\Sigcheck.exe" + '"'

                        + " -e -s -v "


                        + '"' + filename + '"' + System.Environment.NewLine;
                }

                writeWork.WriteLine(filenameList);
                writeWork.Close();

                System.Windows.MessageBox.Show(filenameList);
                */

                string pa = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE");
                System.Windows.MessageBox.Show( "system (bit): "+ ( (String.IsNullOrEmpty(pa) || String.Compare(pa, 0, "x86",0, 3, true) == 0 ) ? 32 : 64) , "System");

                string serverUrl = "http://team2010app.teamsystems.egemin.be:8080/tfs/Development";
                Uri serverUri = new Uri(serverUrl);
                System.Net.ICredentials tfsCredentials 
                    = new System.Net.NetworkCredential("TfsBuild", "Egemin01", "TeamSystems.Egemin.Be");

                TfsTeamProjectCollection tfsProjectCollection 
                    = new TfsTeamProjectCollection(serverUri, tfsCredentials);
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
            catch (TeamFoundationServerUnauthorizedException ex)
            {
                System.Windows.MessageBox.Show(ex.Message + "" + ex.StackTrace);
            }

        }

        private void btnFolderBrowser_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dlg = new FolderBrowserDialog();
            System.Windows.Interop.HwndSource source = PresentationSource.FromVisual(this)
                as System.Windows.Interop.HwndSource;
            System.Windows.Forms.IWin32Window win = new OldWindow(source.Handle);
            System.Windows.Forms.DialogResult result = dlg.ShowDialog(win);
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                txtBuildPath.Text = dlg.SelectedPath;
            }
        }

        private void Startup()
        {
            //Load the configuration
            tester.LoadConfiguration(lbSelectedConfigFile.Content.ToString());
            //string logPath = tester.Configuration.BuildInformationfilePath;
            //string path = Path.Combine( logPath, tester.Configuration.LogFilename );
            //logger = new Logger(@"C:\Epia3Log.txt");
            //if (tester.Configuration.EnableLog)
            //{
            logger.LogMessageToFile("--------------------------EPIA Application Deployment and Test Tool----------------------", 0, 0);
            logger.LogMessageToFile("------    " + " TestTools Version " + testToolVersion + " ------- start up --->", 0, 0);
            logger.LogMessageToFile("------    TEST MACHINE: " + System.Environment.MachineName, 0, 0);
            logger.LogMessageToFile("------    OS Version: " + System.Environment.OSVersion, 0, 0);
            //logger.LogMessageToFile("------    Configuration file: " + CONFIGPATH);
            logger.LogMessageToFile("--------------------------------------------------------------------------", 0, 0);
            //}

            //disable start button
            btnStartAuto.IsEnabled = false;
            btnStopAuto.IsEnabled = !btnStartAuto.IsEnabled;

            Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
            Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
            //Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);

            dispatcherTimer.IsEnabled = true;
            tester.Log("Searching for available build...");
            dispatcherTimer.Start();
        }

        private delegate void tester_OnLoggingChangedDelegate(object sender, System.EventArgs e);

        private void tester_OnLoggingChanged(object sender, System.EventArgs e)
        {
            // make threadsafe
            if (!this.Dispatcher.CheckAccess())
            {
                this.Dispatcher.BeginInvoke(new tester_OnLoggingChangedDelegate(tester_OnLoggingChanged), new object[] { sender, e });
                return;
            }
            //update the multiline textBox
            string newLog = string.Empty;

            foreach (string logLine in tester.Logging)
                newLog += logLine + Environment.NewLine;

            txtResult.Text = newLog;
        }

        private void selectToolStripMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "Test Config files (*.config*)| *.config*"; // Filter files by extension
            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                //CONFIGPATH_SELECTION = true;
                CONFIGPATH = dlg.FileName;
                lbSelectedConfigFile.Content = dlg.FileName;
                tester.LoadConfiguration(CONFIGPATH);
            }
        }

        private void viewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string configfile = lbSelectedConfigFile.Content.ToString();
            ConfigXMLViewer configXMLscreen = new ConfigXMLViewer(configfile);
            configXMLscreen.Show();
        }

        private void mnuAbout_Click(object sender, EventArgs e)
        {
            System.Windows.MessageBox.Show("Egemin E'pia Software Application\n\n E'pia 4 Deployment and Testing Tool\n\n Version:"
                + testToolVersion);
        }

        
        private class OldWindow : System.Windows.Forms.IWin32Window
        {
            IntPtr _handle;
            public OldWindow(IntPtr handle)
            {
                _handle = handle;
            }

            #region IWin32Window Members
            IntPtr System.Windows.Forms.IWin32Window.Handle
            {
                get { return _handle; }
            }
            #endregion
        }


    }
}
