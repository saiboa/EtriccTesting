using System;
using System.Windows.Forms;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.Build.Client;
using System.Xml;

using TestTools;


namespace Epia3Deployment
{
    public partial class FormEPIA3Deployment : System.Windows.Forms.Form
    {
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Fields of FormEPIA3Deployment : System.Windows.Forms.Form (3)
        public static string CONFIGPATH = string.Empty;
        public static DateTime sStartUpTime;
        public Tester tester = new Tester();
        public string testToolVersion = "1.10.5.20";
        internal static TestTools.Logger logger = null;
        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Constructors/Destructors/Cleanup of FormEPIA3Deployment : System.Windows.Forms.Form (1)
        public FormEPIA3Deployment()
        {
            InitializeComponent();

            sStartUpTime = System.DateTime.Now;
            lbStartTime.Text = sStartUpTime.ToString("yyyyMMdd HH:mm:ss");
            tester.sTestStartUpTime = sStartUpTime;
            tester.m_TestWorkingDirectory = System.IO.Directory.GetCurrentDirectory();
            tester.TESTTOOL_VERSION = testToolVersion;
        }
        #endregion // —— Constructors/Destructors/Cleanup •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Methods of FormEPIA3Deployment : System.Windows.Forms.Form (16)
   

        [STAThread]
        static void Main(string[] args)
        {
            Utilities.CloseProcess("EXCEL");
            Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
            Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
            //Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);
            Application.Run(new FormEPIA3Deployment());
        }

        private void btnClearLog_Click(object sender, EventArgs e)
        {
            txtResult.Text = String.Empty;
            tester.m_Logging.Clear();
        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            string path = lbSelectedConfigFile.Text;
            Configuration configScreen = new Configuration(Settings.GetSettings(path));
            configScreen.ShowDialog(this);
        }

        private void btnConn_Click(object sender, EventArgs e)
        {
            try
            { 
                TeamFoundationServer TFS = TeamFoundationServerFactory.GetServer("http://teamApplication.teamSystems.egemin.be:8080");
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
                    MessageBox.Show("Connection OK \n Server Name is " + TFS.Name + " \nwith projects:" + prjs);
                }
                else
                    MessageBox.Show("Connection Failed ");
            }
            catch (TeamFoundationServerUnauthorizedException ex)
            {
                MessageBox.Show(ex.Message+""+ex.StackTrace);
            }
        }

        private void btnStartAuto_Click(object sender, EventArgs e)
        {
            Startup();
        }

        private void btnStartManual_Click(object sender, EventArgs e)
        {
            if (!chkContinueAuto.Checked)
                btnStopAuto_Click(sender, e);

            // Start manual test here
            if (txtBuildPath.Text == string.Empty)
                MessageBox.Show("Please fill in the Build Path");
            else
            {
                tester.LoadConfiguration(lbSelectedConfigFile.Text);
                tester.m_TestAutoMode = false;
               
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SHELL);
                Utilities.CloseProcess(ConstCommon.EGEMIN_EPIA_SERVER);
                Utilities.CloseProcess(ConstCommon.EGEMIN_ETRICC_SERVER);

                tester.Start(this.txtBuildPath.Text, ref sStartUpTime);
            }
        }

        private void btnStopAuto_Click(object sender, EventArgs e)
        {
            timStart.Stop();
            timStart.Enabled = false;
            //toggle enabled state of buttons
            btnStopAuto.Enabled = false;
            btnStartAuto.Enabled = !btnStopAuto.Enabled;
            //save the configuration
            tester.SaveConfiguration();
            //if (tester.Configuration.EnableLog)
            //{
            //    logger.LogMessageToFile("auto-deployment setup stoped");
            //    logger.LogMessageToFile("--------------------------------------------------------------------------");
            //}

            tester.State = Tester.STATE.PENDING;
        }

        private void button_FolderBrowser_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                txtBuildPath.Text = folderBrowserDialog.SelectedPath;
            }
        }

        private void FormEPIA3Deployment_Load(object sender, System.EventArgs e)
        {
            lbSelectedConfigFile.Text = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(),
                Constants.EPIA3_DEPLOYMENT_CONFIG_FILE);

            string sLogFilePath = tester.getLogPath();
            logger = new TestTools.Logger(sLogFilePath);
            logger.LogMessageToFile("Start Log: " + sLogFilePath, 0, 0);

            CONFIGPATH = lbSelectedConfigFile.Text;
            timStart.Enabled = true;
            timStart.Start();
            tester.OnLoggingChanged += new EventHandler(tester_OnLoggingChanged);
        }

        private void mnuAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Egemin E'pia Software Application\n\n E'pia 3 Deployment and Testing Tool\n\n Version:" 
                + testToolVersion);
        }

        private void mnuExit_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void Startup()
        {
            //Load the configuration
            tester.LoadConfiguration(lbSelectedConfigFile.Text);
            //string logPath = tester.Configuration.BuildInformationfilePath;
            //string path = Path.Combine( logPath, tester.Configuration.LogFilename );
            //logger = new Logger(@"C:\Epia3Log.txt");
            //if (tester.Configuration.EnableLog)
            //{
            logger.LogMessageToFile("--------------------------EPIA Application Test Tool----------------------", 0, 0);
            logger.LogMessageToFile("------    " + " TestTools Version" + testToolVersion + " ------- start up --->", 0, 0);
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
            timStart.Start();
        }

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

        private void timStart_Tick(object sender, EventArgs e)
        {
            if (btnStartAuto.Enabled)
                return;


            TimeSpan mTime = DateTime.Now - sStartUpTime;
            if (mTime.Hours > 1)
                timStart.Interval = 120000;

            logger.LogMessageToFile("TimStart_Tick mTime.Hours: " + mTime.Hours, 0, 0);
               
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
                tester.SetDeploymentStatus(false);
                tester.Start(ref sStartUpTime);
                lbStartTime.Text = sStartUpTime.ToString("yyyyMMdd HH:mm:ss");
            }
        }
        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        private delegate void tester_OnLoggingChangedDelegate(object sender, System.EventArgs e);

        private void txtBuildPath_DoubleClick(object sender, EventArgs e)
        {
            this.openFileDialog.Filter = "Epia Setup files (*.*)|*.*";
            this.openFileDialog.Title = "Select setup to test";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string st = System.IO.Path.GetDirectoryName(openFileDialog.FileName);
                //string dir = System.IO.Path.GetDirectoryName(st);
                txtBuildPath.Text = st;
            }
        }

        private void selectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.openFileDialog.Filter = "Test Config files (*.config*)| *.config*";
            this.openFileDialog.Title = "Select a configuration file";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //CONFIGPATH_SELECTION = true;
                CONFIGPATH = openFileDialog.FileName;
                lbSelectedConfigFile.Text = openFileDialog.FileName;
                tester.LoadConfiguration(CONFIGPATH);
            }
        }

        private void viewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string configfile = lbSelectedConfigFile.Text;
            ConfigXMLViewer configXMLscreen = new ConfigXMLViewer(configfile);
            configXMLscreen.Show();
        }

    }
}
