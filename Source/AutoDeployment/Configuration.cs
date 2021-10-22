using System;
using System.Windows.Forms;

namespace Epia3Deployment
{
    public partial class Configuration : Form
    {

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Fields of Configuration : Form (1)
        Settings m_settings;
        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Constructors/Destructors/Cleanup of Configuration : Form (2)
        public Configuration()
        {
            InitializeComponent();
        }

        public Configuration(Settings settings)
        {
            InitializeComponent();
            cmbProjectFile.SelectedItem = settings.SelectedProjectFile;
            txtBoxDeployLocation.Text   = settings.Epia3InstallPath;
            cmbBuildApp.SelectedItem    = settings.BuildApplication;
            cmbBranch.SelectedItem      = settings.Branch;
            ckbCI.Checked               = settings.BuildDefCI;
            ckbNightly.Checked          = settings.BuildDefNightly;
            ckbWeekly.Checked           = settings.BuildDefWeekly;
            ckbVersion.Checked          = settings.BuildDefVersion;
            cmbExcelVisible.SelectedItem = settings.ExcelVisible;
            cmbPlatformTarget.SelectedItem = settings.PlatformTarget;
            ckbFunctionalTesting.Checked = settings.FunctionalTesting;
            cmbServerRunAs.SelectedItem  = settings.ServerRunAs;
            ckbMail.Checked              = settings.Mail;
        }
        #endregion // —— Constructors/Destructors/Cleanup •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Methods of Configuration : Form (4)
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            m_settings = new Settings();
            m_settings.SelectedProjectFile  = cmbProjectFile.SelectedItem.ToString();
            m_settings.Epia3InstallPath     = txtBoxDeployLocation.Text;
            m_settings.BuildApplication     = cmbBuildApp.SelectedItem.ToString();
            m_settings.Branch               = cmbBranch.SelectedItem.ToString();
            m_settings.BuildDefCI           = ckbCI.Checked;
            m_settings.BuildDefNightly      = ckbNightly.Checked;
            m_settings.BuildDefWeekly       = ckbWeekly.Checked;
            m_settings.BuildDefVersion      = ckbVersion.Checked;
            m_settings.ExcelVisible         = cmbExcelVisible.SelectedItem.ToString();
            m_settings.PlatformTarget       = cmbPlatformTarget.SelectedItem.ToString();
            m_settings.FunctionalTesting    = ckbFunctionalTesting.Checked;
            m_settings.ServerRunAs          = cmbServerRunAs.SelectedItem.ToString();
            m_settings.Mail                 = ckbMail.Checked;

            string configpath = WpfMainWindow.CONFIGPATH;
            Settings.SaveSettings(m_settings, configpath);
            Close();
        }

        private void button_DeploymentLocation_Click(object sender, EventArgs e)
        {
            // temp disabled, it is defined in install scripts
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                txtBoxDeployLocation.Text = folderBrowserDialog.SelectedPath;
            }
        }


        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

    }
}
