using System;
using System.Configuration;
using System.Windows.Forms;
using System.Xml;

namespace TFSQATestTools
{
    public partial class Configuration : Form
    {

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Constructors/Destructors/Cleanup of Configuration : Form (2)
        public Configuration()
        {
            InitializeComponent();
        }

        public Configuration(TestsConfigSection customSection)
        {
            InitializeComponent();
            cmbProjectFile.SelectedItem = customSection.Element.SelectedProjectFile;
            cmbExcelVisible.SelectedItem = customSection.Element.Excel;
            ckbFunctionalTesting.Checked = customSection.Element.FunctionalTesting;
            cmbServerRunAs.SelectedItem =  customSection.Element.ServerRunAs;
            ckbMail.Checked = customSection.Element.Mail;
            ckbRemoteVMSwitchMode.Checked = customSection.Element.VMSwitchMode;
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
            SaveTestConfigSectionSettings();
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

        private void SaveTestConfigSectionSettings()
        {
            string sectionName = Constants.TestConfigSection;
            System.Configuration.Configuration config =
                  System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            try
            {
                // Create a custom configuration section having the same name used in the roaming configuration file.
                // This is possible because the configuration section can be overridden by other configuration files. 
                TestsConfigSection customSection = new TestsConfigSection();
                if (config.Sections[sectionName] == null)
                {
                    // Store console settings.
                    customSection.Element.SelectedProjectFile = cmbProjectFile.SelectedItem.ToString();
                    customSection.Element.Excel = cmbExcelVisible.SelectedItem.ToString();
                    customSection.Element.FunctionalTesting = ckbFunctionalTesting.Checked;
                    customSection.Element.ServerRunAs = cmbServerRunAs.SelectedItem.ToString();
                    customSection.Element.Mail = ckbMail.Checked;
                    customSection.Element.VMSwitchMode = ckbRemoteVMSwitchMode.Checked;

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
                                if (node.Name.Equals("testElement"))
                                {
                                    XmlNode pNode = node.ParentNode;
                                    XmlNode oldNode = node;
                                    XmlNode newNode = node;
                                    newNode.Attributes["selectedProjectFile"].Value = cmbProjectFile.SelectedItem.ToString();
                                    newNode.Attributes["excel"].Value = cmbExcelVisible.SelectedItem.ToString();
                                    newNode.Attributes["functionalTesting"].Value = ckbFunctionalTesting.Checked.ToString();
                                    newNode.Attributes["serverRunAs"].Value = cmbServerRunAs.SelectedItem.ToString();
                                    newNode.Attributes["mail"].Value = ckbMail.Checked.ToString();
                                    //newNode.Attributes["vmSwitchMode"].Value = ckbRemoteVMSwitchMode.Checked.ToString();
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
        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


    }
}
