using System;
using System.Collections.Specialized;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace Epia3Deployment
{
    /// <summary>
    /// This class will Load and Save the needed Configuration setting for the Tester Class
    /// </summary>
    /// 
    //[Serializable]
    public class Settings
    {
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Fields of Settings (8)
        private static string root      = "";
        private string m_BuildApp       = "Etricc";
        private string m_Branch         = "Main";
        private bool m_BuildDefCI       = false;
        private bool m_BuildDefNightly  = true;
        private bool m_BuildDefWeekly   = false;
        private bool m_BuildDefVersion  = false;
        private bool    m_DefaultProjectFile    = true;
        private string  m_Epia3InstallPath      = "  ";
        private string  m_ExcelVisible          = "Invisible";
        private string m_PlatformTarget         = "Any CPU";
        private string  m_SelectedProjectFile   = "   ";
        private bool    m_FunctionalTesting     = true;
        private string  m_ServerRunAs           = "Service";
        private bool    m_Mail                  = false;
        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Constructors/Destructors/Cleanup of Settings (1)
        /// <summary>
        /// Default Constructor.
        /// </summary>
        public Settings()
        {
            root = getRootPath();
        }
        #endregion // —— Constructors/Destructors/Cleanup •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Properties of Settings (7)
        [XmlElement]
        public string BuildApplication
        {
            get
            {
                return m_BuildApp;
            }
            set
            {
                m_BuildApp = value;
            }
        }

        [XmlElement]
        public string Branch
        {
            get
            {
                return m_Branch;
            }
            set
            {
                m_Branch = value;
            }
        }

        [XmlElement]
        public bool BuildDefCI
        {
            get
            {
                return m_BuildDefCI;
            }
            set
            {
                m_BuildDefCI = value;
            }
        }

        [XmlElement]
        public bool BuildDefNightly
        {
            get
            {
                return m_BuildDefNightly;
            }
            set
            {
                m_BuildDefNightly = value;
            }
        }

        [XmlElement]
        public bool BuildDefWeekly
        {
            get
            {
                return m_BuildDefWeekly;
            }
            set
            {
                m_BuildDefWeekly = value;
            }
        }

        [XmlElement]
        public bool BuildDefVersion
        {
            get
            {
                return m_BuildDefVersion;
            }
            set
            {
                m_BuildDefVersion = value;
            }
        }

        [XmlElement]
        public bool DefaultProjectFile
        {
            get
            {
                return m_DefaultProjectFile;
            }
            set
            {
                m_DefaultProjectFile = value;
            }
        }

        [XmlElement]
        public string Epia3InstallPath
        {
            get
            {
                return m_Epia3InstallPath;
            }
            set
            {
                m_Epia3InstallPath = value;
            }
        }

        [XmlElement]
        public string ExcelVisible
        {
            get
            {
                return m_ExcelVisible;
            }
            set
            {
                m_ExcelVisible = value;
            }
        }

        [XmlElement]
        public string PlatformTarget
        {
            get
            {
                return m_PlatformTarget;
            }
            set
            {
                m_PlatformTarget = value;
            }
        }

        [XmlElement]
        public string SelectedProjectFile
        {
            get
            {
                return m_SelectedProjectFile;
            }
            set
            {
                m_SelectedProjectFile = value;
            }
        }

        [XmlElement]
        public bool FunctionalTesting
        {
            get
            {
                return m_FunctionalTesting;
            }
            set
            {
                m_FunctionalTesting = value;
            }
        }

        [XmlElement]
        public string ServerRunAs
        {
            get
            {
                return m_ServerRunAs;
            }
            set
            {
                m_ServerRunAs = value;
            }
        }

        [XmlElement]
        public bool Mail
        {
            get
            {
                return m_Mail;
            }
            set
            {
                m_Mail = value;
            }
        }
        #endregion // —— Properties •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••


        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Methods of Settings (5)
        private static string getRootPath()
        {
            string root = Application.StartupPath;
            StringCollection tokens = new StringCollection();
            tokens.AddRange(root.Split(new char[] { '\\' }));
            //Remove last three tokens and form a root path
            string rootPath = "";
            for (int i = 0; i < tokens.Count - 3; i++)
            {
                tokens[i] = tokens[i].Trim();
                rootPath = rootPath + tokens[i] + "\\";
            }
            //MessageBox.Show(rootPath);
            return rootPath;
        }

        /// <summary>
        /// Get the Setting values from a config file
        /// </summary>
        /// <returns>Module Setting with all settings value</returns>
        public static Settings GetSettings()
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Settings));
            Settings settings = null;
            //TODO fixed test.config naam ergens anders halen
            string filePath = System.IO.Path.Combine(getRootPath(), Constants.EPIA3_DEPLOYMENT_CONFIG_FILE/*GetSettingsFile()*/);

            try
            {
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                //De-serialise the data
                settings = (Settings)serializer.Deserialize(fs);
                fs.Close();
            }
            catch (System.IO.FileNotFoundException ex)
            {
                //do something, show exception
                MessageBox.Show(ex.Message);
            }
            return settings;
        }

        public static Settings GetSettings(string configfullpath)
        {
            Settings settings = new Settings(); ;
            try
            {
                //MessageBox.Show("1:" + configfullpath);
                XmlReaderSettings Rsettings = new XmlReaderSettings();
                Rsettings.IgnoreWhitespace = true;
                
                using (XmlReader r = XmlReader.Create(configfullpath, Rsettings))
                {
                    r.MoveToContent();                // Skip over the XML declaration
                    r.ReadStartElement("configuration");

                    while (r.NodeType == XmlNodeType.Element)
                    {
                        r.ReadStartElement("Settings");
                        //MessageBox.Show(r.Name);
                        settings.DefaultProjectFile     = r.ReadElementContentAsBoolean("DefaultProjectFile", "");
                        settings.SelectedProjectFile    = r.ReadElementContentAsString("SelectedProjectFile", "");
                        settings.Epia3InstallPath       = r.ReadElementContentAsString("Epia3InstallPath", "");
                        settings.BuildApplication       = r.ReadElementContentAsString("BuildApplication", "");
                        settings.Branch                 = r.ReadElementContentAsString("Branch", "");
                        settings.BuildDefCI             = r.ReadElementContentAsBoolean("BuildDefCI", "");
                        settings.BuildDefNightly        = r.ReadElementContentAsBoolean("BuildDefNightly", "");
                        settings.BuildDefWeekly         = r.ReadElementContentAsBoolean("BuildDefWeekly", "");
                        settings.BuildDefVersion        = r.ReadElementContentAsBoolean("BuildDefVersion", "");
                        settings.ExcelVisible           = r.ReadElementContentAsString("ExcelVisible", "");
                        settings.PlatformTarget         = r.ReadElementContentAsString("PlatformTarget", "");
                        settings.FunctionalTesting      = r.ReadElementContentAsBoolean("FunctionalTesting", "");
                        settings.ServerRunAs            = r.ReadElementContentAsString("ServerRunAs", "");
                        settings.Mail                   = r.ReadElementContentAsBoolean("Mail", "");
                        
                        //r.MoveToContent();            // Skip over that pesky comment
                        r.ReadEndElement();             // Read the closing customer tag
                    }
                    r.ReadEndElement();

                }

                //XmlSerializer serializer = new XmlSerializer(typeof(Settings));


                //TODO fixed test.config naam ergens anders halen
                //string filePath = System.IO.Path.Combine( Application.StartupPath , "AutoDeploymentTool.config"/*GetSettingsFile()*/);
                //string filepath = configfullpath;

                //FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read);
                //De-serialise the data
                //settings = (Settings)serializer.Deserialize(fs);
                //fs.Close();
            }
            catch (System.IO.FileNotFoundException ex)
            {
                //do something, show exception
                MessageBox.Show(ex.Message);
            }
            catch (System.Exception ex2)
            {
                //do something, show exception
                MessageBox.Show(ex2.Message + "----" + ex2.StackTrace);
            }
            return settings;
        }

        public static void SaveSettings(Settings settings)
        {
            try
            {
                string fileName = System.IO.Path.Combine(Application.StartupPath, Constants.EPIA3_DEPLOYMENT_CONFIG_FILE/*GetSettingsFile()*/);
                //XmlSerializer seriallizer = new XmlSerializer(typeof(Settings));
                //FileStream fs = new FileStream(fileName, FileMode.Create);
                //seriallizer.Serialize(fs, settings);
                //fs.Close();

                XmlWriterSettings Wsettings = new XmlWriterSettings();
                Wsettings.Indent = true;
                using (XmlWriter writer = XmlWriter.Create(fileName, Wsettings))
                {
                    writer.WriteStartElement("configuration");
                    writer.WriteStartElement("Settings");
                    writer.WriteElementString("DefaultProjectFile", settings.DefaultProjectFile.ToString().ToLower());
                    writer.WriteElementString("SelectedProjectFile", settings.SelectedProjectFile);
                    writer.WriteElementString("Epia3InstallPath", settings.Epia3InstallPath);
                    writer.WriteElementString("BuildApplication", settings.BuildApplication);
                    writer.WriteElementString("Branch", settings.Branch);
                    writer.WriteElementString("BuildDefCI", settings.BuildDefCI.ToString().ToLower());
                    writer.WriteElementString("BuildDefNightly", settings.BuildDefNightly.ToString().ToLower());
                    writer.WriteElementString("BuildDefWeekly", settings.BuildDefWeekly.ToString().ToLower());
                    writer.WriteElementString("BuildDefVersion", settings.BuildDefVersion.ToString().ToLower());
                    writer.WriteElementString("ExcelVisible", settings.ExcelVisible);
                    writer.WriteElementString("PlatformTarget", settings.PlatformTarget);
                    writer.WriteElementString("FunctionalTesting", settings.FunctionalTesting.ToString().ToLower());
                    writer.WriteElementString("ServerRunAs", settings.ServerRunAs);
                    writer.WriteElementString("Mail", settings.Mail.ToString().ToLower());
                    
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
                MessageBox.Show(ex.Message + System.Environment.NewLine + ex.StackTrace);
            }
        }

        public static void SaveSettings(Settings settings, string fileName)
        {
            try
            {
                //FileInfo file = new FileInfo(fileName);
                //File.SetAttributes(file.FullName, FileAttributes.Normal);

                //string fileName= System.IO.Path.Combine( Application.StartupPath , "AutoDeploymentTool.config"/*GetSettingsFile()*/);
                //XmlSerializer seriallizer = new XmlSerializer(typeof(Settings));
                //FileStream fs = new FileStream(fileName, FileMode.Create);
                //seriallizer.Serialize(fs, settings);
                //fs.Close();

                XmlWriterSettings Wsettings = new XmlWriterSettings();
                Wsettings.Indent = true;
                using (XmlWriter writer = XmlWriter.Create(fileName, Wsettings))
                {
                    writer.WriteStartElement("configuration");
                    writer.WriteStartElement("Settings");
                    writer.WriteElementString("DefaultProjectFile", settings.DefaultProjectFile.ToString().ToLower());
                    writer.WriteElementString("SelectedProjectFile", settings.SelectedProjectFile);
                    writer.WriteElementString("Epia3InstallPath", settings.Epia3InstallPath);
                    writer.WriteElementString("BuildApplication", settings.BuildApplication);
                    writer.WriteElementString("Branch", settings.Branch);
                    writer.WriteElementString("BuildDefCI", settings.BuildDefCI.ToString().ToLower());
                    writer.WriteElementString("BuildDefNightly", settings.BuildDefNightly.ToString().ToLower());
                    writer.WriteElementString("BuildDefWeekly", settings.BuildDefWeekly.ToString().ToLower());
                    writer.WriteElementString("BuildDefVersion", settings.BuildDefVersion.ToString().ToLower());
                    writer.WriteElementString("ExcelVisible", settings.ExcelVisible);
                    writer.WriteElementString("PlatformTarget", settings.PlatformTarget);
                    writer.WriteElementString("FunctionalTesting", settings.FunctionalTesting.ToString().ToLower());
                    writer.WriteElementString("ServerRunAs", settings.ServerRunAs);
                    writer.WriteElementString("Mail", settings.Mail.ToString().ToLower());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
                MessageBox.Show(ex.Message + System.Environment.NewLine + ex.StackTrace);
            }
        }
        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

    }
}