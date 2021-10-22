using System;
using System.Configuration;

namespace TFS2010AutoDeploymentTool
{
    public class TestsConfigSection : ConfigurationSection
    {
        // Create a configuration section.
        public TestsConfigSection()
        { }

        // Set or get the TestConfigElement. 
        [ConfigurationProperty("testElement")]
        public TestConfigElement Element
        {
            get
            {
                return (
                  (TestConfigElement)this["testElement"]);
            }
            set
            {
                this["testElement"] = value;
            }
        }
    }

    // Define a custom configuration element to be 
    // contained by the ConsoleSection. This element 
    // stores background and foreground colors that
    // the application applies to the console window.
    public class TestConfigElement : ConfigurationElement
    {
        // Create the element.
        public TestConfigElement()
        { }

        // Create the element.
        public TestConfigElement(string thisServerRunAs,
            string thisExcel, bool thisFunctionalTesting, bool thisMail, bool thisVMSwitchMode, 
            string thisSelectedProjectFile, bool thisInstallOldEtriccService,
            bool thisInstallEtriccLauncher)
        {
            ServerRunAs = thisServerRunAs;
            Excel = thisExcel;
            FunctionalTesting = thisFunctionalTesting;
            Mail = thisMail;
            VMSwitchMode = thisVMSwitchMode;
            SelectedProjectFile = thisSelectedProjectFile;
            InstallOldEtriccService = thisInstallOldEtriccService;
            InstallEtriccLauncher = thisInstallEtriccLauncher;
        }

        // Get or set the server run as param: Service or Console.
        [ConfigurationProperty("serverRunAs", DefaultValue = "Service", IsRequired = false)]
        public string ServerRunAs
        {
            get
            {
                return (string)this["serverRunAs"];
            }
            set
            {
                this["serverRunAs"] = value;
            }
        }

        // Get or set the excel Visible or Invisible
        [ConfigurationProperty("excel", DefaultValue = "Visible", IsRequired = false)]
        public string Excel
        {
            get
            {
                return (string)this["excel"];
            }
            set
            {
                this["excel"] = value;
            }
        }

        // Get or set Allow Functional Testing true or false
        [ConfigurationProperty("functionalTesting", DefaultValue = false, IsRequired = false)]
        public bool FunctionalTesting
        {
            get
            {
                return (bool)this["functionalTesting"];
            }
            set
            {
                this["functionalTesting"] = value;
            }
        }

        // Get or set Send mail after testing?  true or false
        [ConfigurationProperty("mail", DefaultValue = false, IsRequired = false)]
        public bool Mail
        {
            get
            {
                return (bool)this["mail"];
            }
            set
            {
                this["mail"] = value;
            }
        }

        // Get or set VM Switch Mode?  true or false
        [ConfigurationProperty("vmSwitchMode", DefaultValue = false, IsRequired = false)]
        public bool VMSwitchMode
        {
            get
            {
                return (bool)this["vmSwitchMode"];
            }
            set
            {
                this["vmSwitchMode"] = value;
            }
        }




        // Get or set the console buildDefinitions.
        [ConfigurationProperty("selectedProjectFile", DefaultValue = "Demo.xml", IsRequired = false)]
        public string SelectedProjectFile
        {
            get
            {
                return (string)this["selectedProjectFile"];
            }
            set
            {
                this["selectedProjectFile"] = value;
            }
        }

        // Get or set Install old etricc 5 serice true or false
        [ConfigurationProperty("installOldEtricc5Service", DefaultValue = false, IsRequired = false)]
        public bool InstallOldEtriccService
        {
            get
            {
                return (bool)this["installOldEtricc5Service"];
            }
            set
            {
                this["installOldEtricc5Service"] = value;
            }
        }

        // Get or set Install old etricc 5 launcher true or false
        [ConfigurationProperty("installEtriccLauncher", DefaultValue = false, IsRequired = false)]
        public bool InstallEtriccLauncher
        {
            get
            {
                return (bool)this["installEtriccLauncher"];
            }
            set
            {
                this["installEtriccLauncher"] = value;
            }
        }


    }
}
