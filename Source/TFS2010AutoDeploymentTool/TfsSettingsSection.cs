using System;
using System.Configuration;
using System.Windows.Forms;

namespace TFS2010AutoDeploymentTool
{
    class TfsSettingsSection : ConfigurationSection
    {
        // Create a configuration section.
        public TfsSettingsSection()
        { }

        // Set or get the ConsoleElement. 
        [ConfigurationProperty("settingsElement")]
        public ConsoleConfigElement Element
        {
            get
            {
                return (
                  (ConsoleConfigElement)this["settingsElement"]);
            }
            set
            {
                this["settingsElement"] = value;
            }
        }
    }

    // Define a custom configuration element to be 
    // contained by the ConsoleSection. This element 
    // stores background and foreground colors that
    // the application applies to the console window.
    public class ConsoleConfigElement : ConfigurationElement
    {
        // Create the element.
        public ConsoleConfigElement()
        { }

        // Create the element.
        public ConsoleConfigElement(string thisProject,
            string thisTestApp, string thisTargetPlatform, string thisDateFilter, bool thisBuildProtected)
        {
            TestProject = thisProject;
            TestApp = thisTestApp;
            TargetPlatform = thisTargetPlatform;
            DateFilter = thisDateFilter;
            BuildProtected = thisBuildProtected;

            
        }

        // Get or set the console background color.
        [ConfigurationProperty("project", DefaultValue = "Epia 4", IsRequired = false)]
        public string TestProject
        {
            get
            {
                return (string)this["project"];
            }
            set
            {
                this["project"] = value;
            }
        }

        // Get or set the console testApp.
        [ConfigurationProperty("testApp", DefaultValue = "Epia4", IsRequired = false)]
        public string TestApp
        {
            get
            {
                return (string)this["testApp"];
            }
            set
            {
                this["testApp"] = value;
            }
        }

        // Get or set the console targetPlatform.
        [ConfigurationProperty("targetPlatform", DefaultValue = "AnyCPU", IsRequired = false)]
        public string TargetPlatform
        {
            get
            {
                return (string)this["targetPlatform"];
            }
            set
            {
                this["targetPlatform"] = value;
            }
        }

        // Get or set the console buildDefinitions.
        [ConfigurationProperty("buildDefinitions", DefaultValue = "Epia.Main.CI;Epia.Main.Nightly", 
            IsRequired = false)]
        public string BuildDefinitions
        {
            get
            {
                return (string)this["buildDefinitions"];
            }
            set
            {
                this["buildDefinitions"] = value;
            }
        }

        // Get or set the console dateFilter.
        [ConfigurationProperty("dateFilter", DefaultValue = "<Any Time>",
            IsRequired = false)]
        public string DateFilter
        {
            get
            {
                return (string)this["dateFilter"];
            }
            set
            {
                this["dateFilter"] = value;
            }
        }

        // Get or set the console buildProtected.
        [ConfigurationProperty("buildProtected", DefaultValue = true, IsRequired = false)]
        public bool BuildProtected
        {
            get
            {
                return (bool)this["buildProtected"];
            }
            set
            {
                this["buildProtected"] = value;
            }
        }

    }
}
