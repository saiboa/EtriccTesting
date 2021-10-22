using System.Configuration;

namespace TFSQATestTools
{
    internal class TfsSettingsSection : ConfigurationSection
    {
        // Create a configuration section.

        // Set or get the ConsoleElement. 
        [ConfigurationProperty("settingsElement")]
        public ConsoleConfigElement Element
        {
            get
            {
                return (
                           (ConsoleConfigElement) this["settingsElement"]);
            }
            set { this["settingsElement"] = value; }
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
        {
        }

        // Create the element.
        public ConsoleConfigElement(
            string thisDateFilter, /*bool thisBuildRelease, */
            string thisTestApp1, string thisApp1TestDefName, string thisTestApp2, string thisApp2TestDefName,
            string thisTestApp3, string thisApp3TestDefName)
        {
            DateFilter = thisDateFilter;
            //BuildRelease = thisBuildRelease;

            // new fields
            TestApp1 = thisTestApp1;
            App1TestDefName = thisApp1TestDefName;
            TestApp2 = thisTestApp2;
            App2TestDefName = thisApp2TestDefName;
            TestApp3 = thisTestApp3;
            App3TestDefName = thisApp3TestDefName;
        }

        // Get or set the console background color.
        [ConfigurationProperty("testApplication1", DefaultValue = "Epia4", IsRequired = false)]
        public string TestApp1
        {
            get { return (string) this["testApplication1"]; }
            set { this["testApplication1"] = value; }
        }

        // Get or set the console background color.
        [ConfigurationProperty("application1TestDefName", DefaultValue = "Epia4TestDefinition.txt", IsRequired = false)]
        public string App1TestDefName
        {
            get { return (string) this["application1TestDefName"]; }
            set { this["application1TestDefName"] = value; }
        }

        // Get or set the console background color.
        [ConfigurationProperty("testApplication2", DefaultValue = "EtriccUI", IsRequired = false)]
        public string TestApp2
        {
            get { return (string) this["testApplication2"]; }
            set { this["testApplication2"] = value; }
        }

        // Get or set the console background color.
        [ConfigurationProperty("application2TestDefName", DefaultValue = "EtriccUITestTypeDefinition.txt",
            IsRequired = false)]
        public string App2TestDefName
        {
            get { return (string) this["application2TestDefName"]; }
            set { this["application2TestDefName"] = value; }
        }

        // Get or set the console background color.
        [ConfigurationProperty("testApplication3", DefaultValue = "EtriccStatistics", IsRequired = false)]
        public string TestApp3
        {
            get { return (string) this["testApplication3"]; }
            set { this["testApplication3"] = value; }
        }

        // Get or set the console background color.
        [ConfigurationProperty("application3TestDefName", DefaultValue = "StatisticsTestTypeDefinition.txt",
            IsRequired = false)]
        public string App3TestDefName
        {
            get { return (string) this["application3TestDefName"]; }
            set { this["application3TestDefName"] = value; }
        }

        //========================================================================================================================================
        // Get or set the console buildDefinitions.
        [ConfigurationProperty("buildDefinitions", DefaultValue = "Epia.Main.CI;Epia.Main.Nightly",
            IsRequired = false)]
        public string BuildDefinitions
        {
            get { return (string) this["buildDefinitions"]; }
            set { this["buildDefinitions"] = value; }
        }

        // Get or set the console dateFilter.
        [ConfigurationProperty("dateFilter", DefaultValue = "<Any Time>",
            IsRequired = false)]
        public string DateFilter
        {
            get { return (string) this["dateFilter"]; }
            set { this["dateFilter"] = value; }
        }


        // Get or set the console buildRelease.
        /*[ConfigurationProperty("buildRelease", DefaultValue = true, IsRequired = false)]
        public bool BuildRelease
        {
            get
            {
                return (bool)this["buildRelease"];
            }
            set
            {
                this["buildRelease"] = value;
            }
        }*/
    }
}