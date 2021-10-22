using System;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using TestTools;

namespace QATestEtriccUI
{
    internal class EtriccUtilities
    {
        private static string uninstallWindowName = "Programs and Features";
        private static string sgrid = "Folder View";

        public static AutomationElement GetMainWindow(string mainFormId)
        {
            AutomationElement aeWindow = null;
            AutomationElementCollection aeAllWindows = null;
            // find main window
            Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            int k = 0;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while (aeWindow == null && mTime.TotalSeconds <= 120)
            {
                Console.WriteLine("aeWindow[k]=" + k++);
                k++;
                aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                Thread.Sleep(3000);
                for (int i = 0; i < aeAllWindows.Count; i++)
                {
                    try
                    {
                        if (aeAllWindows[i].Current.AutomationId.Equals(mainFormId))
                        {
                            aeWindow = aeAllWindows[i];
                            Console.WriteLine("aeWindow[" + i + "]=" + aeWindow.Current.Name);
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Get window exception: exceptionmessage:"+ex.Message);
                        //throw ex;
                    }
                }
                mTime = DateTime.Now - mStartTime;
            }

            return aeWindow;
        }

        public static AutomationElement GetMainWindow(string mainFormId, int seconds)
        {
            AutomationElement aeWindow = null;
            AutomationElementCollection aeAllWindows = null;
            // find main window
            Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            int k = 0;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while (aeWindow == null && mTime.TotalSeconds <= seconds)
            {
                Console.WriteLine("aeWindow[k]=" + k++);
                k++;
                aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                Thread.Sleep(3000);
                for (int i = 0; i < aeAllWindows.Count; i++)
                {
                    try
                    {
                        if (aeAllWindows[i].Current.AutomationId.Equals(mainFormId))
                        {
                            aeWindow = aeAllWindows[i];
                            Console.WriteLine("aeWindow[" + i + "]=" + aeWindow.Current.Name);
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Get window exception: exceptionmessage:" + ex.Message);
                        //throw ex;
                    }
                }
                mTime = DateTime.Now - mStartTime;
            }

            return aeWindow;
        }

        public static AutomationElement WaitUntilEtriccServiceModalWindow(AutomationElement EtriccServerCmdElement)
        {
            AutomationElement aeServiceModalWindow = null;
            AutomationElementCollection aeAllWindows = null;
            // find main window
            Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            int k = 0;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while (aeServiceModalWindow == null && mTime.TotalSeconds <= 120)
            {
                Console.WriteLine("aeWindow[k]=" + k++);
                aeAllWindows = EtriccServerCmdElement.FindAll(TreeScope.Descendants, cWindows);
                Console.WriteLine("aeAllWindows.Count=" + aeAllWindows.Count);
                Thread.Sleep(5000);
                for (int i = 0; i < aeAllWindows.Count; i++)
                {
                    if (aeAllWindows[i].Current.Name.Equals("Etricc Service"))
                    {
                        var wp = aeAllWindows[i].GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                        bool isModal = wp.Current.IsModal;
                        if (isModal)
                        {
                            aeServiceModalWindow = aeAllWindows[i];
                            Console.WriteLine("aeWindow[" + i + "]=" + aeServiceModalWindow.Current.Name);
                            break;
                        }
                    }
                }
                mTime = DateTime.Now - mStartTime;
            }

            return aeServiceModalWindow;
        }

        public static AutomationElement GetCategoryWindow(string windowName, ref string errorMsg)
        {
            AutomationElement aeReportWindow = null;
            bool result = true;

            Console.WriteLine("GetReportWindow:: ");
            Console.WriteLine("GetMainWindow:: ");
            AutomationElement aeWindow = GetMainWindow("MainForm");
            if (aeWindow != null)
            {
                Console.WriteLine("MainWindow found name is: " + aeWindow.Current.Name);
                Thread.Sleep(3000);
                result = true;
            }
            else
            {
                errorMsg = "MainWindow not found : ";
                Console.WriteLine(errorMsg);
                Console.WriteLine(errorMsg);
                result = false;
            }

            if (result)
            {
                // find report window
                /*
                Console.WriteLine("find report window: ");
                System.Windows.Automation.Condition cWindows = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.Window);
                
                int k = 0;
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                while (aeReportWindow == null && mTime.TotalSeconds <= 120)
                {
                    Console.WriteLine("aeReportWindow[k]=");
                    k++;
                    aeAllWindows = aeWindow.FindAll(TreeScope.Descendants, cWindows);
                    Console.WriteLine("aeAllWindows.Count=" + aeAllWindows.Count);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        if (aeAllWindows[i].Current.Name.Equals(reportWindowName))
                        {
                            aeReportWindow = aeAllWindows[i];
                            Console.WriteLine("aeReportWindow[" + i + "]=" + aeReportWindow.Current.Name);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }*/
                //aeReportWindow = AUIUtilities.FindElementByName(reportWindowName, aeWindow);
                aeReportWindow = AUIUtilities.FindElementByType(ControlType.Window, aeWindow);
                if (aeReportWindow == null)
                {
                    Console.WriteLine("aeReportWindow not found ");
                }
                else
                {
                    Console.WriteLine("aeReportWindow found: " + aeReportWindow.Current.Name);
                }
            }

            return aeReportWindow;
        }

        public static void ErrorWindowHandling(AutomationElement element, ref string errorMsg)
        {
            const string close = "Close";
            string error;
            AutomationElement aeError1 = AUIUtilities.FindElementByID("m_LblCaption", element);
            if (aeError1 == null)
            {
                error = "Error Message Element not Fund";
                Console.WriteLine(error);
                return;
            }
            else
            {
                errorMsg = aeError1.Current.Name;
                Console.WriteLine("aeError is found ------------:");
                AutomationElement aeErrorText = AUIUtilities.FindElementByID("m_LblErrorText", element);
                if (aeErrorText != null)
                {
                    errorMsg = errorMsg + "\n" + aeErrorText.Current.Name;
                }
            }


            Console.WriteLine("Error Msg is ------------:" + errorMsg);

            AutomationElement aeClose = AUIUtilities.FindElementByID(close, element);
            if (aeClose == null)
            {
                error = "Close button element not Found";
                Console.WriteLine(error);
                return;
            }
            else
            {
                Console.WriteLine("aeClose is found ------------:");
            }

            Thread.Sleep(1000);
            var ivp = (InvokePattern) aeClose.GetCurrentPattern(InvokePattern.Pattern);
            ivp.Invoke();
        }

        public static void TryToGetErrorMessageAndCloseErrorScreen(ref string ErrorMsg)
        {
            const string close = "Continue";
            string error;
            AutomationElement aeErrorWindow = GetCategoryWindow("Egemin Shell", ref ErrorMsg);
            if (aeErrorWindow == null)
            {
                error = "Egemin Shell Error Message Window not Fund";
                Console.WriteLine(error);
                return;
            }
            else
            {
                ErrorMsg = "" + aeErrorWindow.Current.Name;
                Console.WriteLine("aeError is found ------------:");
                AutomationElement aeErrorText = AUIUtilities.FindElementByType(ControlType.Text, aeErrorWindow);
                if (aeErrorText != null)
                {
                    ErrorMsg = ErrorMsg + "\n" + aeErrorText.Current.Name;
                }
            }


            Console.WriteLine("Error Msg is ------------:" + ErrorMsg);

            AutomationElement aeClose = AUIUtilities.FindElementByName(close, aeErrorWindow);
            if (aeClose == null)
            {
                error = "Continue button element not Found";
                Console.WriteLine(error);
                return;
            }
            else
            {
                Console.WriteLine("aeContinue is found ------------:");
            }

            Thread.Sleep(1000);
            var ivp = (InvokePattern) aeClose.GetCurrentPattern(InvokePattern.Pattern);
            ivp.Invoke();
        }

        public static bool SwitchLanguageAndFindText(string resourcesFolder, string fileName, ref string errorMsg)
        {
            bool result = true;
            string language = string.Empty;
            if (fileName.IndexOf("_cn") > 0)
                language = "中文(简体)";
            else if (fileName.IndexOf("_de") > 0)
                language = "Deutsch";
            else if (fileName.IndexOf("_el") > 0)
                language = "Eλληνικά";
            else if (fileName.IndexOf("_en") > 0)
                language = "English";
            else if (fileName.IndexOf("_es") > 0)
                language = "Español";
            else if (fileName.IndexOf("_fr") > 0)
                language = "Français ";
            else if (fileName.IndexOf("_nl") > 0)
                language = "Nederlands";
            else if (fileName.IndexOf("_pl") > 0)
                language = "Polski";
            else
            {
                language = "Extra Language detected";
                errorMsg = "Extra Language detected" + fileName;
                return false;
            }

            AutomationElement aeWindow = null;
            var dirInfo = new DirectoryInfo(resourcesFolder);
            FileInfo[] serverFolderFiles = dirInfo.GetFiles(fileName);
            if (serverFolderFiles.Length == 0)
            {
                result = false;
                errorMsg = resourcesFolder + " has no resource file:" + fileName;
            }
            else // switch to this language
            {
                aeWindow = GetMainWindow("MainForm");
                if (aeWindow == null)
                {
                    result = false;
                    errorMsg = fileName + "SwitchLanguageAndFindText:: Min window noty found ";
                }
                else
                {
                    // open my setting window
                    aeWindow.SetFocus();
                    const string titleBarId = "_MainForm_Toolbars_Dock_Area_Top";
                    AutomationElement aeTitleBar = AUIUtilities.FindElementByID(titleBarId, aeWindow);
                    if (aeTitleBar == null)
                    {
                        result = false;
                        errorMsg = titleBarId + "not found" + fileName;
                    }
                    else
                    {
                        double x = aeTitleBar.Current.BoundingRectangle.Left + 100;
                        double y = (aeTitleBar.Current.BoundingRectangle.Top +
                                    aeTitleBar.Current.BoundingRectangle.Bottom)/2;
                        var myPlacePoint = new Point(x, y);
                        Input.MoveTo(myPlacePoint);
                        Thread.Sleep(2000);

                        Console.WriteLine("re click myPlacePoint :");
                        Input.MoveToAndClick(myPlacePoint);
                        Thread.Sleep(5000);

                        Input.MoveToAndClick(new Point(x, y + 30));
                        Thread.Sleep(5000);
                    }
                }
            }


            AutomationElement aeMySettingsWindow = null;
            if (result)
            {
                const string settingsWindowId = "Dialog - Egemin.Epia.Modules.RnD.Screens.UserSettingsDetailsScreen";
                aeWindow = GetMainWindow("MainForm");
                aeMySettingsWindow = AUIUtilities.FindElementByID(settingsWindowId, aeWindow);
                if (aeMySettingsWindow == null)
                {
                    result = false;
                    errorMsg = fileName + " aeMySettingsWindow not found";
                }
                else
                {
                    Console.WriteLine("aeMySettingsWindow found");
                    //Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeMySettings));
                }
            }


            // change language
            // "中文(简体)", "我的位置"
            if (result)
            {
                AutomationElement aeCombo = AUIUtilities.FindElementByID("languageIdComboBox", aeMySettingsWindow);
                if (aeCombo == null)
                {
                    result = false;
                    errorMsg = fileName + " LanguageSettings failed to find aeCombo at time: " +
                               DateTime.Now.ToString("HH:mm:ss");
                }
                else
                {
                    var selectPattern =
                        aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                    AutomationElement item = null;

                    if (fileName.IndexOf("_cn") > 0)
                        item = AUIUtilities.FindElementByName("中文(简体)", aeCombo); //我的位置
                    else if (fileName.IndexOf("_de") > 0)
                        item = AUIUtilities.FindElementByName("Deutsch", aeCombo); // Meine Einstellungen  
                    else if (fileName.IndexOf("_el") > 0)
                        item = AUIUtilities.FindElementByName("Eλληνικά", aeCombo); // Η τοποθεσία μου             
                    else if (fileName.IndexOf("_en") > 0)
                        item = AUIUtilities.FindElementByName("English", aeCombo); // My Place     
                    else if (fileName.IndexOf("_es") > 0)
                        item = AUIUtilities.FindElementByName("Español", aeCombo); // TEMP    My Place             
                    else if (fileName.IndexOf("_fr") > 0)
                        item = AUIUtilities.FindElementByName("Français ", aeCombo); // Ma place               
                    else if (fileName.IndexOf("_nl") > 0)
                        item = AUIUtilities.FindElementByName("Nederlands", aeCombo); // Mijn plek       
                    else if (fileName.IndexOf("_pl") > 0)
                        item = AUIUtilities.FindElementByName("Polski", aeCombo); // Moje miejsce


                    if (item != null)
                    {
                        Console.WriteLine("LanguageSettings item found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(2000);

                        var itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemPattern.Select();
                    }
                    else
                    {
                        result = false;
                        errorMsg = fileName + " Finding Language in combo failed: " + DateTime.Now.ToString("HH:mm:ss");
                    }
                }
            }

            // save or cancel
            if (result)
            {
                if (AUIUtilities.FindElementAndClickPoint("m_btnSave", aeMySettingsWindow))
                    Thread.Sleep(3000);
                else
                {
                    result = false;
                    errorMsg = fileName + " FindElementAndClick failed:" + "m_btnSave";
                }
                Thread.Sleep(3000);
            }
            else
            {
                if (AUIUtilities.FindElementAndClickPoint("m_btnCancel", aeMySettingsWindow))
                    Thread.Sleep(3000);
                else
                {
                    result = false;
                    errorMsg = fileName + " FindElementAndClick failed:" + "m_btnCancel";
                }

                Thread.Sleep(3000);
            }

            // Validation
            AutomationElement aeValidationText = null;
            int kxx = 0;
            if (result)
            {
                aeWindow = GetMainWindow("MainForm");
                if (aeWindow == null)
                {
                    result = false;
                    errorMsg = fileName + "SwitchLanguageAndFindText:: Main window not found ";
                }
                else
                {
                    if (fileName.IndexOf("_cn") > 0)
                    {
                        while (aeValidationText == null && kxx++ < 5)
                            aeValidationText = AUIUtilities.FindElementByName("现场总线功能", aeWindow); //现场总线功能
                    }
                    else if (fileName.IndexOf("_de") > 0)
                    {
                        while (aeValidationText == null && kxx++ < 5)
                            aeValidationText = AUIUtilities.FindElementByName("Feldfunktionen", aeWindow);
                        // duits    ""Feldfunktionen 
                    }
                    else if (fileName.IndexOf("_el") > 0)
                    {
                        while (aeValidationText == null && kxx++ < 5)
                            aeValidationText = AUIUtilities.FindElementByName("Field functions", aeWindow);
                        // Η τοποθεσία μου             
                    }
                    else if (fileName.IndexOf("_en") > 0)
                    {
                        while (aeValidationText == null && kxx++ < 5)
                            aeValidationText = AUIUtilities.FindElementByName("Field functions", aeWindow); // My Place     
                    }
                    else if (fileName.IndexOf("_es") > 0)
                    {
                        while (aeValidationText == null && kxx++ < 5)
                            aeValidationText = AUIUtilities.FindElementByName("Funciones E/S", aeWindow);
                            // Spanish    "Funciones E/S 
                    }
                    else if (fileName.IndexOf("_fr") > 0)
                    {
                        while (aeValidationText == null && kxx++ < 5)
                            aeValidationText = AUIUtilities.FindElementByName("Fonctions E/S", aeWindow);
                            // Fonctions E/S              
                    }
                    else if (fileName.IndexOf("_nl") > 0)
                    {
                        while (aeValidationText == null && kxx++ < 5)
                            aeValidationText = AUIUtilities.FindElementByName("I/O functies", aeWindow);
                            // I/O functies       
                    }
                    else if (fileName.IndexOf("_pl") > 0)
                    {
                        while (aeValidationText == null && kxx++ < 5)
                            aeValidationText = AUIUtilities.FindElementByName("Field functions", aeWindow); // Moje miejsce
                    }


                    if (aeValidationText == null)
                    {
                        result = false;
                        errorMsg = "translation text 'Field functions' not found in: " + fileName +" ---- kxx:"+kxx;
                    }
                }
            }


            return result;
        }

        public static bool ValidateGridData(AutomationElement aeGrid, string colName1, string val1, string colName2,
                                            string val2, int numRows, ref string errorMSG)
        {
            bool validation = false;
            var agvsIdCells = new string[numRows];
            for (int i = 0; i < numRows; i++)
            {
                // Construct the Grid Cell Element Name
                string cellname = colName1 + " Row " + i;
                // Get the Element with the Row Col Coordinates
                AutomationElement aeCell1 = AUIUtilities.FindElementByName(cellname, aeGrid);

                if (aeCell1 == null)
                {
                    errorMSG = "Find DataGridView aeCell failed:" + cellname;
                    Console.WriteLine(errorMSG);
                    validation = false;
                }
                else
                {
                    Console.WriteLine("cell DataGridView found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                    // find cell value
                    string value1;
                    try
                    {
                        var vp = aeCell1.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        value1 = vp.Current.Value;
                        Console.WriteLine("Get element.Current Value:" + value1);
                    }
                    catch (NullReferenceException)
                    {
                        value1 = string.Empty;
                    }

                    if (string.IsNullOrEmpty(value1))
                    {
                        errorMSG = "DataGridView aeCell Value not found:" + cellname;
                        Console.WriteLine(errorMSG);
                        validation = false;
                    }
                    else if (!value1.Equals(val1))
                    {
                        errorMSG = value1 + " DataGridView aeCell Value not equal:" + val1;
                        Console.WriteLine(errorMSG);
                    }
                    else
                    {
                        #region validate val2

                        string cellname2 = colName2 + " Row " + i;
                        // Get the Element with the Row Col Coordinates
                        AutomationElement aeCell2 = AUIUtilities.FindElementByName(cellname2, aeGrid);
                        if (aeCell1 == null)
                        {
                            errorMSG = "Find DataGridView aeCell failed:" + cellname2;
                            Console.WriteLine(errorMSG);
                            validation = false;
                        }
                        else
                        {
                            // find cell value
                            string value2;
                            try
                            {
                                var vp = aeCell2.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                value2 = vp.Current.Value;
                                Console.WriteLine("Get element.Current Value:" + value2);
                            }
                            catch (NullReferenceException)
                            {
                                value2 = string.Empty;
                            }

                            if (value2 == null || value2 == string.Empty)
                            {
                                errorMSG = "DataGridView aeCell Value not found:" + cellname2;
                                Console.WriteLine(errorMSG);
                                validation = false;
                            }
                            else if (!value2.Equals(val2))
                            {
                                errorMSG = value2 + " DataGridView aeCell Value not equal:" + val2;
                                Console.WriteLine(errorMSG);
                                validation = false;
                            }
                            else
                            {
                                validation = true;
                                errorMSG = string.Empty;
                                i = numRows;
                            }
                        }

                        #endregion
                    }
                }
            }

            return validation;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="root"></param>
        /// <param name="simCommandName"></param>
        /// <param name="errorMsg"></param>
        /// <returns> -1 error, 1 found command in activeList, 0 no command in Active list</returns>
        public static int ValidateAgvSimulation(AutomationElement root, string simCommandName, ref string errorMsg)
        {
            AutomationElement aeBatteryLowBtn = null;

            #region Refind three screens

            AutomationElement aeSimulationWindow = null;
            AutomationElement aeAgvDetailWindow = null;
            AutomationElement aeAgvOverview = null;
            Condition cWindow = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            // Find the simulation element.
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while ((aeSimulationWindow == null || aeAgvOverview == null || aeAgvDetailWindow == null) &&
                   mTime.TotalSeconds < 60)
            {
                AutomationElementCollection aeAllWindows = root.FindAll(TreeScope.Element | TreeScope.Descendants,
                                                                        cWindow);
                for (int i = 0; i < aeAllWindows.Count; i++)
                {
                    Console.WriteLine("All sub window: " + aeAllWindows[i].Current.Name);
                    if (aeAllWindows[i].Current.Name.EndsWith("mulation"))
                        aeSimulationWindow = aeAllWindows[i];
                    else if (aeAllWindows[i].Current.Name.StartsWith("Agv detail"))
                        aeAgvDetailWindow = aeAllWindows[i];
                    else if (aeAllWindows[i].Current.Name.Equals("Agvs"))
                        aeAgvOverview = aeAllWindows[i];
                }

                Thread.Sleep(2000);
                mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time2 is :" + mTime.TotalSeconds);
            }

            if (aeSimulationWindow == null || aeAgvOverview == null || aeAgvDetailWindow == null)
            {
                errorMsg = "aeSimulationWindow or aeAgvOverview or aeAgvDetailWindow not found";
                return -1;
            }

            #endregion

            // test simulation
            Condition c = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, simCommandName),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                );

            // Find the element.
            aeBatteryLowBtn = aeSimulationWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
            if (aeBatteryLowBtn != null)
            {
                Point pt = AUIUtilities.GetElementCenterPoint(aeBatteryLowBtn);
                Input.MoveTo(pt);
                Thread.Sleep(300);
                Input.ClickAtPoint(pt);
                Thread.Sleep(5000);
            }
            else
            {
                errorMsg = "aeBatteryLowBtn Button not Found ------------:";
                return -1;
            }

            AutomationElement aeActiveStatusList = null;
            root = GetMainWindow("MainForm");
            aeAgvDetailWindow = null;

            #region

            // Find Agv DetailWindows
            mStartTime = DateTime.Now;
            mTime = DateTime.Now - mStartTime;
            while (aeAgvDetailWindow == null && mTime.TotalSeconds < 60)
            {
                AutomationElementCollection aeAllWindows = root.FindAll(TreeScope.Descendants, cWindow);
                for (int i = 0; i < aeAllWindows.Count; i++)
                {
                    Console.WriteLine("------   All sub windowS: " + aeAllWindows[i].Current.Name);
                    if (aeAllWindows[i].Current.Name.StartsWith("Agv detail"))
                    {
                        aeAgvDetailWindow = aeAllWindows[i];
                        break;
                    }
                }

                Thread.Sleep(2000);
                mTime = DateTime.Now - mStartTime;
                Console.WriteLine(" time2 is :" + mTime.TotalSeconds);
            }

            if (aeAgvDetailWindow == null)
            {
                errorMsg = "aeAgvDetailWindow not found";
                return -1;
            }
            else
            {
                aeActiveStatusList = AUIUtilities.FindElementByID("activeStatusListWrapperListBox", aeAgvDetailWindow);
                if (aeActiveStatusList == null)
                {
                    errorMsg = "aeActiveStatusList not found";
                    return -1;
                }
                else
                {
                    #region //Get the all the listitems in List control

                    AutomationElementCollection aeAllItems = aeActiveStatusList.FindAll(TreeScope.Children,
                                                                                        new PropertyCondition(
                                                                                            AutomationElement.
                                                                                                ControlTypeProperty,
                                                                                            ControlType.ListItem));

                    bool search = false;
                    for (int i = 0; i < aeAllItems.Count; i++)
                    {
                        Console.WriteLine("aeAllItems[" + i + "]=" + aeAllItems[i].Current.Name);
                        if (aeAllItems[i].Current.Name.Equals("Low battery"))
                        {
                            search = true;
                            break;
                        }
                    }

                    if (search == false)
                    {
                        return 0;
                    }
                    else
                        return 1;

                    #endregion
                }
            }

            #endregion
        }

        public static bool CopyFilesWithWildcards(string fromPath, string toPath, string filenameWithWildcards)
        {
            if (!Directory.Exists(fromPath))
            {
                Directory.CreateDirectory(fromPath);
            }


            if (Directory.Exists(toPath))
            {
                var DirInfo = new DirectoryInfo(toPath);
                FileInfo[] FilesToDelete = DirInfo.GetFiles();

                foreach (FileInfo file in FilesToDelete)
                {
                    try
                    {
                        var attributes = FileAttributes.Normal;
                        File.SetAttributes(file.FullName, attributes);
                        if (file.FullName.IndexOf("Script") > 0)
                            file.Delete();
                    }
                    catch (Exception ex)
                    {
                        //if (m_Settings.EnableLog)
                        //{
                        //string logPath = Configuration.BuildInformationfilePath;
                        Console.WriteLine("Delete file in folder exception  folder : " + toPath);
                        Console.WriteLine("Delete file in folder exception  file : " + file.FullName);
                        Console.WriteLine("Delete file in folder exception: "+ex.Message);
                        return false;
                    }
                }
            }
            else
                Directory.CreateDirectory(toPath);

            FileInfo[] FilesToCopy;
            string filename = string.Empty;
            try
            {
                var DirInfo = new DirectoryInfo(fromPath);
                FilesToCopy = DirInfo.GetFiles(filenameWithWildcards);

                foreach (FileInfo file in FilesToCopy)
                {
                    filename = file.Name;
                    file.CopyTo(Path.Combine(toPath, file.Name));
                }
                //Log("Copied Setup from " + fromPath + " to " + toPath);
                //try
                //{
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    //logger.LogMessageToFile("Copied Setup from " + fromPath + " to " + toPath, 0, 0);
                    //}
                //}
                //catch (Exception ex1)
                //{
                    //if (m_Settings.EnableLog)
                    //logger.LogMessageToFile("------ Test Exception : " + ex1.Message + "\r\n" + ex1.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                //}
            }
            catch (Exception ex)
            {
                //try
                //{
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    //logger.LogMessageToFile("----------Setup Error --------", 0, 0);
                    //logger.LogMessageToFile("CopySetup Exception:" + ex.ToString(), 0, 0);
                    //Log("CopySetup Exception:" + ex.ToString());
                    //}
                //}
                //catch (Exception ex2)
                //{
                    //if (m_Settings.EnableLog)
                    //logger.LogMessageToFile("------ Test Exception : " + ex2.Message + "\r\n" + ex2.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                //}
                Console.WriteLine("Copy file in folder exception  from folder : " + fromPath);
                Console.WriteLine("Copy file in folder exception  file : " + filename);
                Console.WriteLine("Copy file in folder exception: " + ex.Message);
                return false;
            }
            return true;
        }

        public static bool ScrollToThisFolderItemAndDoubleClick(string folderName, ref string sErrorMessage)
        {
            bool status = true;
            string WindowID = "frmMain";
            AutomationElement aeWindow = null;
            AutomationElement aeFolderListView = null;
            AutomationElement aeFolderName = null;
            DateTime mTime = DateTime.Now;
            AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, WindowID, mTime, 300);
            Thread.Sleep(2000);
            AutomationElement aeLoadWindow = AUIUtilities.FindElementByName("Open", aeWindow);
            if (aeLoadWindow == null)
            {
                status = false;
                sErrorMessage = "aeLoadWindow NOT found ------------:";
                Console.WriteLine(sErrorMessage);
            }
            else
            {
                Console.WriteLine("aeLoadWindow FOUND ------------:");
                // Find C:... Treeitem
                Condition cList = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Items View"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List)
                    );

                aeFolderListView = aeLoadWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cList);
                if (aeFolderListView != null)
                {
                    status = true;
                }
                else
                {
                    status = false;
                    sErrorMessage = "aeFolderListView not found ------------:";
                    Console.WriteLine(sErrorMessage);
                }
            }

            // find folderName folder
            if (status)
            {
                #region

                Console.WriteLine("aeListView found  .........");
                Thread.Sleep(1000);
                // Set a property condition that will be used to find the control.
                Condition c = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem);
                AutomationElementCollection aeAllItems = aeFolderListView.FindAll(TreeScope.Children, c);
                //string programFilesFolderName = TestTools.OSVersionInfoClass.ProgramFilesx86FolderName();
                Console.WriteLine("All items count ..." + aeAllItems.Count);
                for (int i = 0; i < aeAllItems.Count; i++)
                {
                    Console.WriteLine("item name: " + aeAllItems[i].Current.Name);
                    if (aeAllItems[i].Current.Name.StartsWith(folderName))
                    {
                        aeFolderName = aeAllItems[i];
                        Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeFolderName));
                        break;
                    }
                }
                Thread.Sleep(2000);
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTimeSpan = DateTime.Now - mStartTime;
                Console.WriteLine(" time is :" + mTimeSpan.TotalSeconds);
                while (aeFolderName == null && mTimeSpan.TotalSeconds <= 120)
                {
                    //ScrollPattern scrollPattern = GetScrollPattern(element);
                    var scrollPattern = (ScrollPattern) aeFolderListView.GetCurrentPattern(ScrollPattern.Pattern);
                    if (scrollPattern.Current.VerticallyScrollable)
                    {
                        scrollPattern.ScrollVertical(ScrollAmount.LargeIncrement);
                    }

                    Thread.Sleep(2000);
                    aeAllItems = aeFolderListView.FindAll(TreeScope.Children, c);
                    Console.WriteLine("FolderName ..." + folderName);
                    for (int i = 0; i < aeAllItems.Count; i++)
                    {
                        Console.WriteLine("item name: " + aeAllItems[i].Current.Name);
                        if (aeAllItems[i].Current.Name.StartsWith(folderName))
                        {
                            Console.WriteLine("FOUND     ..." + aeAllItems[i].Current.Name);
                            aeFolderName = aeAllItems[i];
                            Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeFolderName));
                            break;
                        }
                    }
                }

                if (aeFolderName == null)
                {
                    status = false;
                    sErrorMessage = "FolderName ..." + folderName + "  not found after 2 min ------------:";
                    Console.WriteLine(sErrorMessage);
                }

                #endregion
            }

            return status;
        }

        public static bool ScrollToThisFolderItemAndDoubleClick(string windowID, string folderName,
                                                                ref string sErrorMessage)
        {
            bool status = true;
            //string WindowID = "frmMain";
            AutomationElement aeWindow = null;
            AutomationElement aeFolderListView = null;
            AutomationElement aeFolderName = null;
            DateTime mTime = DateTime.Now;
            AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, windowID, mTime, 300);
            Thread.Sleep(2000);
            AutomationElement aeLoadWindow = AUIUtilities.FindElementByName("Open", aeWindow);
            if (aeLoadWindow == null)
            {
                status = false;
                sErrorMessage = "aeLoadWindow NOT found ------------:";
                Console.WriteLine(sErrorMessage);
            }
            else
            {
                Console.WriteLine("aeLoadWindow FOUND ------------:");
                // Find C:... Treeitem
                Condition cList = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Items View"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List)
                    );

                aeFolderListView = aeLoadWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cList);
                if (aeFolderListView != null)
                {
                    status = true;
                    //ScrollPattern scrollPattern = GetScrollPattern(element);
                    var scrollPattern = (ScrollPattern)aeFolderListView.GetCurrentPattern(ScrollPattern.Pattern);
                    int k = 0;
                    while (scrollPattern.Current.VerticallyScrollable && k++ < 4 )
                    {
                        scrollPattern.ScrollVertical(ScrollAmount.LargeDecrement);
                        scrollPattern = (ScrollPattern)aeFolderListView.GetCurrentPattern(ScrollPattern.Pattern);
                        Thread.Sleep(1000);
                        Console.WriteLine("k= "+k);
                    }
                }
                else
                {
                    status = false;
                    sErrorMessage = "aeFolderListView not found ------------:";
                    Console.WriteLine(sErrorMessage);
                }
            }

            // find folderName folder
            if (status)
            {
                #region

                Console.WriteLine("aeListView found  .........");
                Thread.Sleep(1000);
                // Set a property condition that will be used to find the control.
                Condition c = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.ListItem);

                AutomationElementCollection aeAllItems = aeFolderListView.FindAll(TreeScope.Children, c);
                //string programFilesFolderName = TestTools.OSVersionInfoClass.ProgramFilesx86FolderName();
                Console.WriteLine("All items count ..." + aeAllItems.Count);
                for (int i = 0; i < aeAllItems.Count; i++)
                {
                    Console.WriteLine("item name: " + aeAllItems[i].Current.Name);
                    if (aeAllItems[i].Current.Name.StartsWith(folderName))
                    {
                        aeFolderName = aeAllItems[i];
                        Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeFolderName));
                        break;
                    }
                }
                Thread.Sleep(2000);
                while (aeFolderName == null)
                {
                    //ScrollPattern scrollPattern = GetScrollPattern(element);
                    var scrollPattern = (ScrollPattern) aeFolderListView.GetCurrentPattern(ScrollPattern.Pattern);
                    if (scrollPattern.Current.VerticallyScrollable)
                    {
                        scrollPattern.ScrollVertical(ScrollAmount.LargeIncrement);
                    }

                    Thread.Sleep(2000);
                    aeAllItems = aeFolderListView.FindAll(TreeScope.Children, c);
                    Console.WriteLine("FolderName ..." + folderName);
                    for (int i = 0; i < aeAllItems.Count; i++)
                    {
                        Console.WriteLine("item name: " + aeAllItems[i].Current.Name);
                        if (aeAllItems[i].Current.Name.StartsWith(folderName))
                        {
                            Console.WriteLine("FOUND     ..." + aeAllItems[i].Current.Name);
                            aeFolderName = aeAllItems[i];
                            Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeFolderName));
                            break;
                        }
                    }
                }

                #endregion
            }

            return status;
        }

        public static bool ScrollToThisFolderItemAndDoubleClick(string windowID, string subWindowNamw, string folderName,
                                                                ref string sErrorMessage)
        {
            bool status = true;
            //string WindowID = "frmMain";
            AutomationElement aeWindow = null;
            AutomationElement aeFolderListView = null;
            AutomationElement aeFolderName = null;
            DateTime mTime = DateTime.Now;
            AUIUtilities.WaitUntilElementByIDFound(AutomationElement.RootElement, ref aeWindow, windowID, mTime, 300);
            Thread.Sleep(2000);
            AutomationElement aeLoadWindow = AUIUtilities.FindElementByName(subWindowNamw, aeWindow);
            if (aeLoadWindow == null)
            {
                status = false;
                sErrorMessage = subWindowNamw + " Window NOT found ------------:";
                Console.WriteLine(sErrorMessage);
            }
            else
            {
                Console.WriteLine(subWindowNamw + "aeLoadWindow FOUND ------------:");
                // Find C:... Treeitem
                Condition cList = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "Items View"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List)
                    );

                aeFolderListView = aeLoadWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cList);
                if (aeFolderListView != null)
                {
                    status = true;
                }
                else
                {
                    status = false;
                    sErrorMessage = "aeFolderListView not found ------------:";
                    Console.WriteLine(sErrorMessage);
                }
            }

            // find folderName folder
            if (status)
            {
                #region

                Console.WriteLine("aeListView found  .........");
                Thread.Sleep(1000);
                // Set a property condition that will be used to find the control.
                Condition c = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.ListItem);

                AutomationElementCollection aeAllItems = aeFolderListView.FindAll(TreeScope.Children, c);
                //string programFilesFolderName = TestTools.OSVersionInfoClass.ProgramFilesx86FolderName();
                Console.WriteLine("All items count ..." + aeAllItems.Count);
                for (int i = 0; i < aeAllItems.Count; i++)
                {
                    Console.WriteLine("item name: " + aeAllItems[i].Current.Name);
                    if (aeAllItems[i].Current.Name.StartsWith(folderName))
                    {
                        aeFolderName = aeAllItems[i];
                        Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeFolderName));
                        break;
                    }
                }
                Thread.Sleep(2000);
                while (aeFolderName == null)
                {
                    //ScrollPattern scrollPattern = GetScrollPattern(element);
                    var scrollPattern = (ScrollPattern) aeFolderListView.GetCurrentPattern(ScrollPattern.Pattern);
                    if (scrollPattern.Current.VerticallyScrollable)
                    {
                        scrollPattern.ScrollVertical(ScrollAmount.LargeIncrement);
                    }

                    Thread.Sleep(2000);
                    aeAllItems = aeFolderListView.FindAll(TreeScope.Children, c);
                    Console.WriteLine("FolderName ..." + folderName);
                    for (int i = 0; i < aeAllItems.Count; i++)
                    {
                        Console.WriteLine("item name: " + aeAllItems[i].Current.Name);
                        if (aeAllItems[i].Current.Name.StartsWith(folderName))
                        {
                            Console.WriteLine("FOUND     ..." + aeAllItems[i].Current.Name);
                            aeFolderName = aeAllItems[i];
                            Input.MoveToAndDoubleClick(AUIUtilities.GetElementCenterPoint(aeFolderName));
                            break;
                        }
                    }
                }

                #endregion
            }

            return status;
        }

        public static bool GetSampleProject(TfsTeamProjectCollection tfsProjectCollection, ref string sErrorMessage)
        {
            bool result = true;
            var versionControlServer =
                (VersionControlServer) tfsProjectCollection.GetService(typeof (VersionControlServer));
            //=============
            string workspaceName = Environment.MachineName;
            //string workspaceName = "PCC7 - 201109029";
            string projectPath = @"$/Etricc - Projects/Sample";
            // the container Project (like a tabel in sql/ or like a folder) containing the projects sources in a collection (like a database in sql/ or also like a folder) from TFS          
            string workingDirectory = @"C:\Etricc 5.0.0\Sample";
            // local folder where to save projects sources          
            //TeamFoundationServer tfs = new TeamFoundationServer("http://test-server:8080/tfs/CollectionName", System.Net.CredentialCache.DefaultCredentials); // tfs server url including the  Collection Name --  CollectionName as the existing name of the collection from the tfs server          
            //tfs.EnsureAuthenticated();

            //VersionControlServer sourceControl = (VersionControlServer)tfs.GetService(typeof(VersionControlServer));
            Workspace[] workspaces = versionControlServer.QueryWorkspaces(workspaceName,
                                                                          versionControlServer.AuthorizedUser,
                                                                          Workstation.Current.Name);
            if (workspaces.Length > 0)
            {
                versionControlServer.DeleteWorkspace(workspaceName, versionControlServer.AuthorizedUser);
            }

            Workspace workspace = versionControlServer.CreateWorkspace(workspaceName,
                                                                       versionControlServer.AuthorizedUser,
                                                                       "Temporary Workspace");
            try
            {
                workspace.Map(projectPath, workingDirectory);
                var request = new GetRequest(new ItemSpec(projectPath, RecursionType.Full), VersionSpec.Latest);
                GetStatus status = workspace.Get(request, GetOptions.GetAll | GetOptions.Overwrite);
                // this line doesn't do anything - no failures or errors         
            }
            catch (Exception ex)
            {
                sErrorMessage = ex.Message + "-----" + ex.StackTrace;
                result = false;
                ;
            }
            finally
            {
                if (workspace != null)
                {
                    workspace.Delete();
                    //System.Windows.Forms.MessageBox.Show("The Projects have been brought into the Folder  " + workingDirectory);
                }
            }

            return result;
        }

        public static bool UpdateCreateScriptNoRunApplication(ref string sErrorMessage)
        {
            bool result = true;
            string b = string.Empty;
            string readline = string.Empty;
            StreamReader readerInfo = File.OpenText(@"C:\Etricc 5.0.0\Sample\Source\Script\Main\Create.cs");

            try
            {
                readline = readerInfo.ReadLine();
                while (readline != null)
                {
                    if (readline.IndexOf("Run<Applications>();") >= 0)
                        b = b + "      //Run<Applications>();" + Environment.NewLine;
                    else if (readline.IndexOf("if (!ScriptContext.AutomatedBuild)") >= 0)
                        b = b + "      //if (!ScriptContext.AutomatedBuild)" + Environment.NewLine;
                    else
                        b = b + readline + Environment.NewLine;
                    ;
                    readline = readerInfo.ReadLine();
                }
                readerInfo.Close();

                var attributes = FileAttributes.Normal;
                File.SetAttributes(@"C:\Etricc 5.0.0\Sample\Source\Script\Main\Create.cs", attributes);
                File.Delete(@"C:\Etricc 5.0.0\Sample\Source\Script\Main\Create.cs");


                StreamWriter writeInfo = File.CreateText(@"C:\Etricc 5.0.0\Sample\Source\Script\Main\Create.cs");
                writeInfo.WriteLine(b);
                writeInfo.Close();
            }
            catch (Exception ex)
            {
                sErrorMessage = ex.Message + "-----" + ex.StackTrace;
                result = false;
            }

            return result;
        }

        public static bool ReCompileSampleWorker(TfsTeamProjectCollection tfsProjectCollection, ref string sErrorMessage)
        {
            bool result = true;
            var versionControlServer =
                (VersionControlServer) tfsProjectCollection.GetService(typeof (VersionControlServer));
            //=============
            string workspaceName = Environment.MachineName;
            //string workspaceName = "PCC7 - 201109029";
            string projectPath = @"$/Etricc - Projects/Sample";
            // the container Project (like a tabel in sql/ or like a folder) containing the projects sources in a collection (like a database in sql/ or also like a folder) from TFS          
            string workingDirectory = @"C:\Etricc 5.0.0\Sample";
            // local folder where to save projects sources          
            //TeamFoundationServer tfs = new TeamFoundationServer("http://test-server:8080/tfs/CollectionName", System.Net.CredentialCache.DefaultCredentials); // tfs server url including the  Collection Name --  CollectionName as the existing name of the collection from the tfs server          
            //tfs.EnsureAuthenticated();

            //VersionControlServer sourceControl = (VersionControlServer)tfs.GetService(typeof(VersionControlServer));
            Workspace[] workspaces = versionControlServer.QueryWorkspaces(workspaceName,
                                                                          versionControlServer.AuthorizedUser,
                                                                          Workstation.Current.Name);
            if (workspaces.Length > 0)
            {
                versionControlServer.DeleteWorkspace(workspaceName, versionControlServer.AuthorizedUser);
            }

            Workspace workspace = versionControlServer.CreateWorkspace(workspaceName,
                                                                       versionControlServer.AuthorizedUser,
                                                                       "Temporary Workspace");
            try
            {
                workspace.Map(projectPath, workingDirectory);
                var request = new GetRequest(new ItemSpec(projectPath, RecursionType.Full), VersionSpec.Latest);
                GetStatus status = workspace.Get(request, GetOptions.GetAll | GetOptions.Overwrite);
                // this line doesn't do anything - no failures or errors         
            }
            catch (Exception ex)
            {
                sErrorMessage = ex.Message + "-----" + ex.StackTrace;
                result = false;
                ;
            }
            finally
            {
                if (workspace != null)
                {
                    workspace.Delete();
                    //System.Windows.Forms.MessageBox.Show("The Projects have been brought into the Folder  " + workingDirectory);
                }
            }

            return result;
        }

        //-----------------------------------------------------------------------------------------------------------------------

        public static bool UninstallPragram(string programName, out string msg)
        {
            bool status = false;
            msg = string.Empty;
            AutomationElement element = null;

            Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            AutomationElementCollection aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children,
                                                                                             cWindows);
            Thread.Sleep(3000);
            for (int i = 0; i < aeAllWindows.Count; i++)
            {
                Console.WriteLine("aeWindow[" + i + "]=" + aeAllWindows[i].Current.Name);
                if (aeAllWindows[i].Current.Name.IndexOf("Program") >= 0)
                {
                    element = aeAllWindows[i];
                    break;
                }
            }

            if (element == null) // Programs and Features panel not found:
            {
                msg = "Programs and Features panel not found: ";
                Console.WriteLine(msg);
                Thread.Sleep(3000);
                status = false;
            }

            AutomationElement rootElement = AutomationElement.RootElement;
            //string uninstallWindowName = "Programs and Features";
            string uninstallWindowName = element.Current.Name;
            string sgrid = "Folder View";

            AutomationElement aeEtriccProgram = null;
            //string sYesButtonName = "Yes";
            string sCloseButtonName = "Close";

            #region // Uninstall Etricc Program

            // (1) Programs and Features main window
            Console.WriteLine("Programs and Features Main Form found ... Welcom Main window");
            Console.WriteLine("Searching programs item element...");
            AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(element, sgrid);
            if (aeGridView != null)
                Console.WriteLine("Gridview found...");

            Thread.Sleep(2000);

            // Set a property condition that will be used to find the control.
            Condition c = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.DataItem);

            AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children, c);

            Thread.Sleep(5000);

            Console.WriteLine("Programs count ..." + aeProgram.Count);
            for (int i = 0; i < aeProgram.Count; i++)
            {
                Console.WriteLine("programs name: " + aeProgram[i].Current.Name);
                if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                    && aeProgram[i].Current.Name.IndexOf(programName) > 0)
                    aeEtriccProgram = aeProgram[i];
            }

            if (aeEtriccProgram == null) // Etricc Core not in Programs list
            {
                Console.WriteLine("No Etricc: " + programName);
                AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(element, sCloseButtonName);
                if (btnClose != null)
                {
                    // (2) Components
                    //UnInstalled = true;
                    AUIUtilities.ClickElement(btnClose);
                    status = true;
                }
            }
            else
            {
                Console.WriteLine("Etricc program name: " + aeEtriccProgram.Current.Name);
                string x = aeEtriccProgram.Current.Name;
                AutomationElement dialogElement = null;

                var pattern = aeEtriccProgram.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                pattern.Invoke();
                Thread.Sleep(20000);

                DateTime startTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - startTime;
                while (dialogElement == null && mTime.TotalSeconds < 60)
                {
                    Thread.Sleep(8000);
                    aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        Console.WriteLine("aeWindow[" + i + "].name=" + aeAllWindows[i].Current.Name);
                        Console.WriteLine("aeWindow[" + i + "].automationId=" + aeAllWindows[i].Current.AutomationId);
                        if (aeAllWindows[i].Current.Name.StartsWith("E'tricc") &&
                            aeAllWindows[i].Current.AutomationId.Length < 20)
                        {
                            dialogElement = aeAllWindows[i];
                            break;
                        }
                    }
                    mTime = DateTime.Now - startTime;
                }


                if (dialogElement != null)
                {
                    AutomationElement aeTitleBar =
                        AUIUtilities.FindElementByID("TitleBar", dialogElement);

                    var pt1 = new Point(
                        (aeTitleBar.Current.BoundingRectangle.Left + aeTitleBar.Current.BoundingRectangle.Right)/2,
                        (aeTitleBar.Current.BoundingRectangle.Bottom + aeTitleBar.Current.BoundingRectangle.Top)/2);

                    var newPt1 = new Point(pt1.X - 500, pt1.Y - 400);
                    Input.MoveTo(pt1);

                    Thread.Sleep(1000);
                    Input.SendMouseInput(pt1.X, pt1.Y, 0, SendMouseInputFlags.LeftDown | SendMouseInputFlags.Absolute);

                    Thread.Sleep(1000);
                    Input.SendMouseInput(newPt1.X, newPt1.Y, 0, SendMouseInputFlags.Move | SendMouseInputFlags.Absolute);
                    //Input.MoveTo(newPt1);

                    Thread.Sleep(1000);
                    Input.SendMouseInput(newPt1.X, newPt1.Y, 0,
                                         SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);

                    Console.WriteLine("dlalog element moved =");
                    Thread.Sleep(10000);
                    Console.WriteLine("----- Windows dialog window still open ...");

                    startTime = DateTime.Now;
                    mTime = DateTime.Now - startTime;

                    /*AutomationElement aeRegisterdialog = null;
                    AutomationElement aeYesButton = null;
                    while (aeRegisterdialog == null && mTime.TotalSeconds < 120)
                    {
                        Console.WriteLine("----- WAIT UNTIL Windows RegistryKeysDialog open ...");
                        aeRegisterdialog =
                            AUIUtilities.FindElementByID("FrmRemoveRegistryKeysDialog", AutomationElement.RootElement);

                        if (aeRegisterdialog != null)
                        {
                            Console.WriteLine("aeRegisterdialog dialog found ...");
                            aeYesButton =
                                AUIUtilities.GetElementByNameProperty(rootElement, sYesButtonName);

                            if (aeYesButton != null)
                            {
                                Console.WriteLine("Click Yes Button ...");
                                AUIUtilities.ClickElement(aeYesButton);
                                break;
                            }
                        }
                        else
                            Console.WriteLine("aeRegisterdialog dialog ----  NOT found ...");
                    }*/
                }
                // wait until application uninstalled
                startTime = DateTime.Now;
                mTime = DateTime.Now - startTime;
                bool hasApplication = IsApplicationInstalled(programName);
                while (hasApplication && mTime.TotalSeconds < 360)
                {
                    Thread.Sleep(8000);
                    mTime = DateTime.Now - startTime;
                    if (mTime.TotalSeconds > 300)
                    {
                        MessageBox.Show("Uninstall EtriccCore run timeout " + mTime.TotalSeconds);
                        break;
                    }
                    hasApplication = IsApplicationInstalled("EtriccPlayback");
                }

                Thread.Sleep(2000);
                AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(element, sCloseButtonName);
                if (btnClose != null)
                {
                    AUIUtilities.ClickElement(btnClose);
                    status = true;
                }
            }

            #endregion

            return status;
        }

        public static void Wait(int seconds)
        {
            Thread.Sleep(seconds*1000);
        }

        private static bool IsApplicationInstalled(string ApplicationType)
        {
            bool applicationInstalled = false;

            Condition condition = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName);
            AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
            if (appElement != null)
            {
                // (1) Programs and Features main window
                Console.WriteLine("Programs and Features main window opend");
                Wait(1);
                Console.WriteLine("Searching programs item button...");
                AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(appElement, sgrid);
                if (aeGridView != null)
                    Console.WriteLine("Gridview found...");
                Wait(1);
                // Set a property condition that will be used to find the control.
                Condition c = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.DataItem);

                AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children, c);
                Console.WriteLine("Programs count ..." + aeProgram.Count);
                for (int i = 0; i < aeProgram.Count; i++)
                {
                    switch (ApplicationType)
                    {
                        case "Playback":
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                && aeProgram[i].Current.Name.IndexOf("Playback") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case "Epia":
                            if (aeProgram[i].Current.Name.StartsWith("E'pia Fr"))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case "EtriccCore":
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                && aeProgram[i].Current.Name.IndexOf("Statistics") < 0
                                && aeProgram[i].Current.Name.IndexOf("Shell") < 0
                                && aeProgram[i].Current.Name.IndexOf("Playback") < 0
                                && aeProgram[i].Current.Name.IndexOf("HostTest") < 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case "EtriccShell":
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                && aeProgram[i].Current.Name.IndexOf("Shell") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case "Ewcs":
                            if (aeProgram[i].Current.Name.StartsWith("E'wcs")
                                && aeProgram[i].Current.Name.IndexOf("TestPrograms") < 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case "EwcsTestProgram":
                            if (aeProgram[i].Current.Name.StartsWith("E'wcs")
                                && aeProgram[i].Current.Name.IndexOf("TestPrograms") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                    }
                }
            }


            return applicationInstalled;
        }
    }
}