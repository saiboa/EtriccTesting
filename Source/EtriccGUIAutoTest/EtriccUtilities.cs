using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Automation;
using System.Threading;
using TestTools;

namespace EtriccGUIAutoTest
{
    class EtriccUtilities
    {
        static public AutomationElement GetMainWindow(string mainFormId)
        {
            AutomationElement aeWindow = null;
            AutomationElementCollection aeAllWindows = null;
            // find main window
            System.Windows.Automation.Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            int k = 0;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while (aeWindow == null && mTime.TotalSeconds <= 120)
            {
                Console.WriteLine("aeWindow[k]=");
                k++;
                aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, cWindows);
                Thread.Sleep(3000);
                for (int i = 0; i < aeAllWindows.Count; i++)
                {
                    if (aeAllWindows[i].Current.AutomationId.Equals(mainFormId))
                    {
                        aeWindow = aeAllWindows[i];
                        Console.WriteLine("aeWindow[" + i + "]=" + aeWindow.Current.Name);
                        break;
                    }
                }
                mTime = DateTime.Now - mStartTime;
            }

            return aeWindow;
        }

        static public AutomationElement GetCategoryWindow(string windowName, ref string errorMsg)
        {
            AutomationElement aeReportWindow = null;
            AutomationElementCollection aeAllWindows = null;
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

            if (result == true)
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

        public static void ErrorWindowHandling(AutomationElement element, ref string ErrorMSG)
        {
            string close = "Close";
            string error = string.Empty;
            AutomationElement aeError1 = AUIUtilities.FindElementByID("m_LblCaption", element);
            if (aeError1 == null)
            {
                error = "Error Message Element not Fund";
                Console.WriteLine(error);
                return;
            }
            else
            {
                ErrorMSG = aeError1.Current.Name;
                Console.WriteLine("aeError is found ------------:");
                AutomationElement aeErrorText = AUIUtilities.FindElementByID("m_LblErrorText", element);
                if (aeErrorText != null)
                {
                    ErrorMSG = ErrorMSG + "\n" + aeErrorText.Current.Name;
                }
            }


            Console.WriteLine("Error Msg is ------------:" + ErrorMSG);

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
            InvokePattern ivp = (InvokePattern)aeClose.GetCurrentPattern(InvokePattern.Pattern);
            ivp.Invoke();
        }

        public static void TryToGetErrorMessageAndCloseErrorScreen(ref string ErrorMsg)
        {
            
            string close = "Continue";
            string error = string.Empty;
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
            InvokePattern ivp = (InvokePattern)aeClose.GetCurrentPattern(InvokePattern.Pattern);
            ivp.Invoke();
        }

        static public bool SwitchLanguageAndFindText(string resourcesFolder, string fileName, ref string errorMSG)
        {
            bool result = true;
            string language = "English";
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

            AutomationElement aeWindow = null;
            DirectoryInfo DirInfo = new DirectoryInfo(resourcesFolder);
            FileInfo[] serverFolderFiles = DirInfo.GetFiles(fileName);
            if (serverFolderFiles.Length == 0)
            {
                result = false;
                errorMSG = resourcesFolder + " has no resource file:" + fileName;
            }
            else // switch to this language
            {
                aeWindow = GetMainWindow("MainForm");
                if (aeWindow == null)
                {
                    result = false;
                    errorMSG = fileName + "SwitchLanguageAndFindText:: Min window noty found "; ;
                }
                else
                {   // open my setting window
                    aeWindow.SetFocus();
                    string titleBarID = "_MainForm_Toolbars_Dock_Area_Top";
                    AutomationElement aeTitleBar = AUIUtilities.FindElementByID(titleBarID, aeWindow);
                    if (aeTitleBar == null)
                    {
                        result = false;
                        errorMSG = titleBarID + "not found" + fileName;
                    }
                    else
                    {
                        double x = aeTitleBar.Current.BoundingRectangle.Left + 100;
                        double y = (aeTitleBar.Current.BoundingRectangle.Top + aeTitleBar.Current.BoundingRectangle.Bottom) / 2;
                        System.Windows.Point myPlacePoint = new System.Windows.Point(x, y);
                        Input.MoveTo(myPlacePoint);
                        Thread.Sleep(2000);
                        //while (root.Current.IsEnabled)
                        //{
                        Console.WriteLine("re click myPlacePoint :");
                        Input.MoveToAndClick(myPlacePoint);
                        Thread.Sleep(5000);

                        Input.MoveToAndClick(new System.Windows.Point(x, y + 30));
                        Thread.Sleep(5000);
                    }
                }
            }


            AutomationElement aeMySettingsWindow = null;
            if (result)
            {
                string settingsWindowId = "Dialog - Egemin.Epia.Modules.RnD.Screens.UserSettingsDetailsScreen";
                aeWindow = GetMainWindow("MainForm");
                aeMySettingsWindow = AUIUtilities.FindElementByID(settingsWindowId, aeWindow);
                if (aeMySettingsWindow == null)
                {
                    result = false;
                    errorMSG = fileName + " aeMySettingsWindow not found";
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
                    errorMSG = fileName + " LanguageSettings failed to find aeCombo at time: " + System.DateTime.Now.ToString("HH:mm:ss");
                }
                else
                {
                    SelectionPattern selectPattern =
                        aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                    AutomationElement item = null;

                    if (fileName.IndexOf("_cn") > 0)
                        item = AUIUtilities.FindElementByName("中文(简体)", aeCombo);   //我的位置
                    else if (fileName.IndexOf("_de") > 0)
                        item = AUIUtilities.FindElementByName("Deutsch", aeCombo);        // Meine Einstellungen  
                    else if (fileName.IndexOf("_el") > 0)
                        item = AUIUtilities.FindElementByName("Eλληνικά", aeCombo);      // Η τοποθεσία μου             
                    else if (fileName.IndexOf("_en") > 0)
                        item = AUIUtilities.FindElementByName("English", aeCombo);        // My Place     
                    else if (fileName.IndexOf("_es") > 0)
                        item = AUIUtilities.FindElementByName("Español", aeCombo);       // TEMP    My Place             
                    else if (fileName.IndexOf("_fr") > 0)
                        item = AUIUtilities.FindElementByName("Français ", aeCombo);      // Ma place               
                    else if (fileName.IndexOf("_nl") > 0)
                        item = AUIUtilities.FindElementByName("Nederlands", aeCombo);           // Mijn plek       
                    else if (fileName.IndexOf("_pl") > 0)
                        item = AUIUtilities.FindElementByName("Polski", aeCombo);  // Moje miejsce


                    if (item != null)
                    {
                        Console.WriteLine("LanguageSettings item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(2000);

                        SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemPattern.Select();
                    }
                    else
                    {
                        result = false;
                        errorMSG = fileName + " Finding Language in combo failed: " + System.DateTime.Now.ToString("HH:mm:ss");
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
                    errorMSG = fileName + " FindElementAndClick failed:" + "m_btnSave";
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
                    errorMSG = fileName + " FindElementAndClick failed:" + "m_btnCancel";
                }

                Thread.Sleep(3000);
            }

            // Validation
            AutomationElement aeValidationText = null;
            if (result)
            {
                aeWindow = GetMainWindow("MainForm");
                if (aeWindow == null)
                {
                    result = false;
                    errorMSG = fileName + "SwitchLanguageAndFindText:: Main window not found "; ;
                }
                else
                {
                    if (fileName.IndexOf("_cn") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("现场总线功能", aeWindow);   //现场总线功能
                    else if (fileName.IndexOf("_de") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Field functions", aeWindow);// Meine Einstellungen  
                    else if (fileName.IndexOf("_el") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Field functions", aeWindow);      // Η τοποθεσία μου             
                    else if (fileName.IndexOf("_en") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Field functions", aeWindow);        // My Place     
                    else if (fileName.IndexOf("_es") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Field functions", aeWindow);       // TEMP    My Place             
                    else if (fileName.IndexOf("_fr") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Fonctions E/S", aeWindow);      // Ma place               
                    else if (fileName.IndexOf("_nl") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("I/O functies", aeWindow);           // Mijn plek       
                    else if (fileName.IndexOf("_pl") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Field functions", aeWindow);  // Moje miejsce


                    if (aeValidationText == null)
                    {
                        result = false;
                        errorMSG = "translation text 'Field functions' not found in: " + fileName;
                    }
                }

            }


            return result;
        }

        static public bool ValidateGridData(AutomationElement aeGrid, string colName1, string val1, string colName2, string val2, int numRows, ref string errorMSG)
        {
            bool validation = false;
            string[] AgvsIdCells = new string[numRows];
            for (int i = 0; i < numRows; i++)
            {
                // Construct the Grid Cell Element Name
                string cellname = colName1+" Row " + i;
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
                    Console.WriteLine("cell DataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    // find cell value
                    string Value1 = string.Empty;
                    try
                    {
                        ValuePattern vp = aeCell1.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        Value1 = vp.Current.Value;
                        Console.WriteLine("Get element.Current Value:" + Value1);
                    }
                    catch (System.NullReferenceException)
                    {
                        Value1 = string.Empty;
                    }

                    if (Value1 == null || Value1 == string.Empty)
                    {
                        errorMSG = "DataGridView aeCell Value not found:" + cellname;
                        Console.WriteLine(errorMSG);
                        validation = false;
                    }
                    else if (!Value1.Equals(val1))
                    {
                        errorMSG = Value1+ " DataGridView aeCell Value not equal:" + val1;
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
                            string Value2 = string.Empty;
                            try
                            {
                                ValuePattern vp = aeCell2.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                Value2 = vp.Current.Value;
                                Console.WriteLine("Get element.Current Value:" + Value2);
                            }
                            catch (System.NullReferenceException)
                            {
                                Value2 = string.Empty;
                            }
                        
                            if (Value2 == null || Value2 == string.Empty)
                            {
                                errorMSG = "DataGridView aeCell Value not found:" + cellname2;
                                Console.WriteLine(errorMSG);
                                validation = false;
                            }
                            else if (!Value2.Equals(val2))
                            {
                                 errorMSG = Value2 + " DataGridView aeCell Value not equal:" + val2;
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

        static public bool CopyFilesWithWildcards(string fromPath, string toPath, string filenameWithWildcards)
        {
            if (!Directory.Exists(fromPath))
            {
                Directory.CreateDirectory(fromPath);
            }


            if (Directory.Exists(toPath))
            {
                DirectoryInfo DirInfo = new DirectoryInfo(toPath);
                FileInfo[] FilesToDelete = DirInfo.GetFiles();

                foreach (FileInfo file in FilesToDelete)
                {
                    try
                    {
                        FileAttributes attributes = FileAttributes.Normal;
                        File.SetAttributes(file.FullName, attributes);
                        if (file.FullName.IndexOf("Script") >0)
                            file.Delete();
                    }
                    catch (Exception ex)
                    {
                        //if (m_Settings.EnableLog)
                        //{
                        //string logPath = Configuration.BuildInformationfilePath;
                        //string path = Path.Combine( logPath, Configuration.LogFilename );
                        //Logger logger = new Logger(path );
                        return false;
                    }
                }
            }
            else
                Directory.CreateDirectory(toPath);

            FileInfo[] FilesToCopy;
            try
            {
                DirectoryInfo DirInfo = new DirectoryInfo(fromPath);
                FilesToCopy = DirInfo.GetFiles(filenameWithWildcards);

                foreach (FileInfo file in FilesToCopy)
                {
                    file.CopyTo(Path.Combine(toPath, file.Name));
                }
                //Log("Copied Setup from " + fromPath + " to " + toPath);
                try
                {
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    //logger.LogMessageToFile("Copied Setup from " + fromPath + " to " + toPath, 0, 0);
                    //}
                }
                catch (Exception ex1)
                {
                    //if (m_Settings.EnableLog)
                    //logger.LogMessageToFile("------ Test Exception : " + ex1.Message + "\r\n" + ex1.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                }
            }
            catch (Exception ex)
            {
                try
                {
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    //logger.LogMessageToFile("----------Setup Error --------", 0, 0);
                    //logger.LogMessageToFile("CopySetup Exception:" + ex.ToString(), 0, 0);
                    //Log("CopySetup Exception:" + ex.ToString());
                    //}
                }
                catch (Exception ex2)
                {
                    //if (m_Settings.EnableLog)
                    //logger.LogMessageToFile("------ Test Exception : " + ex2.Message + "\r\n" + ex2.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                }
                //System.Windows.MessageBox.Show("FromPath=" + fromPath + "   " + ex.ToString() + "\r\n" + ex.StackTrace);
                //m_State = Tester.STATE.EXCEPTION;
                return false;
            }
            return true;
        }

    }
}
