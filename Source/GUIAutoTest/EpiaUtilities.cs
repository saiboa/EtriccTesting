using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Automation;
using System.Text;
using System.Linq;
using System.IO;
using System.Collections;
using System.Windows;

using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
//using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.Build.Client;

using TestTools;

namespace Epia4GUIAutoTest
{
    class EpiaUtilities
    {
        public static void ClearDisplayedScreens(AutomationElement root)
        {
            AutomationElementCollection aeAllTabs = root.FindAll(TreeScope.Descendants, new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Tab));

            for (int k = 0; k < aeAllTabs.Count; k++)
            {
                if (aeAllTabs[k] != null)
                {
                    double right = aeAllTabs[k].Current.BoundingRectangle.Right;
                    double bottom = aeAllTabs[k].Current.BoundingRectangle.Bottom;
                    double top = aeAllTabs[k].Current.BoundingRectangle.Top;

                    double x = right - 5;
                    double y = (top + bottom) / 2;
                    Point p = new Point(x, y);

                    for (int i = 0; i < 5; i++)
                    {
                        Input.MoveToAndClick(p);
                        Thread.Sleep(300);
                    }
                }
            }
        }

        /// <summary>
        ///     get searched files
        /// </summary>
        /// <param name="path">folder where files located</param>
        /// <param name="searchPattern">file type sepearted by ; --> "*.exe;*.config;*.dll;*.pdb"</param>
        /// <returns></returns>
        public static string[] GetFiles( string path,  string searchPattern)
        {    
            string[] fileTypes = searchPattern.Split(';');   
            List<string> strFiles = new List<string>();
            foreach (string filter in fileTypes)    
                strFiles.AddRange(System.IO.Directory.GetFiles(path, filter));    

            string[] absolutefileNames = strFiles.ToArray();

            string[] fileNames = new string[absolutefileNames.Length];
            for (int i = 0; i < absolutefileNames.Length; i++)
            {
                FileInfo fileInfo = new FileInfo(absolutefileNames[i]);
                fileNames[i] = fileInfo.Name;
                //Console.WriteLine("installedDllsName[i] : " + installedDllsName[i]);
                //dlls = dlls + System.Environment.NewLine + installedDllsName[i];
            }
            return fileNames;
        }

        public static void CompareFileLists(string[] thisFiles, string[] standardFiles, ref string errorMsg)
        {
            string message = "";
            // Create the query. Note that method syntax must be used here.
            IEnumerable<string> differenceQuery =
                standardFiles.Except(thisFiles);

            IEnumerator EmpEnumerator = differenceQuery.GetEnumerator(); //Getting the Enumerator
            //EmpEnumerator.Reset(); //Position at the Beginning
            while (EmpEnumerator.MoveNext()) //Till not finished do print
            {
                string b = (string)EmpEnumerator.Current;
                Console.WriteLine("The following file are missing  --- " + b);
                message = message + "; " + b;
            }
            errorMsg = message;
            // Execute the query.
            //Console.WriteLine("The following lines are in names1.txt but not names2.txt");
            //foreach (string s in differenceQuery)
            //    Console.WriteLine(s);
        }

        public static void CompareLists()
        {
            // Create the IEnumerable data sources.
            string[] names1 = System.IO.File.ReadAllLines(@"../../../names1.txt");
            string[] names2 = System.IO.File.ReadAllLines(@"../../../names2.txt");

            // Create the query. Note that method syntax must be used here.
            IEnumerable<string> differenceQuery =
                names1.Except(names2);

            IEnumerator EmpEnumerator = differenceQuery.GetEnumerator(); //Getting the Enumerator
            EmpEnumerator.Reset(); //Position at the Beginning
            while (EmpEnumerator.MoveNext()) //Till not finished do print
            {
                string b = (string)EmpEnumerator.Current;
                Console.WriteLine("The following lines are in names1.txt but not names2.txt  --- "+b);
            }

            // Execute the query.
            Console.WriteLine("The following lines are in names1.txt but not names2.txt");
            foreach (string s in differenceQuery)
                Console.WriteLine(s);
        }

        public static bool IsFileListEqual(List<string> ExternalList, List<string> InternalList, ref string difFiles) 
        { 
            if (InternalList.Count != ExternalList.Count) 
            { 
                return false; 
            } 
            else 
            { 
                for (int i = 0; i < InternalList.Count; i++) 
                { 
                    if (InternalList[i] != ExternalList[i])                     
                        return false; 
                } 
            } 
            
            return true; 
        }

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
                try
                {
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
                }
                catch(Exception ex)
                {
                    aeWindow = null;
                    Console.WriteLine(ex.Message +" ---- "+ex.StackTrace);
                    Thread.Sleep(5000);
                }
                mTime = DateTime.Now - mStartTime;
            }

            return aeWindow;
        }

        static public AutomationElement GetSelectedOverviewWindow(string reportWindowName, ref string errorMsg)
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
                // opened window (in test only one window is opend
                aeReportWindow = AUIUtilities.FindElementByType(ControlType.Window, aeWindow);
                if (aeReportWindow == null)
                {
                    Console.WriteLine("aeSelectedtWindow not found ");
                }
                else
                {
                    Console.WriteLine("aeSelectedWindow found: " + aeReportWindow.Current.Name);
                }
            }

            return aeReportWindow;
        }

        static public AutomationElement GetCellElementFromOverviewWindow(AutomationElement aeOverview, string colName, int row, ref string errorMsg)
        {
            AutomationElement aeCell = null;
            #region // Find GridView
            AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeOverview);
            if (aeGrid == null)
            {
                errorMsg = aeOverview.Current.Name+ " DataGridView not found";
            }
            else
            {
                Console.WriteLine("DataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                Thread.Sleep(3000);
                // Construct the Grid Cell Element Name
                string cellname = colName+" Row "+row;
                // Get the Element with the Row Col Coordinates
                aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                if (aeCell == null)
                {
                    errorMsg = "Find DataGridView aeCell failed:" + cellname;
                }
            }
            #endregion

            return aeCell;
        }

        static public bool WindowMenuAction(AutomationElement aeWindow, string colHeaderName, int row, string menuItemName, ref string errorMsg)
        {
            bool result = true;
            #region // Find GridView
            AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeWindow);
            if (aeGrid == null)
            {
                errorMsg = aeWindow.Current.Name + " GridData not found";
                Console.WriteLine(errorMsg);
                result = false;
            }
            else
            {
                Console.WriteLine(aeWindow.Current.Name + " GridData found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                Thread.Sleep(3000);

                // Construct the Grid Cell Element Name
                string cellname = colHeaderName + " Row " + row;
                // Get the Element with the Row Col Coordinates
                AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
               
                if (aeCell == null)
                {
                    errorMsg = "Find aeCell failed:" + cellname;
                    Console.WriteLine(errorMsg);
                    result = false;
                }
                else
                {
                    Console.WriteLine(aeCell.Current.Name+"    -----  cell found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                    Thread.Sleep(3000);
                    System.Windows.Point cellPoint = AUIUtilities.GetElementCenterPoint(aeCell);
                    Input.MoveToAndRightClick(cellPoint);
                    Thread.Sleep(2000);

                    AutomationElement aeMenuItem = GetMenuItemFromElement(aeWindow, menuItemName, 120, ref errorMsg);
                    if (aeMenuItem == null)
                    {
                        errorMsg = aeWindow.Current.Name + " aeMenuItem not found -->  " + menuItemName;
                        Console.WriteLine(errorMsg);
                        result = false;
                    }
                    else
                    {
                        System.Windows.Point menuItemPoint = AUIUtilities.GetElementCenterPoint(aeMenuItem);
                        Input.MoveToAndClick(menuItemPoint);
                        result = true;
                    }
                }                
            }
            #endregion

            return result;
        }

        static public AutomationElement GetMenuItemFromElement(AutomationElement element, string menuItemId, int seconds, ref string errorMsg)
        {
            AutomationElement aeMenuItem = null;
            AutomationElementCollection aeAllMenuItems = null;
            System.Windows.Automation.Condition cMenuItems = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem);

            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            try
            {
                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                while (aeMenuItem == null && mTime.TotalSeconds <= 120)
                {
                    aeAllMenuItems = element.FindAll(TreeScope.Descendants, cMenuItems);
                    Thread.Sleep(3000);
                    Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);
                    for (int i = 0; i < aeAllMenuItems.Count; i++)
                    {
                        Console.WriteLine("aeAllMenuItems[" + i + "]=" + aeAllMenuItems[i].Current.Name);
                        if (aeAllMenuItems[i].Current.Name.Equals(menuItemId))
                        {
                            aeMenuItem = aeAllMenuItems[i];
                            Console.WriteLine("aeAllMenuItems[" + i + "]=" + aeMenuItem.Current.Name);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message + "---------------" + ex.StackTrace;
                Console.WriteLine(errorMsg);
            }

            return aeMenuItem;
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="choice">Shell or Server or Both</param>
        static public bool InstallEpia(string choice, ref string errorMsg)
        {
            AutomationElement rootElement = AutomationElement.RootElement;
            AutomationElement btnNext = null;

            #region Install Epia
            Console.WriteLine("Searching for main installer window");
            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
            AutomationElement appElement = rootElement.FindFirst(TreeScope.Children, condition);

            DateTime startTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - startTime;
            while (appElement == null && mTime.TotalMilliseconds < 60000)
            {
                Wait(2);
                appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds > 60000)
                {
                    errorMsg = "After one minute no Installer Window Form found";
                    System.Windows.Forms.MessageBox.Show("After one minute no Installer Window Form found");
                    return false;
                }
            }

            if (appElement != null)
            {   // (1) Welcom Main window
                Console.WriteLine("Welcom Main window opend ");
                Console.WriteLine("Searching next button...");
                btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                if (btnNext != null)
                {   // (2) Components
                    AUIUtilities.ClickElement(btnNext);
                    appElement = rootElement.FindFirst(TreeScope.Children, condition);
                    Console.WriteLine("Componts window opend");
                    Console.WriteLine("Searching checkbox...");
                    AutomationElement epiaServerCheckbox = AUIUtilities.GetElementByNameProperty(appElement, "E'pia Server");
                    if (epiaServerCheckbox != null && (choice.ToLower().StartsWith("server") || choice.ToLower().StartsWith("both")) )
                    {
                        AUIUtilities.ClickElement(epiaServerCheckbox);
                    }

                    AutomationElement epiaShellCheckbox = AUIUtilities.GetElementByNameProperty(appElement, "E'pia Shell");
                    if (epiaShellCheckbox != null && (choice.ToLower().StartsWith("shell") || choice.ToLower().StartsWith("both")))
                    {
                        AUIUtilities.ClickElement(epiaShellCheckbox);
                    }

                    Wait(2);
                    Console.WriteLine("Searching next button...");
                    btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                    if (btnNext != null)
                    {
                        AUIUtilities.ClickElement(btnNext);

                        appElement = rootElement.FindFirst(TreeScope.Children, condition);
                        if (appElement != null)
                            Console.WriteLine("Installation Folders window is opend");
                        // in the future maybe will edit installation Folder
                        //WaitUntilInstallationComplete(appElement);
                        // wait until isContent close button found
                        Console.WriteLine("wait until Content Close button found");
                        System.Windows.Automation.Condition c2 = new AndCondition(new PropertyCondition(AutomationElement.NameProperty, "Close"),
                                new PropertyCondition(AutomationElement.IsContentElementProperty, true) );

                        AutomationElement aeBtnClose = null;
                        while (aeBtnClose == null)
                        {
                            Console.WriteLine("Wait until Close button found...");
                            appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                            btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                            if (btnNext != null)
                            {
                                Console.WriteLine("Next > button found first --> Click Next > button");
                                if (btnNext.Current.IsKeyboardFocusable)
                                    AUIUtilities.ClickElement(btnNext);
                                else
                                    Console.WriteLine("Next > button IsKeyboardFocusable --> false");
                            }
                            else
                            {
                                aeBtnClose = appElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);
                            }
                            Wait(5);
                        }
                        Console.WriteLine("Close button found... ---> Close Installer Window");
                        AUIUtilities.ClickElement(aeBtnClose);
                        Console.WriteLine("---------- Epia Install Successful ---------");
                    }
                    else
                    {
                        errorMsg = "---------- Next button not found  ---------";
                        Console.WriteLine("---------- Next button not found  ---------");
                        return false;
                    }
                }
            }
            #endregion
            return true;
        }

        static private void Wait(int seconds)
        {
            System.Threading.Thread.Sleep(seconds * 1000);
        }

        static public string GetProgramsFeaturesScreenNaam()
        {
            string screenName = "Programs and Features";
            string MachineName = System.Environment.MachineName;
            if (MachineName.ToUpper().StartsWith("ETRICCSTATAUTO") || MachineName.ToUpper().StartsWith("ETRICCAUTOTEST1"))
                screenName = "Control Panel\\Programs\\Programs and Features";

            return screenName; 
        }

        static public void StartProgramsAndFeaturesExecution()
        {
            System.Diagnostics.Process Proc = new System.Diagnostics.Process();
            Proc.StartInfo.FileName = @"C:\Windows\System32\appwiz.cpl";
            Proc.StartInfo.CreateNoWindow = false;
            Proc.Start();
        }

        static public void StartEpiaResourceFileEditorExecution()
        {
            System.Diagnostics.Process Proc = new System.Diagnostics.Process();
            Proc.StartInfo.FileName = Path.Combine(OSVersionInfoClass.ProgramFilesx86() + "\\Egemin\\Epia Resource File Editor",
                "Egemin.Epia.Foundation.Globalization.ResourceFileEditor.exe");
            Proc.StartInfo.CreateNoWindow = false;
            Proc.Start();
        }

        public static void WaitUntilElementByNameFound(AutomationElement root, ref AutomationElement element, string name,
            DateTime startTime, int duration)
        {
            TimeSpan mTime = DateTime.Now - startTime;
            int wt = 0;
            while (element == null && wt < duration)
            {
                Console.WriteLine("Try to Find " + name + " at : " + System.DateTime.Now);
                element = AUIUtilities.FindElementByName(name, root );
                mTime = DateTime.Now - startTime;
                wt = wt + 2;
                Console.WriteLine(name + " find time is (sec) :" + wt * 2);
            }

            if (element == null)
                Console.WriteLine("after " + duration + " seconds" + name + " is not found time is (sec) :" + mTime.Milliseconds);
            else
                Console.WriteLine(name + " found time is (sec) :" + mTime.Milliseconds);

        }

        static public bool IsApplicationInstalled(string ApplicationType)
        {
            bool applicationInstalled = false;

            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, EpiaUtilities.GetProgramsFeaturesScreenNaam());
            AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
            if (appElement != null)
            {   // (1) Programs and Features main window
                Console.WriteLine("Programs and Features main window opend");
                Thread.Sleep(1000);
                Console.WriteLine("Searching programs item button...");
                AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(appElement, Constants.PROGRAMS_FEATURES_FOLDER_VIEW_ID);
                if (aeGridView != null)
                    Console.WriteLine("Gridview found...");
                Thread.Sleep(1000);
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.DataItem);

                AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children, c);
                Console.WriteLine("Programs count ..." + aeProgram.Count);
                for (int i = 0; i < aeProgram.Count; i++)
                {
                    switch (ApplicationType)
                    {
                        case "Epia":
                            if (aeProgram[i].Current.Name.StartsWith("E'pia Fr"))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "EtriccCore":
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                    && aeProgram[i].Current.Name.IndexOf("Shell") < 0
                                 && aeProgram[i].Current.Name.IndexOf("Playback") < 0
                                   && aeProgram[i].Current.Name.IndexOf("Statistics") < 0
                                && aeProgram[i].Current.Name.IndexOf("HostTest") < 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "EtriccShell":
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                    && aeProgram[i].Current.Name.IndexOf("Shell") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "Ewcs":
                            if (aeProgram[i].Current.Name.StartsWith("E'wcs")
                                && aeProgram[i].Current.Name.IndexOf("TestPrograms") < 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "EwcsTestProgram":
                            if (aeProgram[i].Current.Name.StartsWith("E'wcs")
                                && aeProgram[i].Current.Name.IndexOf("TestPrograms") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                        case "EpiaResourceFileEditor":
                            if (aeProgram[i].Current.Name.StartsWith("E'pia Resource"))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Thread.Sleep(5000);
                                break;
                            }
                            break;
                    }
                }
            }



            return applicationInstalled;
        }

        static public bool IsApplicationInstalled(string ApplicationType, string uninstallWindowName)
        {
            bool applicationInstalled = false;

            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName);
            AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
            if (appElement != null)
            {   // (1) Programs and Features main window
                Console.WriteLine("Programs and Features main window opend");
                Wait(1);
                Console.WriteLine("Searching programs item button...");
                AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(appElement, Constants.PROGRAMS_FEATURES_FOLDER_VIEW_ID);
                if (aeGridView != null)
                    Console.WriteLine("Gridview found...");
                Wait(1);
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.DataItem);

                AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children, c);
                Console.WriteLine("Programs count ..." + aeProgram.Count);
                for (int i = 0; i < aeProgram.Count; i++)
                {
                    switch (ApplicationType)
                    {
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
                        case "EpiaResourceFileEditor":
                            if (aeProgram[i].Current.Name.StartsWith("E'pia Resource File Editor"))
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
    
        static public AutomationElement WalkerTreeViewNextChildNede(AutomationElement aeCurrentNode, string nextChildName,ref string errorMsg)
        {
            AutomationElement aeNextChildNode = null;
            Console.WriteLine("\n=== Find " + nextChildName + " node ===");
            Thread.Sleep(2000);
            System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
            aeNextChildNode = TestTools.AUICommon.WalkEnabledElements(aeCurrentNode, treeNode, nextChildName);
            if (aeNextChildNode == null)
            {
                errorMsg = "\n=== " + nextChildName + " node NOT Exist ===";
            }
            else
            {
                Console.WriteLine("\n=== " + aeNextChildNode + " node Exist ===");
                try
                {
                    ScrollItemPattern sip = (ScrollItemPattern)aeNextChildNode.GetCurrentPattern(ScrollItemPattern.Pattern);
                    sip.ScrollIntoView();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("aeFrogramFilesNode is visible  no scroll needed: " + aeNextChildNode.Current.Name);
                }

                try
                {
                    ExpandCollapsePattern ep = (ExpandCollapsePattern)aeNextChildNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                    ep.Expand();
                    Thread.Sleep(2000);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ae" + nextChildName + " Node can not expaned: " + aeNextChildNode.Current.Name);
                }
            }
            return aeNextChildNode;
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
                        int k = 0;
                        while (aeWindow.Current.IsEnabled)
                        {
                            double x = aeTitleBar.Current.BoundingRectangle.Left + 100;
                            double y = (aeTitleBar.Current.BoundingRectangle.Top + aeTitleBar.Current.BoundingRectangle.Bottom) / 2;
                            System.Windows.Point myPlacePoint = new System.Windows.Point(x, y);
                            Input.MoveTo(myPlacePoint);
                            Thread.Sleep(2000);
                            Input.MoveToAndClick(myPlacePoint);
                            Thread.Sleep(2000);
                            //click my settings
                            Input.MoveTo(new System.Windows.Point(x, y + 25));
                            Thread.Sleep(2000);
                            Console.WriteLine("re click mysetting point : "+k++);
                            Input.MoveToAndClick(new System.Windows.Point(x, y + 25));
                            Thread.Sleep(5000);
                        }
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
                    if (aeWindow.Current.IsEnabled)
                        errorMSG = fileName + " aeMySettingsWindow not found";
                    else
                    {
                        AutomationElement aeError = AUIUtilities.FindElementByID("ErrorScreen", AutomationElement.RootElement);
                        if (aeError != null)
                            ErrorWindowHandling(aeError, ref errorMSG);
                    }
                }
                else
                {
                    Console.WriteLine("aeMySettingsWindow found");
                    AutomationElement aeDropDownBtn = AUIUtilities.FindElementByID("DropDown", aeMySettingsWindow);
                    if (aeDropDownBtn != null)
                        Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeDropDownBtn));
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
                    Thread.Sleep(1000);
                    ExpandCollapsePattern cP = aeCombo.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
                    cP.Expand();
                    Thread.Sleep(1000);
                    //SelectionPattern selectPattern =
                    //    aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                    // find item
                    AutomationElement item = AUIUtilities.FindElementByName(language, aeCombo); 
                    if (item != null)
                    {
                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));
                        Console.WriteLine("LanguageSettings item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(1000);
                        // be sure select again
                        //SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        //itemPattern.Select();
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        result = false;
                        errorMSG = fileName +  " Finding Language in combo failed: " + System.DateTime.Now.ToString("HH:mm:ss");
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
                    errorMSG = fileName +  " FindElementAndClick failed:" + "m_btnSave";
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
                    errorMSG = fileName+  "SwitchLanguageAndFindText:: Main window not found "; ;
                }
                else
                {
                    if (fileName.IndexOf("_cn") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("我的位置", aeWindow);   //我的位置
                    else if (fileName.IndexOf("_de") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Meine Einstellungen", aeWindow);// Meine Einstellungen  
                    else if (fileName.IndexOf("_el") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Η τοποθεσία μου", aeWindow);      // Η τοποθεσία μου             
                    else if (fileName.IndexOf("_en") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("My Place", aeWindow);        // My Place     
                    else if (fileName.IndexOf("_es") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("My Place", aeWindow);       // TEMP    My Place             
                    else if (fileName.IndexOf("_fr") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Ma place", aeWindow);      // Ma place               
                    else if (fileName.IndexOf("_nl") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Mijn plek", aeWindow);           // Mijn plek       
                    else if (fileName.IndexOf("_pl") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Moje miejsce", aeWindow);  // Moje miejsce


                    if (aeValidationText == null)
                    {
                        result = false;
                        errorMSG = "SwitchLanguageAndFindText:: validation text not found: "+fileName;
                    }
                }
               
            }
          

            return result;
        }


        static public string getReleaseFromHotfixVersion(string buildnr)
        {
            int ind = buildnr.IndexOf("of");
            string version = buildnr.Substring(7, ind - 8);

            int indLastComma = version.LastIndexOf('.');
            string releaseVersion = version.Substring(0, indLastComma);
            return releaseVersion;
        }

        static public DateTime getReleaseVersionDate(string version)
        {
            DateTime datetime = DateTime.Today;
            TfsTeamProjectCollection tfsProjectCollection;
            string selectedProject = "Epia 4";

            string sTFSServerUrl ="http://team2010App.teamSystems.egemin.be:8080/tfs/Development";
            string sTFSUsername = "TfsBuild";
            string sTFSPassword = "Egemin01";
            string sTFSDomain = "TeamSystems.Egemin.Be";
            Uri serverUri = new Uri(sTFSServerUrl);
            System.Net.ICredentials tfsCredentials
               = new System.Net.NetworkCredential(sTFSUsername, sTFSPassword, sTFSDomain);

            tfsProjectCollection
                = new TfsTeamProjectCollection(serverUri, tfsCredentials);

            IBuildServer buildServer;
            tfsProjectCollection.EnsureAuthenticated();
            buildServer = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));

            IBuildDetailSpec buildDetailSpec = buildServer.CreateBuildDetailSpec(selectedProject, "Epia.Production.Release");
            //buildDetailSpec.MaxBuildsPerDefinition = 1; 
            buildDetailSpec.QueryOrder = BuildQueryOrder.FinishTimeDescending;
            buildDetailSpec.Status = BuildStatus.Succeeded; //Only get succeeded builds  
            //buildDetailSpec.MinFinishTime = timeFrom;
            IBuildQueryResult results = buildServer.QueryBuilds(buildDetailSpec);
            //if (results.Failures.Length == 0 ) 
            //{ 
            //IBuildDetail buildDetail = results.Builds[0]; 
            //Console.WriteLine("Build: " + buildDetail.BuildNumber); 
            //Console.WriteLine("Account requesting build “ + 
            //“(build service user for triggered builds): " + buildDetail.RequestedBy); 
            //   Console.WriteLine("Build triggered by: " + buildDetail.RequestedFor); 
            //}

            IBuildDetail[] buildnrs = results.Builds;
            //IBuildDetail[] buildnrs = buildServer.QueryBuilds(selectedProject, s);
            string bnrs = string.Empty;
            string quality = string.Empty;
            //BuildObject thisBuild = new BuildObject();

            for (int i = 0; i < buildnrs.Length; i++)
            {
                if (buildnrs[i].BuildNumber.IndexOf(version) > 0)
                {
                    datetime = buildnrs[i].FinishTime;
                    //System.Windows.Forms.MessageBox.Show("release version: " + buildnrs[i].BuildNumber, "count:" + buildnrs.Length);
                    //System.Windows.Forms.MessageBox.Show("date: " + buildnrs[i].FinishTime, "count:" + buildnrs.Length);
                    break;
                }
            }


            return datetime;
        }

        public static void ErrorWindowHandling(AutomationElement element, ref string ErrorMSG)
        {
            string close = "Close";
            string error = string.Empty;
            AutomationElement aeError = AUIUtilities.FindElementByType(ControlType.Text, element);
            if (aeError == null)
            {
                error = "Error Message Element not Fund";
                Console.WriteLine(error);
                return;
            }
            else
            {
                Console.WriteLine("aeError is found ------------:");
            }
            ErrorMSG = aeError.Current.Name;
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

        static public bool ProcessSecurityForm(AutomationElement aeSecurityForm, string tester, string UserPassword, ref string ErrorMSG)
        {
            bool status = true;
                #region
                //Console.WriteLine("Application aeSecurityForm name : " + aeSecurityForm.Current.Name);
                string UserNameID = "m_TextBoxUsername"; //"ControlType.Edit" Name : "with Windows credentials
                string PasswordID = "m_TextBoxPassword";
                string BtnOKID = "m_BtnOK";
                string origUser = string.Empty;
                if (AUIUtilities.FindTextBoxAndChangeValue(PasswordID, aeSecurityForm, out origUser, UserPassword, ref ErrorMSG))
                    Thread.Sleep(3000);
                else
                {
                    ErrorMSG = "FindTextBoxAndChangeValue failed:" + PasswordID;
                    Console.WriteLine(ErrorMSG);
                    status = false;
                }

                if (AUIUtilities.FindTextBoxAndChangeValue(UserNameID, aeSecurityForm, out origUser, tester, ref ErrorMSG))
                    Thread.Sleep(3000);
                else
                {
                    ErrorMSG = "FindTextBoxAndChangeValue failed:" + UserNameID;
                    Console.WriteLine(ErrorMSG);
                    status = false;
                }

                 // Logon into Application
                Thread.Sleep(3000);

                // Find Logon OK Button and click 
                if (AUIUtilities.FindElementAndClick(BtnOKID, aeSecurityForm))
                    Thread.Sleep(3000);
                else
                {
                    ErrorMSG = "FindElementAndClick failed:" + BtnOKID;
                    Console.WriteLine(ErrorMSG);
                    status = false;
                }
                #endregion
        
            return status;
        }

    }
}
