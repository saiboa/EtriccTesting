﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using QATestEpiaNet45UI;
using TestTools;

namespace Egemin.Epia.Testing.QATestEpiaNet45UI
{
    internal class EpiaUtilities
    {
        /// <summary>
        /// Retrieves the top-level window that contains the specified UI Automation element.
        /// </summary>
        /// <param name="element">The contained element.</param>
        /// <returns>The containing top-level window element.</returns>
        public static AutomationElement GetTopLevelWindow(AutomationElement element)
        {
            TreeWalker walker = TreeWalker.ControlViewWalker;
            AutomationElement elementParent;
            AutomationElement node = element;
            //if (node == elementRoot) return node;
            do
            {
                elementParent = walker.GetParent(node);
                if (elementParent == AutomationElement.RootElement) break;
                node = elementParent;
            } while (true);
            return node;
        }

        public static void ClearDisplayedScreens(AutomationElement root)
        {
            AutomationElementCollection aeAllTabs = root.FindAll(TreeScope.Descendants, new PropertyCondition(
                                                                                            AutomationElement.
                                                                                                ControlTypeProperty,
                                                                                            ControlType.Tab));

            for (int k = 0; k < aeAllTabs.Count; k++)
            {
                if (aeAllTabs[k] != null)
                {
                    double right = aeAllTabs[k].Current.BoundingRectangle.Right;
                    double bottom = aeAllTabs[k].Current.BoundingRectangle.Bottom;
                    double top = aeAllTabs[k].Current.BoundingRectangle.Top;

                    double x = right - 5;
                    double y = (top + bottom)/2;
                    var p = new Point(x, y);

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
        public static string[] GetFiles(string path, string searchPattern)
        {
            string[] fileTypes = searchPattern.Split(';');
            var strFiles = new List<string>();
            foreach (string filter in fileTypes)
                strFiles.AddRange(Directory.GetFiles(path, filter));

            string[] absolutefileNames = strFiles.ToArray();

            var fileNames = new string[absolutefileNames.Length];
            for (int i = 0; i < absolutefileNames.Length; i++)
            {
                var fileInfo = new FileInfo(absolutefileNames[i]);
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
                var b = (string) EmpEnumerator.Current;
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
            string[] names1 = File.ReadAllLines(@"../../../names1.txt");
            string[] names2 = File.ReadAllLines(@"../../../names2.txt");

            // Create the query. Note that method syntax must be used here.
            IEnumerable<string> differenceQuery =
                names1.Except(names2);

            IEnumerator EmpEnumerator = differenceQuery.GetEnumerator(); //Getting the Enumerator
            EmpEnumerator.Reset(); //Position at the Beginning
            while (EmpEnumerator.MoveNext()) //Till not finished do print
            {
                var b = (string) EmpEnumerator.Current;
                Console.WriteLine("The following lines are in names1.txt but not names2.txt  --- " + b);
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
                catch (Exception ex)
                {
                    aeWindow = null;
                    Console.WriteLine(ex.Message + " ---- " + ex.StackTrace);
                    Thread.Sleep(5000);
                }
                mTime = DateTime.Now - mStartTime;
            }

            return aeWindow;
        }

        public static AutomationElement GetMainWindow(string mainFormId, int waitSec)
        {
            AutomationElement aeWindow = null;
            AutomationElementCollection aeAllWindows = null;
            // find main window
            Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            int k = 0;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while (aeWindow == null && mTime.TotalSeconds <= waitSec)
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
                catch (Exception ex)
                {
                    aeWindow = null;
                    Console.WriteLine(ex.Message + " ---- " + ex.StackTrace);
                    Thread.Sleep(5000);
                }
                mTime = DateTime.Now - mStartTime;
            }

            return aeWindow;
        }

        public static AutomationElement GetSelectedOverviewWindow(string reportWindowName, ref string errorMsg)
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

            if (result)
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

        public static AutomationElement GetCellElementFromOverviewWindow(AutomationElement aeOverview, string colName,
                                                                         int row, ref string errorMsg)
        {
            AutomationElement aeCell = null;

            #region // Find GridView

            AutomationElement aeGrid = AUIUtilities.FindElementById("m_GridData", aeOverview);
            if (aeGrid == null)
            {
                errorMsg = aeOverview.Current.Name + " DataGridView not found";
            }
            else
            {
                Console.WriteLine("DataGridView found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                Thread.Sleep(3000);
                // Construct the Grid Cell Element Name
                string cellname = colName + " Row " + row;
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

        public static bool WindowMenuAction(AutomationElement aeWindow, string colHeaderName, int row,
                                            string menuItemName, ref string errorMsg)
        {
            bool result = true;

            #region // Find GridView

            AutomationElement aeGrid = AUIUtilities.FindElementById("m_GridData", aeWindow);
            if (aeGrid == null)
            {
                errorMsg = aeWindow.Current.Name + " GridData not found";
                Console.WriteLine(errorMsg);
                result = false;
            }
            else
            {
                Console.WriteLine(aeWindow.Current.Name + " GridData found at time: " +
                                  DateTime.Now.ToString("HH:mm:ss"));
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
                    Console.WriteLine(aeCell.Current.Name + "    -----  cell found at time: " +
                                      DateTime.Now.ToString("HH:mm:ss"));
                    Thread.Sleep(3000);
                    Point cellPoint = AUIUtilities.GetElementCenterPoint(aeCell);
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
                        Point menuItemPoint = AUIUtilities.GetElementCenterPoint(aeMenuItem);
                        Input.MoveToAndClick(menuItemPoint);
                        result = true;
                    }
                }
            }

            #endregion

            return result;
        }


        public static AutomationElement GetMenuItemFromElement(AutomationElement element, string menuItemId, int seconds,
                                                               ref string errorMsg)
        {
            AutomationElement aeMenuItem = null;
            AutomationElementCollection aeAllMenuItems = null;
            Condition cMenuItems = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem);

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

        public static AutomationElement GetCategoryWindow(string windowName, ref string errorMsg)
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
            var ivp = (InvokePattern) aeClose.GetCurrentPattern(InvokePattern.Pattern);
            ivp.Invoke();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="choice">Shell or Server or Both</param>
        public static bool InstallEpia(string choice, ref string errorMsg)
        {
            AutomationElement rootElement = AutomationElement.RootElement;
            AutomationElement btnNext = null;

            #region Install Epia

            Console.WriteLine("Searching for main installer window");
            Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
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
                    MessageBox.Show("After one minute no Installer Window Form found");
                    return false;
                }
            }

            if (appElement != null)
            {
                // (1) Welcom Main window
                Console.WriteLine("Welcom Main window opend ");
                Console.WriteLine("Searching next button...");
                btnNext = AUIUtilities.GetElementByNameProperty(appElement, "Next >");
                if (btnNext != null)
                {
                    // (2) Components
                    AUIUtilities.ClickElement(btnNext);
                    appElement = rootElement.FindFirst(TreeScope.Children, condition);
                    Console.WriteLine("Componts window opend");
                    Console.WriteLine("Searching checkbox...");
                    AutomationElement epiaServerCheckbox = AUIUtilities.GetElementByNameProperty(appElement,
                                                                                                 "E'pia Server");
                    if (epiaServerCheckbox != null &&
                        (choice.ToLower().StartsWith("server") || choice.ToLower().StartsWith("both")))
                    {
                        AUIUtilities.ClickElement(epiaServerCheckbox);
                    }

                    AutomationElement epiaShellCheckbox = AUIUtilities.GetElementByNameProperty(appElement,
                                                                                                "E'pia Shell");
                    if (epiaShellCheckbox != null &&
                        (choice.ToLower().StartsWith("shell") || choice.ToLower().StartsWith("both")))
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
                        Condition c2 = new AndCondition(new PropertyCondition(AutomationElement.NameProperty, "Close"),
                                                        new PropertyCondition(
                                                            AutomationElement.IsContentElementProperty, true));

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

        private static void Wait(int seconds)
        {
            Thread.Sleep(seconds*1000);
        }

        public static string GetProgramsFeaturesScreenNaam()
        {
            string screenName = "Programs and Features";
            string MachineName = Environment.MachineName;
            if (MachineName.ToUpper().StartsWith("ETRICCSTATAUTO") ||
                MachineName.ToUpper().StartsWith("ETRICCAUTOTEST1"))
                screenName = "Control Panel\\Programs\\Programs and Features";

            return screenName;
        }

        public static void StartProgramsAndFeaturesExecution()
        {
            var Proc = new Process();
            Proc.StartInfo.FileName = @"C:\Windows\System32\appwiz.cpl";
            Proc.StartInfo.CreateNoWindow = false;
            Proc.Start();
        }

        public static void StartEpiaResourceFileEditorExecution()
        {
            var Proc = new Process();
            Proc.StartInfo.FileName =
                Path.Combine(OSVersionInfoClass.ProgramFilesx86() + "\\Egemin\\Epia Resource File Editor",
                             "Egemin.Epia.Foundation.Globalization.ResourceFileEditor.exe");
            Proc.StartInfo.CreateNoWindow = false;
            Proc.Start();
        }

        public static void WaitUntilElementByNameFound(AutomationElement root, ref AutomationElement element,
                                                       string name,
                                                       DateTime startTime, int duration)
        {
            TimeSpan mTime = DateTime.Now - startTime;
            int wt = 0;
            while (element == null && wt < duration)
            {
                Console.WriteLine("Try to Find " + name + " at : " + DateTime.Now);
                element = AUIUtilities.FindElementByName(name, root);
                mTime = DateTime.Now - startTime;
                wt = wt + 2;
                Console.WriteLine(name + " find time is (sec) :" + wt*2);
            }

            if (element == null)
                Console.WriteLine("after " + duration + " seconds" + name + " is not found time is (sec) :" +
                                  mTime.Milliseconds);
            else
                Console.WriteLine(name + " found time is (sec) :" + mTime.Milliseconds);
        }

        public static bool IsApplicationInstalled(string ApplicationType)
        {
            bool applicationInstalled = false;

            Condition condition = new PropertyCondition(AutomationElement.NameProperty, GetProgramsFeaturesScreenNaam());
            AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
            if (appElement != null)
            {
                // (1) Programs and Features main window
                Console.WriteLine("Programs and Features main window opend");
                Thread.Sleep(1000);
                Console.WriteLine("Searching programs item button...");
                AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(appElement,
                                                                                     Constants.
                                                                                         PROGRAMS_FEATURES_FOLDER_VIEW_ID);
                if (aeGridView != null)
                    Console.WriteLine("Gridview found...");
                Thread.Sleep(1000);
                // Set a property condition that will be used to find the control.
                Condition c = new PropertyCondition(
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

        public static bool IsApplicationInstalled(string ApplicationType, string uninstallWindowName)
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
                AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(appElement,
                                                                                     Constants.
                                                                                         PROGRAMS_FEATURES_FOLDER_VIEW_ID);
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

        public static AutomationElement WalkerTreeViewNextChildNede(AutomationElement aeCurrentNode,
                                                                    string nextChildName, ref string errorMsg)
        {
            AutomationElement aeNextChildNode = null;
            Console.WriteLine("\n=== Find " + nextChildName + " node ===");
            Thread.Sleep(2000);
            var treeNode = new TreeNode();
            aeNextChildNode = AUICommon.WalkEnabledElements(aeCurrentNode, treeNode, nextChildName);
            if (aeNextChildNode == null)
            {
                errorMsg = "\n=== " + nextChildName + " node NOT Exist ===";
            }
            else
            {
                Console.WriteLine("\n=== " + aeNextChildNode + " node Exist ===");
                try
                {
                    var sip = (ScrollItemPattern) aeNextChildNode.GetCurrentPattern(ScrollItemPattern.Pattern);
                    sip.ScrollIntoView();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("aeFrogramFilesNode is visible  no scroll needed: " + aeNextChildNode.Current.Name);
                }

                try
                {
                    var ep = (ExpandCollapsePattern) aeNextChildNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
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

        public static bool SwitchLanguageAndFindText(string resourcesFolder, string fileName, Point myPlacePt,
                                                     ref string errorMSG)
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
            var DirInfo = new DirectoryInfo(resourcesFolder);
            FileInfo[] serverFolderFiles = DirInfo.GetFiles(fileName);
            if (serverFolderFiles.Length == 0)
            {
                result = false;
                errorMSG = resourcesFolder + " has no resource file:" + fileName;
            }
            else // switch to this language
            {
                //aeWindow = GetMainWindow("MainForm");
                //if (aeWindow == null)
                //{
                //    result = false;
                //    errorMSG = fileName + "SwitchLanguageAndFindText:: Min window noty found "; ;
                //}
                //else
                //{   // open my setting window
                //aeWindow.SetFocus();
                //int k = 0;
                //while (aeWindow.Current.IsEnabled)
                //{
                //Input.MoveTo(myPlacePt);
                //Thread.Sleep(500);
                Input.MoveToAndClick(myPlacePt);
                Thread.Sleep(1500);
                //click my settings
                //Input.MoveTo(new System.Windows.Point(myPlacePt.X, myPlacePt.Y + 25));
                //Thread.Sleep(2000);
                Console.WriteLine("re click mysetting point : ");
                Input.MoveToAndClick(new Point(myPlacePt.X, myPlacePt.Y + 25));
                Thread.Sleep(2000);
                //}
                //}
            }

            AutomationElement aeMySettingsWindow = null;
            if (result)
            {
                string settingsWindowId = "Dialog - Egemin.Epia.Modules.RnD.Screens.UserSettingsDetailsScreen";
                aeWindow = GetMainWindow("MainForm");
                aeMySettingsWindow = AUIUtilities.FindElementById(settingsWindowId, aeWindow);
                if (aeMySettingsWindow == null)
                {
                    result = false;
                    if (aeWindow.Current.IsEnabled)
                        errorMSG = fileName + " aeMySettingsWindow not found";
                    else
                    {
                        AutomationElement aeError = AUIUtilities.FindElementById("ErrorScreen",
                                                                                 AutomationElement.RootElement);
                        if (aeError != null)
                            ErrorWindowHandling(aeError, ref errorMSG);
                    }
                }
                else
                {
                    Console.WriteLine("aeMySettingsWindow found");
                    AutomationElement aeDropDownBtn = AUIUtilities.FindElementById("DropDown", aeMySettingsWindow);
                    if (aeDropDownBtn != null)
                        Input.MoveTo(AUIUtilities.GetElementCenterPoint(aeDropDownBtn));
                }
            }

            // change language
            // "中文(简体)", "我的位置"
            if (result)
            {
                AutomationElement aeCombo = AUIUtilities.FindElementById("languageIdComboBox", aeMySettingsWindow);
                if (aeCombo == null)
                {
                    result = false;
                    errorMSG = fileName + " LanguageSettings failed to find aeCombo at time: " +
                               DateTime.Now.ToString("HH:mm:ss");
                }
                else
                {
                    Thread.Sleep(500);
                    var cP = aeCombo.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
                    cP.Expand();
                    Thread.Sleep(500);
                    //SelectionPattern selectPattern =
                    //    aeCombo.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                    // find item
                    AutomationElement item = AUIUtilities.FindElementByName(language, aeCombo);
                    if (item != null)
                    {
                        Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(item));
                        Console.WriteLine("LanguageSettings item found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                        Thread.Sleep(1000);
                        // be sure select again
                        //SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        //itemPattern.Select();
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        result = false;
                        errorMSG = fileName + " Finding Language in combo failed: " + DateTime.Now.ToString("HH:mm:ss");
                    }
                }
            }

            // save or cancel
            if (result)
            {
                if (AUIUtilities.FindElementAndClickPoint("m_btnSave", aeMySettingsWindow))
                    Thread.Sleep(500);
                else
                {
                    result = false;
                    errorMSG = fileName + " FindElementAndClick failed:" + "m_btnSave";
                }
                Thread.Sleep(1000);
            }
            else
            {
                if (AUIUtilities.FindElementAndClickPoint("m_btnCancel", aeMySettingsWindow))
                    Thread.Sleep(1000);
                else
                {
                    result = false;
                    errorMSG = fileName + " FindElementAndClick failed:" + "m_btnCancel";
                }

                Thread.Sleep(1000);
            }

            // Validation
            AutomationElement aeValidationText = null;
            if (result)
            {
                aeWindow = GetMainWindow("MainForm");
                if (aeWindow == null)
                {
                    result = false;
                    errorMSG = fileName + "SwitchLanguageAndFindText:: Main window not found ";
                    ;
                }
                else
                {
                    if (fileName.IndexOf("_cn") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("我的位置", aeWindow); //我的位置
                    else if (fileName.IndexOf("_de") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Meine Einstellungen", aeWindow);
                        // Meine Einstellungen  
                    else if (fileName.IndexOf("_el") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Η τοποθεσία μου", aeWindow);
                        // Η τοποθεσία μου             
                    else if (fileName.IndexOf("_en") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("My Place", aeWindow); // My Place     
                    else if (fileName.IndexOf("_es") > 0)
                    {
                        aeValidationText = AUIUtilities.FindElementByName("My Place", aeWindow);
                        // TEMP    My Place      
                        if (aeValidationText == null)
                        {
                            aeValidationText = AUIUtilities.FindElementByName("Mi configuración", aeWindow);
                            // Spanish    Mi configuración 
                        }
                    }
                    else if (fileName.IndexOf("_fr") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Ma place", aeWindow);
                        // Ma place               
                    else if (fileName.IndexOf("_nl") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Mijn plek", aeWindow); // Mijn plek       
                    else if (fileName.IndexOf("_pl") > 0)
                        aeValidationText = AUIUtilities.FindElementByName("Moje miejsce", aeWindow); // Moje miejsce


                    if (aeValidationText == null)
                    {
                        result = false;
                        errorMSG = "SwitchLanguageAndFindText:: validation text not found: " + fileName;
                    }
                }
            }


            return result;
        }

        public static bool AddNewRole(string logFilePath, string roleName, string roleDesc, string exitModeId,
                                      bool sOnlyUITest, ref string ErrorMSG)
        {
            bool addOK = true;
            AutomationElement aeWindow = null;
            AutomationElement aeRoleAddEditDialog = null;
            if (addOK)
            {
                string RoleAddEditDialogId = "Dialog - Egemin.Epia.Modules.RnD.Screens.RoleAddEditDialog";
                aeWindow = GetMainWindow("MainForm");
                aeRoleAddEditDialog = AUIUtilities.FindElementById(RoleAddEditDialogId, aeWindow);
                if (aeRoleAddEditDialog == null)
                {
                    ErrorMSG = "aeRoleAddEditDialog not opened :";
                    Console.WriteLine(ErrorMSG);
                    Epia3Common.WriteTestLogFail(logFilePath, ErrorMSG, sOnlyUITest);
                    addOK = false;
                }
                else
                {
                    //ControlType:	"ControlType.Edit"  AutomationId:	"nameTextBox"
                    AutomationElement aeRoleNameEdit = AUIUtilities.FindElementById("nameTextBox", aeRoleAddEditDialog);
                    if (aeRoleNameEdit == null)
                    {
                        ErrorMSG = "FindTextBoxAndChangeValue failed:" + "nameTextBox";
                        Console.WriteLine(ErrorMSG);
                        Epia3Common.WriteTestLogFail(logFilePath, "AddNewRole", sOnlyUITest);
                        addOK = false;
                    }
                    else
                        SendTextToElement(aeRoleNameEdit, roleName);

                    if (addOK)
                    {
                        // ControlType:	"ControlType.Edit",   AutomationId:	"descriptionTextBox"    , Name:	"Description:"
                        AutomationElement aeRoleDescEdit = AUIUtilities.FindElementById("descriptionTextBox",
                                                                                        aeRoleAddEditDialog);
                        if (aeRoleDescEdit == null)
                        {
                            ErrorMSG = "FindTextBoxAndChangeValue failed:" + "descriptionTextBox";
                            Console.WriteLine(ErrorMSG);
                            Epia3Common.WriteTestLogFail(logFilePath, "AddNewRole", sOnlyUITest);
                            addOK = false;
                        }
                        else
                            SendTextToElement(aeRoleDescEdit, roleDesc);
                    }
                }
            }

            AutomationElement aePnlExitConfigMain = null;
            if (addOK)
            {
                aePnlExitConfigMain = AUIUtilities.FindElementById("m_PnlExitConfigMain", aeRoleAddEditDialog);
                if (aePnlExitConfigMain == null)
                {
                    ErrorMSG = "aePnlExitConfigMain not found :";
                    Console.WriteLine(ErrorMSG);
                    Epia3Common.WriteTestLogFail(logFilePath, ErrorMSG, sOnlyUITest);
                    addOK = false;
                }
            }

            if (addOK)
            {
                string inActivitySettingId = "exitConfigurationDefinedCheckbox";
                bool check = AUIUtilities.FindElementAndToggle(inActivitySettingId, aePnlExitConfigMain, ToggleState.On);
                if (check)
                {
                    Thread.Sleep(3000);
                    // find logout radio button
                    AutomationElement aeLogoutRadio = AUIUtilities.FindElementById(exitModeId, aePnlExitConfigMain);
                    if (aeLogoutRadio == null)
                    {
                        ErrorMSG = "aeLogoutRadio not found";
                        Console.WriteLine(ErrorMSG);
                        Epia3Common.WriteTestLogFail(logFilePath, "AddNewRole", sOnlyUITest);
                        addOK = false;
                    }
                    else
                    {
                        var itemRadioPattern =
                            aeLogoutRadio.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        itemRadioPattern.Select();
                        Thread.Sleep(3000);
                    }
                }
                else
                {
                    Console.WriteLine("FindElementAndToggle failed:" + inActivitySettingId);
                    Epia3Common.WriteTestLogFail(logFilePath, "FindElementAndToggle failed:" + inActivitySettingId,
                                                 sOnlyUITest);
                    addOK = false;
                }
            }

            string origValue = string.Empty;
            if (addOK)
            {
                // ControlType:	"ControlType.Edit"
                //AutomationId:	"descriptionTextBox"
                //LocalizedControlType:	"edit"
                //Name:	"Text area"
                AutomationElement aeTextAreaEdit = AUIUtilities.FindElementByType(ControlType.Edit, aePnlExitConfigMain);
                if (aeTextAreaEdit == null)
                {
                    ErrorMSG = "FindTextBoxAndChangeValue failed:" + "Text area";
                    Console.WriteLine(ErrorMSG);
                    Epia3Common.WriteTestLogFail(logFilePath, "AddNewRole", sOnlyUITest);
                    addOK = false;
                }
                else
                {
                    Console.WriteLine("aeCell name is :" + aeTextAreaEdit.Current.Name);
                    Point pt = AUIUtilities.GetElementCenterPoint(aeTextAreaEdit);
                    Thread.Sleep(2000);
                    Input.MoveTo(pt);
                    Thread.Sleep(2000);
                    Input.ClickAtPoint(pt);
                    Thread.Sleep(2000);
                    Input.ClickAtPoint(pt);
                    Thread.Sleep(2000);
                    Input.ClickAtPoint(pt);
                    SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                    Thread.Sleep(1000);
                    SendKeys.SendWait("1");
                    Thread.Sleep(1000);
                }
            }

            if (addOK)
            {
                string BtnSaveId = "m_btnSave";
                AutomationElement aeSave = AUIUtilities.FindElementById(BtnSaveId, aeRoleAddEditDialog);
                if (aeSave == null)
                {
                    ErrorMSG = "failed to find aeSave of aeRoleAddEditDialog";
                    Console.WriteLine("AddNewRole" + " failed to find aeSave at time: " +
                                      DateTime.Now.ToString("HH:mm:ss"));
                    addOK = false;
                }
                else
                {
                    Input.MoveTo(aeSave);
                    Console.WriteLine("AddNewRole" + " aeSave found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                    var ipc =
                        aeSave.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    ipc.Invoke();
                }
                Thread.Sleep(5000);
            }
            else
            {
                string BtnnCancelId = "m_btnCancel";
                AutomationElement aeCancel = AUIUtilities.FindElementById(BtnnCancelId, aeRoleAddEditDialog);
                if (aeCancel == null)
                {
                    ErrorMSG = "failed to find aeCancel of aeRoleAddEditDialog";
                    Console.WriteLine("AddNewRole" + " failed to find aeCancel at time: " +
                                      DateTime.Now.ToString("HH:mm:ss"));
                    addOK = false;
                }
                else
                {
                    Console.WriteLine("AddNewRole" + " aeCancel found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                    var ipc =
                        aeCancel.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    ipc.Invoke();
                }
                Thread.Sleep(5000);
            }

            return addOK;
        }

        public static void SendTextToElement(AutomationElement aeEditElement, string thisText)
        {
            Console.WriteLine("aeCell name is :" + aeEditElement.Current.Name);
            Point pt = AUIUtilities.GetElementCenterPoint(aeEditElement);
            Input.MoveTo(pt);
            Thread.Sleep(500);
            Input.ClickAtPoint(pt);
            Thread.Sleep(1000);
            SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
            Thread.Sleep(1000);
            SendKeys.SendWait(thisText);
            Thread.Sleep(500);
        }

        public static bool AddNewAccount(string logFilePath, string accountName, string UserPassword, string roleDesc,
                                         bool inactivityConfig, string exitModeId, int inactiveMin, string roleName,
                                         bool sOnlyUITest, ref string ErrorMSG)
        {
            bool addOK = true;
            AutomationElement aeWindow = null;
            AutomationElement aeAccountAddEditDialog = null;
            if (addOK)
            {
                string AccountAddEditDialogId = "Dialog - Egemin.Epia.Modules.RnD.Screens.EpiaAccountAddEdit";
                aeWindow = GetMainWindow("MainForm");
                aeAccountAddEditDialog = AUIUtilities.FindElementById(AccountAddEditDialogId, aeWindow);
                if (aeAccountAddEditDialog == null)
                {
                    ErrorMSG = "aeAccountAddEditDialog not opened :";
                    Console.WriteLine(ErrorMSG);
                    Epia3Common.WriteTestLogFail(logFilePath, ErrorMSG, sOnlyUITest);
                    addOK = false;
                }
                else
                {
                    AutomationElement aeAccountNameEdit = AUIUtilities.FindElementById("accountNameTextBox",
                                                                                       aeAccountAddEditDialog);
                    if (aeAccountNameEdit == null)
                    {
                        ErrorMSG = "FindTextBoxAndChangeValue failed:" + "m_Password";
                        Console.WriteLine(ErrorMSG);
                        Epia3Common.WriteTestLogFail(logFilePath, "AddNewAccount", sOnlyUITest);
                        addOK = false;
                    }
                    else
                        SendTextToElement(aeAccountNameEdit, accountName);

                    if (addOK) // Enter password : ControlType:	"ControlType.Edit"  //AutomationId:	"m_Password"
                    {
                        AutomationElement aePasswordEdit = AUIUtilities.FindElementById("m_Password",
                                                                                        aeAccountAddEditDialog);
                        if (aePasswordEdit == null)
                        {
                            ErrorMSG = "FindTextBoxAndChangeValue failed:" + "m_Password";
                            Console.WriteLine(ErrorMSG);
                            Epia3Common.WriteTestLogFail(logFilePath, "AddNewAccount", sOnlyUITest);
                            addOK = false;
                        }
                        else
                            SendTextToElement(aePasswordEdit, UserPassword);
                    }

                    Thread.Sleep(3000);
                    if (addOK) // Reenter password : ControlType:	"ControlType.Edit"  //AutomationId:	"m_Password2"
                    {
                        AutomationElement aePasswordEdit = AUIUtilities.FindElementById("m_Password2",
                                                                                        aeAccountAddEditDialog);
                        if (aePasswordEdit == null)
                        {
                            ErrorMSG = "FindTextBoxAndChangeValue failed:" + "m_Password2";
                            Console.WriteLine(ErrorMSG);
                            Epia3Common.WriteTestLogFail(logFilePath, "AddNewAccount", sOnlyUITest);
                            addOK = false;
                        }
                        else
                            SendTextToElement(aePasswordEdit, UserPassword);
                    }
                }
            }

            AutomationElement aePnlExitConfigMain = null;
            if (addOK && inactivityConfig)
            {
                Console.WriteLine("addOK:" + addOK);
                Console.WriteLine("inactivityConfig:" + inactivityConfig);
                aePnlExitConfigMain = AUIUtilities.FindElementById("m_PnlExitConfigMain", aeAccountAddEditDialog);
                if (aePnlExitConfigMain != null)
                {
                    string inActivitySettingId = "exitConfigurationDefinedCheckbox";
                    bool check = AUIUtilities.FindElementAndToggle(inActivitySettingId, aePnlExitConfigMain,
                                                                   ToggleState.On);
                    if (check)
                    {
                        Thread.Sleep(2000); // find logout radio button
                        AutomationElement aeLogoutRadio = AUIUtilities.FindElementById(exitModeId, aePnlExitConfigMain);
                        if (aeLogoutRadio == null)
                        {
                            ErrorMSG = "aeLogoutRadio not found";
                            Console.WriteLine(ErrorMSG);
                            Epia3Common.WriteTestLogFail(logFilePath, "AddNewAccount", sOnlyUITest);
                            addOK = false;
                        }
                        else
                        {
                            var itemRadioPattern =
                                aeLogoutRadio.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            itemRadioPattern.Select();
                            Thread.Sleep(2000);
                        }

                        if (addOK) // Enter Inactivity timeout: ControlType:	"ControlType.Edit", Name:	"Text area"
                        {
                            AutomationElement aeTextAreaEdit = AUIUtilities.FindElementByType(ControlType.Edit,
                                                                                              aePnlExitConfigMain);
                            if (aeTextAreaEdit == null)
                            {
                                ErrorMSG = "FindTextBoxAndChangeValue failed:" + "Text area";
                                Console.WriteLine(ErrorMSG);
                                Epia3Common.WriteTestLogFail(logFilePath, "AddNewAccount", sOnlyUITest);
                                addOK = false;
                            }
                            else
                                SendTextToElement(aeTextAreaEdit, "" + inactiveMin);
                        }
                    }
                    else
                    {
                        Console.WriteLine("FindElementAndToggle failed:" + inActivitySettingId);
                        Epia3Common.WriteTestLogFail(logFilePath, "FindElementAndToggle failed:" + inActivitySettingId,
                                                     sOnlyUITest);
                        addOK = false;
                    }
                }
                else
                {
                    ErrorMSG = "aePnlExitConfigMain not found :";
                    Console.WriteLine(ErrorMSG);
                    Epia3Common.WriteTestLogFail(logFilePath, ErrorMSG, sOnlyUITest);
                    addOK = false;
                }
            }

            if (addOK && roleName.Length > 1) // select role
            {
                // ControlType:	"ControlType.Tree", AutomationId:	"m_TreeRoles"   , Name:	"Roles:"
                AutomationElement aeTreeRoleArea = AUIUtilities.FindElementById("m_TreeRoles", aeAccountAddEditDialog);
                if (aeTreeRoleArea == null)
                {
                    ErrorMSG = "FindTextBoxAndChangeValue failed:" + "aeTreeRoleArea";
                    Console.WriteLine(ErrorMSG);
                    Epia3Common.WriteTestLogFail(logFilePath, "AddNewAccount", sOnlyUITest);
                    addOK = false;
                }
                else
                {
                    aeTreeRoleArea.SetFocus();
                    AutomationElement aeRoleName = AUIUtilities.FindElementByName(roleName, aeTreeRoleArea);
                    if (aeRoleName == null)
                    {
                        ErrorMSG = aeRoleName + " checkbox not found:";
                        Console.WriteLine(ErrorMSG);
                        Epia3Common.WriteTestLogFail(logFilePath, "AddNewAccount", sOnlyUITest);
                        addOK = false;
                    }
                    else
                    {
                        Console.WriteLine(roleName + " found: ");
                        var tg = aeRoleName.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                        Thread.Sleep(500);
                        ToggleState tgTState = tg.Current.ToggleState;
                        Console.WriteLine("FindElementAndToggle to: " + tgTState.ToString());
                        if (tgTState == ToggleState.Off)
                        {
                            Console.WriteLine("FindElementAndToggle to: " + tgTState.ToString());
                            double x = aeRoleName.Current.BoundingRectangle.Left + 5.0;
                            double y = (aeRoleName.Current.BoundingRectangle.Bottom +
                                        aeRoleName.Current.BoundingRectangle.Top)/2.0;
                            var pt = new Point(x, y);
                            Input.MoveToAndClick(pt);
                            Thread.Sleep(5000);
                            Console.WriteLine("FindElementAndToggle to: " + tgTState.ToString());
                        }
                    }
                }
            }

            if (addOK)
            {
                string BtnSaveId = "m_btnSave";
                AutomationElement aeSave = AUIUtilities.FindElementById(BtnSaveId, aeAccountAddEditDialog);
                if (aeSave == null)
                {
                    ErrorMSG = "failed to find aeSave of aeRoleAddEditDialog";
                    Console.WriteLine("AddNewAccount" + " failed to find aeSave at time: " +
                                      DateTime.Now.ToString("HH:mm:ss"));
                    addOK = false;
                }
                else
                {
                    Input.MoveTo(aeSave);
                    Console.WriteLine("AddNewAccount" + " aeSave found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                    var ipc =
                        aeSave.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    ipc.Invoke();
                }
                Thread.Sleep(5000);
            }
            else
            {
                string BtnnCancelId = "m_btnCancel";
                AutomationElement aeCancel = AUIUtilities.FindElementById(BtnnCancelId, aeAccountAddEditDialog);
                if (aeCancel == null)
                {
                    ErrorMSG = "failed to find aeCancel of aeRoleAddEditDialog";
                    Console.WriteLine("AddNewAccount" + " failed to find aeCancel at time: " +
                                      DateTime.Now.ToString("HH:mm:ss"));
                    addOK = false;
                }
                else
                {
                    Console.WriteLine("AddNewAccount" + " aeCancel found at time: " + DateTime.Now.ToString("HH:mm:ss"));
                    var ipc =
                        aeCancel.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    ipc.Invoke();
                }
                Thread.Sleep(5000);
            }

            return addOK;
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

            AutomationElement aeClose = AUIUtilities.FindElementById(close, element);
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
    }
}