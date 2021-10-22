using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using TestTools;

namespace TFSQATestTools
{
    public class ProjBasicUI
    {
        static public bool IsThisBuildBeforeTestCaseDate(string sBuildNr, string teamProject, string buildDef, DateTime testCaseCreateDate, ref string sErrorMessage)
        {
            bool isBefore = false;
            if (sBuildNr.IndexOf("Hotfix") >= 0) //string buildnr = "Hotfix 4.3.2.1 of Epia.Production.Hotfix_20120405.1";
            {
                //Release 4.4.4 of Epia.Production.Release_20120731.1
                string releaseBuildNr = DeployUtilities.GetReleaseBuildNrFromThisHotfixBuild(sBuildNr);
                DateTime dt = DeployUtilities.GetDateCompletedOfThisBuild(teamProject, buildDef, releaseBuildNr);
                if (dt < testCaseCreateDate)
                {
                    sErrorMessage = "Release date of this hotfix is earlier then this test case created date, Not test";
                    isBefore=  true;
                }
            }
            else if (sBuildNr.IndexOf("Production.Release") >= 0)
            {
                DateTime dt = DeployUtilities.GetDateCompletedOfThisBuild(teamProject, buildDef, sBuildNr);
                if (dt < testCaseCreateDate)
                {
                    sErrorMessage = " date of this release is earlier then this test case created date, Not test";
                    isBefore = true;
                }
            }
            return isBefore;
        }

        static public AutomationElement GetMainWindowWithinTime(string mainFormId, int seconds)
        {
            AutomationElement aeWindow = null;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while (aeWindow == null && mTime.TotalSeconds <= seconds)
            {
                Console.WriteLine("searching window " + mainFormId + " from All Main aeWindows in "+ seconds +" (secs) .....currently:   " + mTime.TotalSeconds);
                try
                {
                    AutomationElementCollection aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children, 
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        if (aeAllWindows[i].Current.AutomationId.Equals(mainFormId))
                        {
                            aeWindow = aeAllWindows[i];
                            Console.WriteLine("Found: aeWindow[" + i + "]=" + aeWindow.Current.Name);
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

        static public AutomationElement GetMainWindowByNameWithinTime(string mainFormId, int seconds)
        {
            AutomationElement aeWindow = null;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while (aeWindow == null && mTime.TotalSeconds <= seconds)
            {
                Console.WriteLine("searching window " + mainFormId + " from All Main aeWindows ...  ");
                try
                {
                    AutomationElementCollection aeAllWindows = AutomationElement.RootElement.FindAll(TreeScope.Children,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        if (aeAllWindows[i].Current.Name.Equals(mainFormId))
                        {
                            aeWindow = aeAllWindows[i];
                            Console.WriteLine("Found: aeWindow[" + i + "]=" + aeWindow.Current.Name);
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

        static public AutomationElement GetSelectedOverviewWindow(string selectedWindowName, ref string errorMsg)
        {
            AutomationElement aeSelectedPane = null;
            AutomationElement aeSelectedWindow = null;
            Console.WriteLine("GetSelectedOverviewWindow: " + selectedWindowName);
            AutomationElement aeWindow = GetMainWindowWithinTime("MainForm", 10);
            if (aeWindow != null)
            {
                DateTime sStartTime = DateTime.Now;
                TimeSpan sTime = DateTime.Now - sStartTime;
                while (aeSelectedWindow == null && sTime.TotalSeconds < 120)
                {
                    AutomationElementCollection aeAllSelectedPanes = aeWindow.FindAll(TreeScope.Children,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane));
                    // There are 3 chile panes possible , left pane for menu selection, Sub windows pane with empty pane name
                    Console.WriteLine(" aeAllSelectedPanes.Count :" + aeAllSelectedPanes.Count);
                    for (int i = 0; i < aeAllSelectedPanes.Count; i++)
                    {
                        string PaneName = aeAllSelectedPanes[i].Current.Name;
                        Console.WriteLine("<PaneName>: " + PaneName + " <PaneName.Length>: " + PaneName.Length + " <Pane.Height>: "
                            + aeAllSelectedPanes[i].Current.BoundingRectangle.Height);

                        Console.WriteLine(" ------------------ <Pane.Width>: " + aeAllSelectedPanes[i].Current.BoundingRectangle.Width);

                        if (PaneName.Length == 0)
                        {
                            // one pane has Height of 5 which include no windows (horizontal), for verticale should > 20
                            // the automationid of left pane is : "windowDockingArea2"
                            if (aeAllSelectedPanes[i].Current.BoundingRectangle.Height > 50 && aeAllSelectedPanes[i].Current.BoundingRectangle.Width > 20)
                            {
                                aeSelectedPane = aeAllSelectedPanes[i];
                                Console.WriteLine("aeSelectedPane with empty name found: " );
                                break;
                            }
                        }
                    }

                    if (aeSelectedPane == null)
                    {
                        errorMsg = "aeSelectedPane with empty pane name not found";
                        Console.WriteLine(errorMsg);
                        break;
                    }
                    else
                    {
                        // find child Pane, normally there are 2 chile panes. the windows pane has empty pane name
                        AutomationElementCollection aeAllWindows = aeSelectedPane.FindAll(TreeScope.Children,
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
                        Console.WriteLine(" aeAllWindows.Count :" + aeAllWindows.Count);
                        for (int i = 0; i < aeAllWindows.Count; i++)
                        {
                            string windowName = aeAllWindows[i].Current.Name;
                            Console.WriteLine("Window name: " + windowName);
                            if (windowName.StartsWith(selectedWindowName))
                            {
                                aeSelectedWindow = aeAllWindows[i];
                                Console.WriteLine("aeSelectedWindow found: " + aeSelectedWindow.Current.Name);
                                break;
                            }
                        }
                        Thread.Sleep(2000);
                        sTime = DateTime.Now - sStartTime;
                        Console.WriteLine("wait " + selectedWindowName + " displayed time is (sec) : " + sTime.TotalSeconds);
                    }
                }
                
                if (aeSelectedWindow == null)
                {
                    errorMsg = "aeSelectedtWindow not found : ";
                    Console.WriteLine(errorMsg);
                }
                else
                {
                    Console.WriteLine("aeSelectedWindow found: " + aeSelectedWindow.Current.Name);
                }
            }
            else
            {
                errorMsg = "MainWindow not found : ";
                Console.WriteLine(errorMsg);
            }
            return aeSelectedWindow;
        }

        public static void SendTextToElement(AutomationElement aeEditElement, string thisText)
        {
            Console.WriteLine("aeCell name is :" + aeEditElement.Current.Name);
            Point pt = AUIUtilities.GetElementCenterPoint(aeEditElement);
            Input.MoveTo(pt);
            Thread.Sleep(200);
            Input.ClickAtPoint(pt);
            Thread.Sleep(500);
            System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
            Thread.Sleep(200);
            System.Windows.Forms.SendKeys.SendWait(thisText);
            Thread.Sleep(500);
        }

        static public bool ShellAction(AutomationElement aeMainForm, string action, ref string errorMsg)
        {
            bool actionOK = true;
            int actionLength = 15;
            if (action.ToLower().Equals("configuration"))
                actionLength = 15;
            else if (action.ToLower().Equals("logoff"))
                actionLength = 40;
            else if (action.ToLower().Equals("shutdown"))
                actionLength = 75;

            Thread.Sleep(1000);
            AutomationElement aeShellBar = AUIUtilities.FindElementByID("_MainForm_Toolbars_Dock_Area_Top", aeMainForm);
            if (aeShellBar == null)
            {
                errorMsg = "_MainForm_Toolbars_Dock_Area_Top " + "not found";
                actionOK = false;
            }
            else
            {
                double x = aeShellBar.Current.BoundingRectangle.Left + 15;
                double y = (aeShellBar.Current.BoundingRectangle.Top + aeShellBar.Current.BoundingRectangle.Bottom) / 2;
                Point shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Console.WriteLine("Click Shell tool bar: ");
                Thread.Sleep(2000);
                double y2 = y + actionLength;
                Point securityPoint = new Point(x, y2);
                Input.MoveTo(securityPoint);

                Console.WriteLine("Click Shell action: "+action);
                Thread.Sleep(2000);

                Input.MoveToAndClick(securityPoint);
                Thread.Sleep(1000);
            }

            return actionOK;
        
        }

        static public bool ShellActionNew(AutomationElement aeMainForm, string action, ref string errorMsg)
        {
            bool actionOK = true;
            int actionLength = 15;
            if (action.ToLower().Equals("configuration"))
                actionLength = 15;
            else if (action.ToLower().Equals("logoff"))
                actionLength = 40;
            else if (action.ToLower().Equals("shutdown"))
                actionLength = 75;

            Thread.Sleep(1000);
            AutomationElement aeShellBar = AUIUtilities.FindElementByID("_MainForm_Toolbars_Dock_Area_Top", aeMainForm);
            if (aeShellBar == null)
            {
                errorMsg = "_MainForm_Toolbars_Dock_Area_Top " + "not found";
                actionOK = false;
            }
            else
            {
                double x = aeShellBar.Current.BoundingRectangle.Left + 15;
                double y = (aeShellBar.Current.BoundingRectangle.Top + aeShellBar.Current.BoundingRectangle.Bottom) / 2;
                Point shellPoint = new Point(x, y);
                Input.MoveTo(shellPoint);
                Thread.Sleep(2000);
                Input.MoveToAndClick(shellPoint);
                Console.WriteLine("Click Shell tool bar: ");
                Thread.Sleep(2000);
                double y2 = y + actionLength;
                Point securityPoint = new Point(x, y2);
                Input.MoveTo(securityPoint);

                Console.WriteLine("Click Shell action: " + action);
                Thread.Sleep(2000);

                Input.MoveToAndClick(securityPoint);
                Thread.Sleep(1000);
            }

            return actionOK;

        }

        static public bool Logon(AutomationElement aeLogonForm, string user, string password, ref string errorMsg)
        {
            bool logonOK = true;
            string UserNameID = "m_TextBoxUsername"; //"ControlType.Edit" Name : "with Windows credentials
            string PasswordID = "m_TextBoxPassword";
            string BtnOKID = "m_BtnOK";

            Console.WriteLine("Application aeSecurityForm name : " + aeLogonForm.Current.Name);
            AutomationElement aePasswordEdit = AUIUtilities.FindElementByID(PasswordID, aeLogonForm);
            if (aePasswordEdit == null)
            {
                errorMsg = PasswordID + " not found";
                logonOK = false;
            }
            else
            {
                ProjBasicUI.SendTextToElement(aePasswordEdit, password);
            }

            if (logonOK == true)
            {
                AutomationElement aeUsernameEdit = AUIUtilities.FindElementByID(UserNameID, aeLogonForm);
                if (aeUsernameEdit == null)
                {
                    errorMsg = UserNameID + " not found";
                    logonOK = false;
                }
                else
                {
                    aeUsernameEdit.SetFocus();
                    Thread.Sleep(2000);

                    System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                    Thread.Sleep(1000);

                    System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                    Thread.Sleep(1000);

                    ProjBasicUI.SendTextToElement(aeUsernameEdit, user);
                }
            }
            // validate input 
            if (logonOK == true)
            {


            }


            if (logonOK == true)
            {
                // Logon into Application // Find Logon OK Button and click 
                Thread.Sleep(1000);
                if (AUIUtilities.FindElementAndClick(BtnOKID, aeLogonForm))
                {
                    Thread.Sleep(3000);
                }
                else
                {
                    errorMsg = "FindElementAndClick failed:" + BtnOKID;
                    logonOK = false;
                }
            }

            return logonOK;
        }

        static public bool ClickButtonInThisElement(string buttonIdOrName, string searchType, AutomationElement aeFromElement, ref string ErrorMsg )
        {
            #region // CLICK BUTTON
            bool clickOK = true;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            AutomationElement aeButton = null;
            while (aeButton == null && mTime.TotalSeconds <= 300)
            {
                Console.WriteLine(buttonIdOrName + " Button is not found yet ....");
                if (searchType.ToLower().StartsWith("name"))
                    aeButton = AUIUtilities.FindElementByName(buttonIdOrName, aeFromElement);
                else if (searchType.ToLower().StartsWith("id"))
                    aeButton = AUIUtilities.FindElementByID(buttonIdOrName, aeFromElement);

                //aeButton = AUIUtilities.FindElementByID(buttonId, aeFromElement);
                Thread.Sleep(10000);
                mTime = DateTime.Now - mStartTime;
            }

            if (aeButton != null)
            {
                Console.WriteLine(buttonIdOrName + " Button is found: ");
                Point buttonPt = AUIUtilities.GetElementCenterPoint(aeButton);
                Thread.Sleep(500);
                Input.MoveToAndClick(buttonPt);
            }
            else
            {
                ErrorMsg = buttonIdOrName + " Button not found";
                Console.WriteLine(ErrorMsg);
                clickOK = false;
            }
            #endregion
            return clickOK;
        }

        static public bool ValidateSystemOverviewMenuItemActionElement(string FormID, string[] menuItemName, ref string errorMsg, string ItemType)
        {
            Console.WriteLine("Start ------- ValidateSystemOverviewMenuItemActionElement: ............"+ menuItemName[0]);
            bool findOK = true;
            #region // mainform --> DropDown -->  menuitem
            AutomationElement aeWindow = null;
            AutomationElement aeDropDownWindow = null;
            AutomationElement aeDropDownWindowMenuItem = null;
            aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 60);
            if (aeWindow != null)
            {
                System.Windows.Automation.Condition cMenu = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, string.Empty),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Menu)
                );

                cMenu =  new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Menu);
                AutomationElementCollection aeAllDropDownWindows = aeWindow.FindAll(TreeScope.Descendants, cMenu);
                // new PropertyCondition(AutomationElement.NameProperty, "DropDown"));
                Console.WriteLine("***------------ aeAllDropDownWindows.Count: " + aeAllDropDownWindows.Count);
               
                if (aeAllDropDownWindows.Count > 0) 
                    aeDropDownWindow = aeAllDropDownWindows[0];

                if (aeDropDownWindow == null)
                {
                    errorMsg = "Find aeDropDownWindow failed:" + "DropDown";
                    Console.WriteLine(errorMsg);
                    findOK = false;
                }
            }
            else
            {
                errorMsg = "Find aeWindow failed:" + "MainForm";
                Console.WriteLine(errorMsg);
                findOK = false;
            }


            // find menuItem
            if (findOK == true)
            {
                System.Windows.Automation.Condition cMenuItem = new AndCondition(
                   new PropertyCondition(AutomationElement.NameProperty, menuItemName[0]),
                   new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
               );

                aeDropDownWindowMenuItem = aeDropDownWindow.FindFirst(TreeScope.Children, cMenuItem);
                if (aeDropDownWindowMenuItem == null)
                {
                    errorMsg = " aeDropDownWindowMenuItem -->   NOT Found " + menuItemName[0];
                    Console.WriteLine(errorMsg);
                    findOK = false;
                }
                else
                {
                    Console.WriteLine("Before Move2...wait 30 sec");
                    Thread.Sleep(1000);
                    Input.MoveTo(aeDropDownWindowMenuItem);
                    Console.WriteLine("After Move...wait 30 sec");
                    Thread.Sleep(1000);
                }

            }

            //Console.WriteLine("Before Move...wait 30 sec");
            //Thread.Sleep(30000);

            if (findOK == true)
            {
                bool validateStatus = true;
                // find Level 2 menuItem 
                //string[] SystemOverviewAllItems = new string[] { "Agv Traffic", "Show Locked Segments", "Show Locked Track", "Show Requested Track", "Show Leave Track",
                //       "Show Hull"
                //    };
                string[] Level2MenuItems = new string[menuItemName.Length - 1];
                for (int i = 0; i < menuItemName.Length-1; i++)
                {
                    Level2MenuItems[i] = menuItemName[i + 1];
                }
                
                validateStatus = ValidateMenuItemFromElement(aeDropDownWindowMenuItem, Level2MenuItems, "name", 120, ref errorMsg);
                if (validateStatus == false)
                {
                    errorMsg = " error message -->  " + errorMsg;
                    Console.WriteLine(errorMsg);
                    findOK = false;
                }
                else
                {
                    Console.WriteLine("Before Move2...wait 5 sec");
                    //Thread.Sleep(30000);
                    //Input.MoveTo(aeDropDownWindow);
                    Console.WriteLine("After Move...wait 5 sec");
                    //Thread.Sleep(30000);
                }
                    
            }
            #endregion
            return findOK;
        }

        static public bool ValidateSystemOverviewCheckboxMenuItems(string FormID, string[] menuItemName, ref string errorMsg, string ItemType)
        {
            bool findOK = true;
            #region // mainform --> checkbox -->  menuitem
            AutomationElement aeWindow = null;
            aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 60);
            if (aeWindow != null)
            {
                AutomationElementCollection aeCol = aeWindow.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox));
                //AutomationElement aeLayersBtn = AUIUtilities.FindElementByType( ControlType.ListItem, aeWindow);
                if (aeCol == null || aeCol.Count < menuItemName.Length)
                {
                    errorMsg = " error message --> Layers button not found ";
                    Console.WriteLine(errorMsg);
                    findOK = false;
                }
                else
                {
                    Console.WriteLine("----     aeCol.Count: " + aeCol.Count);
                    for (int i = 0; i < aeCol.Count; i++)
                    {
                        Console.WriteLine("----     Checkbox  Name: " + aeCol[i].Current.Name);
                        //Input.MoveTo(new Point(aeCol[i].Current.BoundingRectangle.TopLeft.X, aeCol[i].Current.BoundingRectangle.TopLeft.Y));
                        //Thread.Sleep(500);
                    }
                    Console.WriteLine(" --> Layers button found " + aeCol.Count);
                    //var items = aeLayersBtn.FindAll(TreeScope.Element, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
                    for (int i = 0; i < menuItemName.Length; i++)
                    {
                        Console.WriteLine("Checkbox  Name: " + aeCol[i].Current.Name);
                        if (aeCol[i].Current.Name.Equals(menuItemName[i]))
                        {

                        }
                        else
                        {
                            //aeMenuItem = aeAllMenuItems[i];
                            errorMsg = aeCol[i].Current.Name + " menuitem not equal to  " + menuItemName[i];
                            Console.WriteLine("---   " + errorMsg);
                            findOK = false;
                            break;
                        }
                    }
                }
            }
            else
            {
                errorMsg = "Find aeWindow failed:" + "MainForm";
                Console.WriteLine(errorMsg);
                findOK = false;
            }
            #endregion
            return findOK;
        }


        static public bool ValidateWindowMenuItemActionElement(string FormID, string[] menuItemName, ref string errorMsg, string ItemType)
        {
            Console.WriteLine("Start - ValidateWindowMenuItemActionElement  ----------  ItemType:"+ ItemType);
            bool findOK = true;
            #region // mainform --> DropDown -->  menuitem
            AutomationElement aeWindow = null;
            AutomationElement aeDropDownWindow = null;
            aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 60);
            if (aeWindow != null)
            {
                System.Windows.Automation.Condition cMenu = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "DropDown"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Menu)
                );

                AutomationElementCollection aeAllDropDownWindows = aeWindow.FindAll(TreeScope.Children, cMenu);
                // new PropertyCondition(AutomationElement.NameProperty, "DropDown"));
                Console.WriteLine("*** aeAllDropDownWindows.Count: " + aeAllDropDownWindows.Count);
                for (int i = 0; i < aeAllDropDownWindows.Count; i++)
                {
                    Console.WriteLine("000 aeAllDropDownWindows[i].Current.Name: " + aeAllDropDownWindows[i].Current.Name);
                    if (aeAllDropDownWindows[i].Current.Name.StartsWith("DropDown"))
                    {
                        Console.WriteLine("--- ControlType:" + aeAllDropDownWindows[i].Current.ControlType.ProgrammaticName);
                        aeDropDownWindow = aeAllDropDownWindows[i];
                        break;
                    }
                }

                if (aeDropDownWindow == null)
                {
                    errorMsg = "Find aeDropDownWindow failed:" + "DropDown";
                    Console.WriteLine(errorMsg);
                    findOK = false;
                }
            }
            else
            {
                errorMsg = "Find aeWindow failed:" + "MainForm";
                Console.WriteLine(errorMsg);
                findOK = false;
            }

            if (findOK == true)
            {
                bool validateStatus = true;
                if ( ItemType.StartsWith("All"))
                    validateStatus = ValidateAllMenuItemFromElement(aeDropDownWindow, menuItemName, "name", 120, ref errorMsg);
                else
                    validateStatus = ValidateMenuItemFromElement(aeDropDownWindow, menuItemName, "name", 120, ref errorMsg);

                if (validateStatus == false)
                {
                    errorMsg = " error message -->  " + errorMsg;
                    Console.WriteLine(errorMsg);
                    findOK = false;
                }
            }
            #endregion
            Console.WriteLine("END -------------------------- ValidateWindowMenuItemActionElement -------------- findOK:" + findOK);
            return findOK;
        }

        static public AutomationElement GetWindowMenuItemActionElement(string FormID, string menuItemName, ref string errorMsg)
        {
            bool findOK = true;
            AutomationElement aeMenuItem = null;
            #region // mainform --> DropDown -->  menuitem
            AutomationElement aeWindow = null;
            AutomationElement aeDropDownWindow = null;
            aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 60);
            if (aeWindow != null)
            {
                System.Windows.Automation.Condition cMenu = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "DropDown"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Menu)
                );

                AutomationElementCollection aeAllDropDownWindows = aeWindow.FindAll(TreeScope.Children, cMenu);
                // new PropertyCondition(AutomationElement.NameProperty, "DropDown"));
                Console.WriteLine("*** aeAllDropDownWindows.Count" + aeAllDropDownWindows.Count);
                for (int i = 0; i < aeAllDropDownWindows.Count; i++)
                {
                    Console.WriteLine("000 aeAllDropDownWindows[i].Current.Name" + aeAllDropDownWindows[i].Current.Name);
                    if (aeAllDropDownWindows[i].Current.Name.StartsWith("DropDown"))
                    {
                        Console.WriteLine("--- ControlType" + aeAllDropDownWindows[i].Current.ControlType.ProgrammaticName);
                        aeDropDownWindow = aeAllDropDownWindows[i];
                        break;
                    }
                }

                if (aeDropDownWindow == null)
                {
                    errorMsg = "Find aeDropDownWindow failed:" + "DropDown";
                    Console.WriteLine(errorMsg);
                    findOK = false;
                }
            }
            else
            {
                errorMsg = "Find aeWindow failed:" + "MainForm";
                Console.WriteLine(errorMsg);
                findOK = false;
            }

            if (findOK == true)
            {
                aeMenuItem = GetMenuItemFromElement(aeDropDownWindow, menuItemName, "name", 120, ref errorMsg);
                if (aeMenuItem == null)
                {
                    errorMsg = aeWindow.Current.Name + " aeMenuItem not found -->  " + menuItemName;
                    Console.WriteLine(errorMsg);
                    findOK = false;
                }
            }
            #endregion
            return aeMenuItem;
        }

        /// <summary>
        ///     in normal case, a sub menuitem from a menuitem 
        /// </summary>
        /// <param name="element"></param>
        /// <param name="menuItemId"></param>
        /// <param name="seconds"></param>
        /// <param name="errorMsg"></param>
        /// <returns></returns>
        static public AutomationElement GetMenuItemFromElement(AutomationElement element, string menuItemNameOrID, string SearchType, int seconds, ref string errorMsg)
        {
            AutomationElement aeMenuItem = null;
            string NameOrID = string.Empty;
            Console.WriteLine("GetmenuItem: " + menuItemNameOrID + " --- SearchType:" + SearchType);
            AutomationElementCollection aeAllMenuItems = null;
            System.Windows.Automation.Condition cMenuItems = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem);

            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            try
            {
                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                while (aeMenuItem == null && mTime.TotalSeconds <= 300)
                {
                    aeAllMenuItems = element.FindAll(TreeScope.Descendants, cMenuItems);
                    Thread.Sleep(3000);
                    Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);
                    for (int i = 0; i < aeAllMenuItems.Count; i++)
                    {
                        if (SearchType.ToLower().StartsWith("name"))
                            NameOrID = aeAllMenuItems[i].Current.Name;
                        else if (SearchType.ToLower().StartsWith("id"))
                            NameOrID = aeAllMenuItems[i].Current.AutomationId;

                        Console.WriteLine("menuItem NameOrID: " + NameOrID);

                        if (NameOrID.StartsWith(menuItemNameOrID))
                        {
                            aeMenuItem = aeAllMenuItems[i];
                            Console.WriteLine("aeMenuItem found: " + aeMenuItem.Current.Name);
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

        static public bool ValidateMenuItemFromElement(AutomationElement element, string[] menuItemNameOrID, string SearchType, int seconds, ref string errorMsg)
        {
            Console.WriteLine("-----------------------  Start ---- ValidateMenuItemFromElement: "+ element.Current.Name);
            bool status = true;
            //AutomationElement aeMenuItem = null;
            string NameOrID = string.Empty;
            for (int i = 0; i < menuItemNameOrID.Length; i++)
            {
                Console.WriteLine(i+"de Get Level2 menuItem: " + menuItemNameOrID[i]);
            }
            Console.WriteLine( " --- SearchType:" + SearchType);
            AutomationElementCollection aeAllMenuItems = null;
            System.Windows.Automation.Condition cMenuItems = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem);

            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            try
            {
                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                //while (aeMenuItem == null && mTime.TotalSeconds <= 300)
                //{
                    aeAllMenuItems = element.FindAll(TreeScope.Descendants, cMenuItems);
                    Thread.Sleep(3000);
                    Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);

                    List<string> VisibleMenuitemsList = new List<string>();
                    string hideMenuitem = "hide  menuitem name : ";
                    for (int i = 0; i < aeAllMenuItems.Count; i++)
                    {
                        
                        // check if is there hidden menuitems?
                        try
                        {
                            //x = "xxx0 menuitem name : " + aeAllMenuItems[i].Current.Name;
                            Console.WriteLine("xxx all menuitem name : " + aeAllMenuItems[i].Current.Name );
                            aeAllMenuItems[i].GetClickablePoint(); 
                        }
                        catch (Exception)
                        {
                            hideMenuitem = hideMenuitem + " , "+ aeAllMenuItems[i].Current.Name;
                            continue;
                        }


                        if (SearchType.ToLower().StartsWith("name"))
                            VisibleMenuitemsList.Add(aeAllMenuItems[i].Current.Name);
                        else if (SearchType.ToLower().StartsWith("id"))
                            VisibleMenuitemsList.Add(aeAllMenuItems[i].Current.AutomationId);

                        
                        //Console.WriteLine("xxx menuitem name : " + aeAllMenuItems[i].Current.Name   );
                    }
                    Console.WriteLine(hideMenuitem);

                    Console.WriteLine("total visible menuitems = " +VisibleMenuitemsList.Count);
                    Console.WriteLine("==================================================== 1" );
                    for (int i = 0; i < VisibleMenuitemsList.Count; i++)
                    {
                        Console.WriteLine(i+ " de VisibleMenuitemsList NameOrID: " + VisibleMenuitemsList.ElementAt(i));

                    }


                    Console.WriteLine("==================================================== 2");

                    for (int i = 0; i < menuItemNameOrID.Length; i++)
                    {
                        Console.WriteLine(i + " de menuItemNameOrID: " + menuItemNameOrID[i]);

                    }

                    Console.WriteLine("==================================================== 3");




                    for (int i = 0; i < VisibleMenuitemsList.Count; i++)
                    {
                        Console.WriteLine("VisibleMenuitemsList NameOrID: " + VisibleMenuitemsList.ElementAt(i) );
                        if (VisibleMenuitemsList.ElementAt(i).Equals(menuItemNameOrID[i]))
                        {

                        }
                        else
                        {
                            //aeMenuItem = aeAllMenuItems[i];
                            errorMsg = VisibleMenuitemsList.ElementAt(i) + " menuitem not equal to  " + menuItemNameOrID[i];
                            Console.WriteLine("---   " + errorMsg);
                            status = false;
                            break;
                        }


                    }
                    mTime = DateTime.Now - mStartTime;
                //}
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message + "---------------" + ex.StackTrace;
                Console.WriteLine(errorMsg);
                status = false;
            }

            Console.WriteLine("-----------------------  End ---- ValidateMenuItemFromElement status: " + status);
            return status;
        }

        /// <summary>
        ///  Validate all menu items. also include invisible menuitems
        /// </summary>
        /// <param name="element"></param>
        /// <param name="menuItemNameOrID"></param>
        /// <param name="SearchType"></param>
        /// <param name="seconds"></param>
        /// <param name="errorMsg"></param>
        /// <returns></returns>
        static public bool ValidateAllMenuItemFromElement(AutomationElement element, string[] menuItemNameOrID, string SearchType, int seconds, ref string errorMsg)
        {
            Console.WriteLine("Start ValidateAllMenuItemFromElement ---------- Start ------------- SearchType=" + SearchType);
            bool status = true;
            //AutomationElement aeMenuItem = null;
            string NameOrID = string.Empty;
            Console.WriteLine("GetmenuItem: " + menuItemNameOrID + " --- SearchType:" + SearchType);
            AutomationElementCollection aeAllMenuItems = null;
            System.Windows.Automation.Condition cMenuItems = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem);

            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            try
            {
                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                aeAllMenuItems = element.FindAll(TreeScope.Descendants, cMenuItems);
                Thread.Sleep(3000);
                Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);

                
                for (int i = 0; i < aeAllMenuItems.Count; i++)
                {
                    if (SearchType.ToLower().StartsWith("name"))
                    {
                        NameOrID = aeAllMenuItems[i].Current.Name;
                        Console.WriteLine("aeAllMenuItems["+i+"]=" + aeAllMenuItems[i].Current.Name);
                    }
                    else if (SearchType.ToLower().StartsWith("id"))
                        NameOrID = aeAllMenuItems[i].Current.AutomationId;

                    Console.WriteLine("VisibleMenuitemsList NameOrID: " + NameOrID);
                    if (NameOrID.Equals(menuItemNameOrID[i]))
                    {

                    }
                    else
                    {
                        //aeMenuItem = aeAllMenuItems[i];
                        errorMsg = NameOrID + " menuitem not equal to "+i +" th menuItemNameOrID " + menuItemNameOrID[i];
                        Console.WriteLine("---   " + errorMsg);
                        status = false;
                        break;
                    }
                }
                mTime = DateTime.Now - mStartTime;
                //}
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message + "---------------" + ex.StackTrace;
                Console.WriteLine(errorMsg);
                status = false;
            }

            return status;
        }

        static public AutomationElement GetPopupDialogFromMainWindow(string mainformId, string dialogNameOrID, string dialogSearchType, ref string errorMsg)
        {
            AutomationElement aePopupWindow = null;
            string NameOrID = string.Empty;
            Console.WriteLine("GetPopupDialog: " + dialogNameOrID + " --- SearchType:" + dialogSearchType);
            AutomationElement aeWindow = GetMainWindowWithinTime(mainformId, 10);
            if (aeWindow != null)
            {
                DateTime sStartTime = DateTime.Now;
                TimeSpan sTime = DateTime.Now - sStartTime;
                while (aePopupWindow == null && sTime.TotalSeconds < 300)
                {
                    AutomationElementCollection aeAllWindows = aeWindow.FindAll(TreeScope.Children,
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
                    Console.WriteLine(" aeAllWindows.Count :" + aeAllWindows.Count);
                    for (int i = 0; i < aeAllWindows.Count; i++)
                    {
                        if ( dialogSearchType.ToLower().StartsWith("name"))
                            NameOrID = aeAllWindows[i].Current.Name;
                        else if ( dialogSearchType.ToLower().StartsWith("id"))
                            NameOrID = aeAllWindows[i].Current.AutomationId;

                        Console.WriteLine("Window NameOrID: " + NameOrID);

                        if (NameOrID.StartsWith(dialogNameOrID))
                        {
                            aePopupWindow = aeAllWindows[i];
                            Console.WriteLine("aeSelectedWindow found: " + aePopupWindow.Current.Name);
                            break;
                        }
                    }
                    Thread.Sleep(2000);
                    sTime = DateTime.Now - sStartTime;
                    Console.WriteLine("wait " + dialogNameOrID + " displayed time is (sec) : " + sTime.TotalSeconds);
                    
                }

                if (aePopupWindow == null)
                {
                    errorMsg = "aePopupWindow not found : ";
                    Console.WriteLine(errorMsg);
                }
                else
                {
                    Console.WriteLine("aePopupWindow found: " + aePopupWindow.Current.Name);
                }
            }
            else
            {
                errorMsg = "MainWindow not found : ";
                Console.WriteLine(errorMsg);
            }
            return aePopupWindow;
        }

        static public bool GetThisWindowIsTopMost(AutomationElement aeWindow)
        {
            WindowPattern pattern = (WindowPattern)aeWindow.GetCurrentPattern(WindowPattern.Pattern);
            return pattern.Current.IsTopmost;

        }
    }
}
