using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Automation;
using System.Threading;
using TestTools;

namespace EtriccStatisticsProgTest
{
    class StatUtilities
    {
        public static string getFQDN()
        {
            string domainName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
            string hostName = System.Net.Dns.GetHostName();
            string fqdn = "";
            if (!hostName.Contains(domainName))
                fqdn = hostName + "." + domainName;
            else
                fqdn = hostName;
            //System.Windows.Forms.MessageBox.Show(fqdn);

            return fqdn;
        }

        public static AutomationElement ClearMainWindow()
        {
            AutomationElementCollection aeAllMenuItems = null;
            AutomationElement aeItemArrow = null;
            Point ItemArrowPt = new Point();
            AutomationElement aeItemFewButtons = null;
            Point ItemFewButtonsPt = new Point();


            Condition cMenuItems = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.MenuItem);

            AutomationElement aeWindow = GetMainWindow("MainForm");

            TransformPattern tranform =
                     aeWindow.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
            if (tranform != null)
            {
                tranform.Resize(System.Windows.Forms.SystemInformation.VirtualScreen.Width - 60, 
                    System.Windows.Forms.SystemInformation.VirtualScreen.Height - 60);
                Thread.Sleep(1000);
                tranform.Move(0, 0);
            }

            Thread.Sleep(2000);

            aeWindow = GetMainWindow("MainForm");
            aeWindow.SetFocus();
            AUICommon.ClearDisplayedScreens(aeWindow);
            string overflowstripId = "overflowStrip";
            
            AutomationElement aeOverflowStrip = AUIUtilities.FindElementByID(overflowstripId, aeWindow);
            Console.WriteLine("find aeAllMenuItems[k]=");
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while (aeItemArrow == null && mTime.TotalSeconds <= 120)
            {
                Console.WriteLine("aeAllMenuItems[k]=");
                aeAllMenuItems = aeWindow.FindAll(TreeScope.Descendants, cMenuItems);
                Thread.Sleep(3000);
                Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);
                for (int i = 0; i < aeAllMenuItems.Count; i++)
                {
                    Console.WriteLine("aeAllMenuItems[" + i + "]=" + aeAllMenuItems[i].Current.Name);
                    if (aeAllMenuItems[i].Current.Name.Equals(""))
                    {
                        aeItemArrow = aeAllMenuItems[i];
                        Console.WriteLine("1  ClearMainWindow::aeAllMenuItems[" + i + "]=" + aeItemArrow.Current.Name);
                        ItemArrowPt = AUIUtilities.GetElementCenterPoint(aeItemArrow);
                    }
                    else if (aeAllMenuItems[i].Current.Name.Equals("Fewer buttons"))
                    {
                        aeItemFewButtons = aeAllMenuItems[i];
                        Console.WriteLine("2 ClearMainWindow::aeAllMenuItems[" + i + "]=" + aeItemFewButtons.Current.Name);
                        ItemFewButtonsPt = AUIUtilities.GetElementCenterPoint(aeItemFewButtons);
                    }
                }
                mTime = DateTime.Now - mStartTime;
            }

            Input.MoveToAndClick(ItemArrowPt);
            Thread.Sleep(2000);
            string errorMsg = "";
            aeItemFewButtons = GetMenuItemFromElement(aeItemArrow, "Fewer buttons", 120, ref errorMsg);
            if (aeItemFewButtons != null)
                ItemFewButtonsPt = AUIUtilities.GetElementCenterPoint(aeItemFewButtons);

            for (int i = 0; i < 5; i++)
            {
                Input.MoveToAndClick(ItemArrowPt);
                Console.WriteLine("move to ItemFewButtonsPt...");
                Thread.Sleep(2000);
                Input.MoveToAndClick(ItemFewButtonsPt);
                Thread.Sleep(2000);
            }

            return aeWindow;
        }

        static public bool DeleteSelectedProject(AutomationElement aeWindow, AutomationElement aeMyProjNode, ref string errorMsg)
        {
            bool status = false;
           
            AutomationElementCollection aeAllMenuItems = null;
            AutomationElement aeDeleteProject = null;
            Condition cMenuItems = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.MenuItem);
            try
            {
                Point myProjectNodePnt = AUIUtilities.GetElementCenterPoint(aeMyProjNode);
                Point deleteProjectPnt = new Point();
                Input.MoveToAndRightClick(myProjectNodePnt);
                Thread.Sleep(2000);

                Console.WriteLine("find Delete project aeAllMenuItems[k]=");
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                while (aeDeleteProject == null && mTime.TotalSeconds <= 120)
                {
                    Console.WriteLine("aeAllMenuItems[k]=");
                    aeAllMenuItems = aeWindow.FindAll(TreeScope.Descendants, cMenuItems);
                    Thread.Sleep(3000);
                    Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);
                    for (int i = 0; i < aeAllMenuItems.Count; i++)
                    {
                        Console.WriteLine("aeAllMenuItems[" + i + "]=" + aeAllMenuItems[i].Current.Name);
                        if (aeAllMenuItems[i].Current.Name.Equals("Delete project"))
                        {
                            aeDeleteProject = aeAllMenuItems[i];
                            Console.WriteLine("DeleteSelectedProject::aeAllMenuItems[" + i + "]=" + aeDeleteProject.Current.Name);
                            deleteProjectPnt = AUIUtilities.GetElementCenterPoint(aeDeleteProject);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }

                if (aeDeleteProject == null)
                {
                    errorMsg = "DeleteSelectedProject:: delete project MenuItem not found";
                    Console.WriteLine(errorMsg);
                }
                else
                {
                    Input.MoveTo(deleteProjectPnt);
                    Thread.Sleep(3000);
                    Input.MoveToAndClick(deleteProjectPnt);
                    Thread.Sleep(3000);

                    AutomationElement aeDialog = GetTopLevelWindow(aeWindow);
                    // Set a property condition that will be used to find the control.
                    Condition c = new PropertyCondition(AutomationElement.NameProperty, "Yes");
                    AutomationElement aeYesButton = TestTools.AUIUtilities.FindElementByName("Yes", aeDialog);
                    Thread.Sleep(3000);
                    Input.MoveToAndClick(aeYesButton);
                    Thread.Sleep(3000);
                    status = true;
                }
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message + "---------------" + ex.StackTrace;
                Console.WriteLine(errorMsg);
                status = false;
            }

            Console.WriteLine("Delete status=" + status);
            return status;
        }

        static public bool DeleteSelectedXsd(AutomationElement aeWindow, AutomationElement aeMyXsdNode)
        {
            bool status = false;
            Point myProjectNodePnt = AUIUtilities.GetElementCenterPoint(aeMyXsdNode);
            Point deleteXsdPnt = new Point();
            Input.MoveToAndRightClick(myProjectNodePnt);
            Thread.Sleep(2000);

            Console.WriteLine("find Actions aeAllMenuItems[k]=");
            AutomationElementCollection aeAllMenuItems = null;
            AutomationElement aeActions = null;
            AutomationElement aeDeleteXsd = null;

            Condition cMenuItems = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.MenuItem);
            try
            {
                int k = 0;
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                while (aeActions == null && mTime.TotalSeconds <= 120)
                {
                    Console.WriteLine("aeAllMenuItems[k]=");
                    k++;
                    aeAllMenuItems = aeWindow.FindAll(TreeScope.Descendants, cMenuItems);
                    Thread.Sleep(3000);
                    Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);
                    for (int i = 0; i < aeAllMenuItems.Count; i++)
                    {
                        Console.WriteLine("aeAllMenuItems[" + i + "]=" + aeAllMenuItems[i].Current.Name);
                        if (aeAllMenuItems[i].Current.Name.Equals("Actions"))
                        {
                            aeActions = aeAllMenuItems[i];
                            Console.WriteLine("aeAllMenuItems[" + i + "]=" + aeActions.Current.Name);
                            Point actionsPnt = AUIUtilities.GetElementCenterPoint(aeActions);
                            Thread.Sleep(3000);
                            Input.MoveTo(actionsPnt);
                            Thread.Sleep(3000);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }


                mStartTime = DateTime.Now;
                mTime = DateTime.Now - mStartTime;
                while (aeDeleteXsd == null && mTime.TotalSeconds <= 120)
                {
                    Console.WriteLine("aeAllMenuItems[k]=");
                    k++;
                    aeAllMenuItems = aeWindow.FindAll(TreeScope.Descendants, cMenuItems);
                    Thread.Sleep(3000);
                    Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);
                    for (int i = 0; i < aeAllMenuItems.Count; i++)
                    {
                        Console.WriteLine("aeAllMenuItems[" + i + "]=" + aeAllMenuItems[i].Current.Name);
                        if (aeAllMenuItems[i].Current.Name.Equals("Delete Xsd"))
                        {
                            aeDeleteXsd = aeAllMenuItems[i];
                            Console.WriteLine("aeAllMenuItems[" + i + "]=" + aeDeleteXsd.Current.Name);
                            deleteXsdPnt = AUIUtilities.GetElementCenterPoint(aeDeleteXsd);
                            Thread.Sleep(3000);
                            Input.MoveTo(deleteXsdPnt);
                            Thread.Sleep(3000);
                            Input.MoveToAndClick(deleteXsdPnt);
                            Thread.Sleep(3000);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }

                AutomationElement aeDialog = GetTopLevelWindow(aeWindow);

                // Set a property condition that will be used to find the control.
                Condition c = new PropertyCondition(AutomationElement.NameProperty, "Yes");
                AutomationElement aeYesButton = TestTools.AUIUtilities.FindElementByName("Yes", aeDialog);
                Thread.Sleep(3000);
                Input.MoveToAndClick(aeYesButton);
                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "---------------" + ex.StackTrace);
                Thread.Sleep(10000);
                status = false;
            }

            return status;
        }

        /// <summary>
        ///       After right click this node then move to selected menuitem and click this menuitem
        /// </summary>
        /// <param name="aeWindow">MainForm</param>
        /// <param name="aeMyNode">Selected node</param>
        /// <param name="menuItemName">clicked MenuItm displayed after right click on this node </param>
        /// <param name="errorMsg">Error message</param>
        /// <returns></returns>
        static public bool FindAndClickMenuItemOnThisNode(AutomationElement aeWindow, AutomationElement aeMyNode, string menuItemName, ref string errorMsg)
        {
            bool status = false;
          
            try
            {
                AutomationElementCollection aeAllMenuItems = null;
                AutomationElement aeMenuItem = null;
                Condition cMenuItems = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.MenuItem);

                // Right click on this node
                Point myNodePnt = AUIUtilities.GetElementCenterPoint(aeMyNode);
                Point myMenuItemClickPnt = new Point();
                Input.MoveToAndRightClick(myNodePnt);
                Thread.Sleep(2000);

                // Find target menuitem 
                Console.WriteLine("// Find target menuitem =" + menuItemName);
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                while (aeMenuItem == null && mTime.TotalSeconds <= 120)
                {
                    aeAllMenuItems = aeWindow.FindAll(TreeScope.Descendants, cMenuItems);
                    Thread.Sleep(3000);
                    Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);
                    for (int i = 0; i < aeAllMenuItems.Count; i++)
                    {
                        Console.WriteLine(menuItemName+ ": = ?  aeAllMenuItems[" + i + "]=" + aeAllMenuItems[i].Current.Name);
                        if (aeAllMenuItems[i].Current.Name.Equals(menuItemName))
                        {
                            aeMenuItem = aeAllMenuItems[i];
                            Console.WriteLine("FindAndClickMenuItemOnThisNode::aeAllMenuItems[" + i + "]=" + aeMenuItem.Current.Name);
                            myMenuItemClickPnt = AUIUtilities.GetElementCenterPoint(aeMenuItem);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }

                if (aeMenuItem == null)
                {
                    errorMsg = "FindAndClickMenuItemOnThisNode:: MenuItem: "+menuItemName+" not found";
                    Console.WriteLine(errorMsg);
                }
                else
                {
                    Input.MoveTo(myMenuItemClickPnt);
                    Thread.Sleep(3000);
                    Input.MoveToAndClick(myMenuItemClickPnt);
                    Thread.Sleep(3000);
                    status = true;
                }
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message + "---------------" + ex.StackTrace;
                Console.WriteLine(errorMsg);
                status = false;
            }

            return status;
        }

        static public bool FindAndClickMenuItemOnXSDsNode(AutomationElement aeWindow, AutomationElement aeXSDsNode, string menuItemActions, string menuItemSelection, ref string errorMsg)
        {
            bool status = false;

            try
            {
                AutomationElementCollection aeAllMenuItems = null;
                AutomationElement aeMenuItemActions = null;
                AutomationElement aeMenuItemSelection = null;
                Condition cMenuItems = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.MenuItem);

                // Right click on this node
                Point myNodePnt = AUIUtilities.GetElementCenterPoint(aeXSDsNode);
                Point myMenuItemClickPnt = new Point();
                Input.MoveToAndRightClick(myNodePnt);
                Thread.Sleep(2000);

                // Find Actions menuitem 
                Console.WriteLine("// Find Actions menuitem =" + menuItemActions);
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                while (aeMenuItemActions == null && mTime.TotalSeconds <= 120)
                {
                    aeAllMenuItems = aeWindow.FindAll(TreeScope.Descendants, cMenuItems);
                    Thread.Sleep(3000);
                    Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);
                    for (int i = 0; i < aeAllMenuItems.Count; i++)
                    {
                        Console.WriteLine(menuItemActions + ": = ?  aeAllMenuItems[" + i + "]=" + aeAllMenuItems[i].Current.Name);
                        if (aeAllMenuItems[i].Current.Name.Equals(menuItemActions))
                        {
                            aeMenuItemActions = aeAllMenuItems[i];
                            Console.WriteLine("FindAndClickMenuItemOnXSDsNode::aeAllMenuItems[" + i + "]=" + aeMenuItemActions.Current.Name);
                            myMenuItemClickPnt = AUIUtilities.GetElementCenterPoint(aeMenuItemActions);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }

                if (aeMenuItemActions == null)
                {
                    errorMsg = "FindAndClickMenuItemOnXSDsNode:: MenuItem: " + menuItemActions + " not found";
                    Console.WriteLine(errorMsg);
                }
                else
                {
                    Thread.Sleep(1000);
                    Input.MoveTo(myMenuItemClickPnt);
                    Thread.Sleep(3000);
                     // Find Actions menuitem 
                    Console.WriteLine("// Find sELECTION menuitem =" + menuItemActions);
                    mStartTime = DateTime.Now;
                    mTime = DateTime.Now - mStartTime;
                    while (aeMenuItemSelection == null && mTime.TotalSeconds <= 120)
                    {
                        aeAllMenuItems = aeWindow.FindAll(TreeScope.Descendants, cMenuItems);
                        Thread.Sleep(3000);
                        Console.WriteLine("aeAllMenuItems.count=" + aeAllMenuItems.Count);
                        for (int i = 0; i < aeAllMenuItems.Count; i++)
                        {
                            Console.WriteLine(menuItemActions + ": = ?  aeAllMenuItems[" + i + "]=" + aeAllMenuItems[i].Current.Name);
                            if (aeAllMenuItems[i].Current.Name.Equals(menuItemSelection))
                            {
                                aeMenuItemSelection = aeAllMenuItems[i];
                                Console.WriteLine("FindAndClickMenuItemOnXSDsNode::aeAllMenuItems[" + i + "]=" + aeMenuItemSelection.Current.Name);
                                myMenuItemClickPnt = AUIUtilities.GetElementCenterPoint(aeMenuItemSelection);
                                break;
                            }
                        }
                        mTime = DateTime.Now - mStartTime;
                    }

                    if (aeMenuItemSelection == null)
                    {
                        errorMsg = "FindAndClickMenuItemOnXSDsNode:: MenuItem: " + menuItemSelection + " not found";
                        Console.WriteLine(errorMsg);
                    }
                    else
                    {
                        Input.MoveTo(myMenuItemClickPnt);
                        Thread.Sleep(3000);
                        Input.MoveToAndClick(myMenuItemClickPnt);
                        Thread.Sleep(3000);
                        status = true;
                    }
                }
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message + "---------------" + ex.StackTrace;
                Console.WriteLine(errorMsg);
                status = false;
            }

            return status;
        }

        /// <summary>
        /// Retrieves the top-level window that contains the specified UI Automation element.
        /// </summary>
        /// <param name="element">The contained element.</param>
        /// <returns>The containing top-level window element.</returns>
        static public AutomationElement GetTopLevelWindow(AutomationElement element)
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
            }
            while (true);
            return node;
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

        static public AutomationElement GetElementByIdFromParserConfigurationMainWindow(string mainFormId, string elementId, int second, ref string errorMsg)
        {
            AutomationElement aeElement = null;
            AutomationElement aeWindow = GetMainWindow(mainFormId);
            if (aeWindow == null)
            {
                errorMsg = "GetElementByIdFromParserConfigurationMainWindow: no MainForm found";
            }
            else
            {
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                while (aeElement == null && mTime.TotalMinutes <= second)
                {
                    aeElement = TestTools.AUIUtilities.FindElementByID(elementId, aeWindow);
                    Thread.Sleep(2000);
                     mTime = DateTime.Now - mStartTime;
                }


                if (aeElement == null)
                {
                    errorMsg = "After "+second+" second element still not found";
                    Console.WriteLine(errorMsg);
                }

               
            }
            return aeElement;
        }

        static public AutomationElement GetMenuItemFromElement(AutomationElement element, string menuItemId, int seconds, ref string errorMsg)
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

        // elementId = "m_TreeView"
        static public AutomationElement GetReportsTrieView(string mainFormId, string treeViewId, int second, ref string errorMsg)
        {
            AutomationElement aeReportTreeView = null;
            AutomationElement aeReportButton = null;
            AutomationElement aeWindow = GetMainWindow(mainFormId);
            Point repotButtonPt = new Point();
            if (aeWindow == null)
            {
                errorMsg = "GetReportsTrieView: no MainForm found";
            }
            else
            {
                aeWindow.SetFocus();
               
                // find reports button
               Condition cButton = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                while (aeReportButton == null && mTime.TotalSeconds <= 120)
                {
                    AutomationElementCollection aeButtons = aeWindow.FindAll(TreeScope.Descendants, cButton);
                    for (int i = 0; i < aeButtons.Count; i++)
                    {
                        if (aeButtons[i].Current.Name.Equals("Egemin.Epia.Modules.SqlRptServices.MenuNode.Reports")
                            || aeButtons[i].Current.Name.Equals("Reports"))
                        {
                            aeReportButton = aeButtons[i];
                            repotButtonPt = AUIUtilities.GetElementCenterPoint(aeReportButton);
                            Input.MoveToAndClick(repotButtonPt);
                            Console.WriteLine("Click Button aeButton[" + i + "]=" + aeReportButton.Current.Name);
                            break;
                        }
                    }

                    if (aeReportButton == null)
                    {
                        Thread.Sleep(3000);
                        mTime = DateTime.Now - mStartTime;
                    }
                }
            }

            if (aeReportButton == null)
            {
                errorMsg = "GetReportsTrieView: no Report button found";
            }
            else 
            {
                Input.MoveToAndClick(repotButtonPt);
                Thread.Sleep(3000);

                aeWindow = GetMainWindow(mainFormId);
                if (aeWindow == null)
                {
                    errorMsg = "GetReportsTrieView: no MainForm found";
                }
                else
                {
                    // find reports trieview
                    int k = 0;
                    DateTime mStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - mStartTime;
                    System.Windows.Automation.Condition c = new PropertyCondition(
                       AutomationElement.AutomationIdProperty, "m_TreeView", PropertyConditionFlags.IgnoreCase);
                    while (aeReportTreeView == null && mTime.TotalSeconds <= 120)
                    {
                        Console.WriteLine("aeTreeView[k]=" + k);
                        k++;
                        AutomationElementCollection aeTreeViews = aeWindow.FindAll(TreeScope.Element | TreeScope.Descendants, c);
                        for (int i = 0; i < aeTreeViews.Count; i++)
                        {
                            Console.WriteLine("AUIUtilities.GetElementCenterPoint( aeTreeViews[i])[" + i + "]=" + AUIUtilities.GetElementCenterPoint(aeTreeViews[i]));
                            if (AUIUtilities.GetElementCenterPoint(aeTreeViews[i]) != null)
                            {
                                aeReportTreeView = aeTreeViews[i];
                                Console.WriteLine("aeTreeViews[i][" + i + "]=" + aeReportTreeView.Current.Name);
                                break;
                            }
                        }

                        if (aeReportTreeView == null)
                        {
                            Thread.Sleep(3000);
                            mTime = DateTime.Now - mStartTime;
                        }
                    }
                }
            }

            return aeReportTreeView;
        }
       
        static public bool FindPerformanceFinalReport(AutomationElement aeTreeView, string ReportType, string reportGroup, string reportName, ref string errorMsg)
        {
            bool result = true;
            try
            {
                AutomationElement aeEtriccNode = null;
                Point EtriccNodePt = new Point();
                TreeWalker walker = TreeWalker.ControlViewWalker;
                aeEtriccNode = walker.GetFirstChild(aeTreeView);
                if (aeEtriccNode != null)
                {
                    EtriccNodePt = AUIUtilities.GetElementCenterPoint(aeEtriccNode);
                    Console.WriteLine("aeEtriccNode node found name is: " + aeEtriccNode.Current.Name);
                    Thread.Sleep(3000);
                    result = true;
                }
                else
                {
                    errorMsg = "aeEtriccNode not found : ";
                    Console.WriteLine("aeEtriccNode not found : ");
                    Console.WriteLine(errorMsg);
                    result = false;
                }

                AutomationElement aePerformance = null;
                Point PerformancePt = new Point();
                AutomationElement aeVehicles = null;
                Point VehiclesPt = new Point();
                AutomationElement aeModeOverview = null;
                Point ModeOverviewPt = new Point();

                if (result == true)
                {
                    Console.WriteLine("\n=== Find " + ReportType + " node ===");         // Performance
                    aePerformance = RefetchNodeTrieView("MainForm", "m_TreeView", ReportType, 120, ref errorMsg);
                    if (aePerformance != null)
                    {
                        PerformancePt = AUIUtilities.GetElementCenterPoint(aePerformance);
                        Console.WriteLine("found ae" + ReportType + " node is displayed: " + aePerformance.Current.Name);
                    }
                    else
                    {
                        Console.WriteLine("ae" + ReportType + " node not displayed: ");
                        Thread.Sleep(2000);
                        Input.MoveToAndDoubleClick(EtriccNodePt);
                    }
                    

                    // aeEtriccNode node is expanded
                    // refetch performance node from treeview
                    aePerformance = RefetchNodeTrieView("MainForm", "m_TreeView", ReportType, 120, ref errorMsg);
                    if (aePerformance == null)
                    {
                        Console.WriteLine(ReportType + " NOT FOUND After aeEtriccNode node is expanded");
                        result = false;
                    }
                    else
                    {
                        PerformancePt = AUIUtilities.GetElementCenterPoint(aePerformance);
                        result = true;
                    }
                }

                // find vehicles node from performance node
                if (result == true)
                {
                    Console.WriteLine("\n=== Find "+reportGroup+"  node ===");
                    aeVehicles = RefetchNodeTrieView("MainForm", "m_TreeView", reportGroup, 120, ref errorMsg);
                    if (aeVehicles != null)
                    {
                        VehiclesPt = AUIUtilities.GetElementCenterPoint(aeVehicles);
                        Console.WriteLine("found "+reportGroup+" node is displayed: " + aeVehicles.Current.Name);
                    }
                    else
                    {
                        Console.WriteLine("ae"+reportGroup+" not displayed: ");
                        Thread.Sleep(2000);
                        Input.MoveToAndDoubleClick(PerformancePt);
                    }

                    // performance node is expanded
                    // refetch Vehicles node from treeview
                    aeVehicles = RefetchNodeTrieView("MainForm", "m_TreeView", reportGroup, 120, ref errorMsg);
                    if (aeVehicles == null)
                    {
                        Console.WriteLine(reportGroup+ " NOT FOUND After " + ReportType +" node is expanded");
                        result = false;
                    }
                    else
                    {
                        VehiclesPt = AUIUtilities.GetElementCenterPoint(aeVehicles);
                        result = true;
                    }
                }
                    
                // find modeOverview node from Vehicles node
                if (result == true)
                {
                    Console.WriteLine("\n=== Find "+reportName+" node ===");
                    aeModeOverview = RefetchNodeTrieView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                    if (aeModeOverview != null)
                    {
                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeVehicles);
                        Console.WriteLine("found "+reportName+" node is displayed: " + aeModeOverview.Current.Name);
                    }
                    else
                    {
                        Console.WriteLine(reportName +  " not displayed: ");
                        Thread.Sleep(2000);
                        Input.MoveToAndDoubleClick(VehiclesPt);
                    }


                    // refetch ModeOverview node from treeview
                    aeModeOverview = RefetchNodeTrieView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                    if (aeModeOverview == null)
                    {
                        Console.WriteLine(reportName + " NOT FOUND After " + reportGroup + " node is expanded");
                        result = false;
                    }
                    else
                    {
                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeModeOverview);
                        Thread.Sleep(2000);
                        Input.MoveToAndDoubleClick(ModeOverviewPt);
                        result = true;
                    }
                    
                }
            }
            catch (Exception ex)
            {
                result = false;
                errorMsg = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine(reportName + " Fatal error: " + errorMsg);
            }

            return result;

        }

        static public bool FindAnalysisFinalReport(AutomationElement aeTreeView, string ReportType, string reportName, ref string errorMsg)
        {
            bool result = true;
            try
            {
                AutomationElement aeEtriccNode = null;
                Point EtriccNodePt = new Point();
                TreeWalker walker = TreeWalker.ControlViewWalker;
                aeEtriccNode = walker.GetFirstChild(aeTreeView);
                if (aeEtriccNode != null)
                {
                    EtriccNodePt = AUIUtilities.GetElementCenterPoint(aeEtriccNode);
                    Console.WriteLine("aeEtriccNode node found name is: " + aeEtriccNode.Current.Name);
                    Thread.Sleep(3000);
                    result = true;
                }
                else
                {
                    errorMsg = "aeEtriccNode not found : ";
                    Console.WriteLine("aeEtriccNode not found : ");
                    Console.WriteLine(errorMsg);
                    result = false;
                }

                AutomationElement aePerformance = null;
                Point PerformancePt = new Point();
                AutomationElement aeModeOverview = null;
                Point ModeOverviewPt = new Point();

                if (result == true)
                {
                    Console.WriteLine("\n=== Find " + ReportType + " node ===");         // Performance
                    aePerformance = RefetchNodeTrieView("MainForm", "m_TreeView", ReportType, 120, ref errorMsg);
                    if (aePerformance != null)
                    {
                        PerformancePt = AUIUtilities.GetElementCenterPoint(aePerformance);
                        Console.WriteLine("found ae" + ReportType + " node is displayed: " + aePerformance.Current.Name);
                    }
                    else
                    {
                        Console.WriteLine("ae" + ReportType + " node not displayed: ");
                        Thread.Sleep(2000);
                        Input.MoveToAndDoubleClick(EtriccNodePt);
                    }

                    // aeEtriccNode node is expanded
                    // refetch performance node from treeview
                    aePerformance = RefetchNodeTrieView("MainForm", "m_TreeView", ReportType, 120, ref errorMsg);
                    if (aePerformance == null)
                    {
                        Console.WriteLine(ReportType + " NOT FOUND After aeEtriccNode node is expanded");
                        result = false;
                    }
                    else
                    {
                        PerformancePt = AUIUtilities.GetElementCenterPoint(aePerformance);
                        result = true;
                    }
                }

                // find reportName node from Analysis node
                if (result == true)
                {
                    DateTime startTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - startTime;
                    Console.WriteLine("\n===try TO Find " + reportName + "  node ===");
                    int keer = 0;
                    while (aeModeOverview == null && mTime.TotalSeconds < 300)
                    {
                        try
                        {
                            Console.WriteLine("\n=== Find " + reportName + "  node ===" + keer++);
                            aeModeOverview = RefetchNodeTrieView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                            if (aeModeOverview != null)
                            {
                                ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeModeOverview);
                                Console.WriteLine("found " + reportName + " node is displayed: " + aeModeOverview.Current.Name);
                            }
                            else
                            {
                                Console.WriteLine(reportName + " not displayed: ");
                                Thread.Sleep(2000);
                                Input.MoveToAndDoubleClick(PerformancePt);
                            }

                            // refetch reportName node from treeview
                            aeModeOverview = RefetchNodeTrieView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                            if (aeModeOverview == null)
                            {
                                Console.WriteLine(reportName + " NOT FOUND After " + ReportType + " node is expanded");
                                mTime = DateTime.Now - startTime;
                                result = false;
                            }
                            else
                            {
                                ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeModeOverview);
                                Thread.Sleep(2000);
                                Input.MoveToAndDoubleClick(ModeOverviewPt);
                                result = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            aeModeOverview = null;
                        }
                        finally 
                        {
                            Thread.Sleep(3000);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result = false;
                errorMsg = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine("Fatal error: " + errorMsg);
            }
            return result;
        }

        static public bool FindGraphicalViewlReport(AutomationElement aeTreeView, string reportName, ref string errorMsg)
        {
            bool result = true;
            try
            {
                AutomationElement aeEtriccNode = null;
                Point EtriccNodePt = new Point();
                TreeWalker walker = TreeWalker.ControlViewWalker;
                aeEtriccNode = walker.GetFirstChild(aeTreeView);
                if (aeEtriccNode != null)
                {
                    EtriccNodePt = AUIUtilities.GetElementCenterPoint(aeEtriccNode);
                    Console.WriteLine("aeEtriccNode node found name is: " + aeEtriccNode.Current.Name);
                    Thread.Sleep(3000);
                    result = true;
                }
                else
                {
                    errorMsg = "aeEtriccNode not found : ";
                    Console.WriteLine("aeEtriccNode not found : ");
                    Console.WriteLine(errorMsg);
                    result = false;
                }

                AutomationElement aeModeOverview = null;
                Point ModeOverviewPt = new Point();

              
                // find reportName node from Vehicles node
                if (result == true)
                {
                    Console.WriteLine("\n=== Find "+reportName +" node ===");
                    aeModeOverview = RefetchNodeTrieView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                    if (aeModeOverview != null)
                    {
                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeEtriccNode);
                        Console.WriteLine("found "+reportName+" node is displayed: " + aeModeOverview.Current.Name);
                    }
                    else
                    {
                        Console.WriteLine(reportName+ " not displayed: ");
                        Thread.Sleep(2000);
                        Input.MoveToAndDoubleClick(EtriccNodePt);
                    }


                    // refetch ModeOverview node from treeview
                    aeModeOverview = RefetchNodeTrieView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                    if (aeModeOverview == null)
                    {
                        Console.WriteLine(reportName + " NOT FOUND After " + aeEtriccNode.Current.Name + " node is expanded");
                        result = false;
                    }
                    else
                    {
                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeModeOverview);
                        Thread.Sleep(2000);
                        Input.MoveToAndDoubleClick(ModeOverviewPt);
                        result = true;
                    }

                }
            }
            catch (Exception ex)
            {
                result = false;
                errorMsg = ex.Message + "----: " + ex.StackTrace;
                Console.WriteLine(reportName+ " Fatal error: " + errorMsg);
            }

            return result;

        }

        // elementId = "m_TreeView"
        static public AutomationElement RefetchNodeTrieView(string mainFormId, string elementId, string node, int second, ref string errorMsg)
        {
            AutomationElement aeReportTreeView = null;
            AutomationElement aeWindow = GetMainWindow(mainFormId);
            if (aeWindow == null)
            {
                errorMsg = "RefetchNodeTrieView: no MainForm found";
            }
            else
            {
                aeWindow.SetFocus();
                // refetch reports trieview
                int k = 0;
                DateTime mStartTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mStartTime;
                System.Windows.Automation.Condition c = new PropertyCondition(
                   AutomationElement.AutomationIdProperty, "m_TreeView", PropertyConditionFlags.IgnoreCase);
                while (aeReportTreeView == null && mTime.TotalSeconds <= 120)
                {
                    Console.WriteLine("aeTreeView[k]=" + k);
                    k++;
                    AutomationElementCollection aeTreeViews = aeWindow.FindAll(TreeScope.Element | TreeScope.Descendants, c);
                    Thread.Sleep(3000);
                    for (int i = 0; i < aeTreeViews.Count; i++)
                    {
                        Console.WriteLine("AUIUtilities.GetElementCenterPoint( aeTreeViews[i])[" + i + "]=" + AUIUtilities.GetElementCenterPoint(aeTreeViews[i]));
                        if (AUIUtilities.GetElementCenterPoint(aeTreeViews[i]) != null)
                        {
                            aeReportTreeView = aeTreeViews[i];
                            Console.WriteLine("aeTreeViews[i][" + i + "]=" + aeReportTreeView.Current.Name);
                            break;
                        }
                    }
                    mTime = DateTime.Now - mStartTime;
                }
               
            }

            System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
            // treeviewNode is nodeNme like Overview
            AutomationElement aeNode = AUICommon.WalkEnabledElements(aeReportTreeView, treeNode, node);
            if (aeNode != null)
                Console.WriteLine("found aeNodeLink node name is: " + aeNode.Current.Name);
            else
            {
                errorMsg = node + " not found";
                Console.WriteLine(errorMsg );
                Thread.Sleep(3000);
            }
            return aeNode;
        }

        static public AutomationElement GetReportWindow(string reportWindowName, ref string errorMsg)
        {
            AutomationElement aeReportWindow = null;
            AutomationElementCollection aeAllWindows = null;
            bool result = true;

            // for graphical view, report name is Status: graphical view but window title is System overview
            // reportWindowName should be "System overview"
            if (reportWindowName.IndexOf("graphical view") > 0)
                reportWindowName = "System overview";

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

        static public AutomationElement GetLoadedReportWindow(AutomationElement aeReportWindow, string reportWindowName, string fromDate, string toDate, ref string errorMsg)
        {
            bool status = true;

            AutomationElement aeLoadedReportWindow = null;
            string ParamAreaId = "rsParams";
            string ViewReportButtionId = "viewReport";
            string FromDateName = "FromDate"; //ControlType.Edit
            string ToDateName = "ToDate"; //ControlType.Edit

            Console.WriteLine("GetLoadedReportWindow:: current window is: " + aeReportWindow.Current.Name);
            /*
            // Set a property condition that will be used to find the control.
            System.Windows.Automation.Condition c = new PropertyCondition(
            AutomationElement.AutomationIdProperty, ParamAreaId, PropertyConditionFlags.IgnoreCase);
            AutomationElementCollection aeAlls = aeReportWindow.FindAll(TreeScope.Descendants, c);
            Console.WriteLine("aeAlls.Count: " + aeAlls.Count);
            */

            AutomationElement aeParameters = AUIUtilities.FindElementByName("Parameters", aeReportWindow);
            if (aeParameters == null)
            {
                Console.WriteLine("aeParameters not found: " + "Parameters");
                status = false;
            }
            else
            {
                status = true;
            }
           

            AutomationElement aeParamArea = AUIUtilities.FindElementByID(ParamAreaId, aeReportWindow);
            if (aeParamArea == null)
            {
                Console.WriteLine("Param Area not found: " + ParamAreaId);
                status = false;
            }
            else
            {
                status = true;
            }

            if (status == true)
            {
                AutomationElement aeFromDateName = AUIUtilities.FindElementByName(FromDateName, aeReportWindow);
                if (aeFromDateName == null)
                {
                    Console.WriteLine("FromDateName not found: " + FromDateName);
                    status = false;
                }
                else
                {
                    aeFromDateName.SetFocus();
                    Thread.Sleep(1000);
                    ValuePattern vp = aeFromDateName.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    vp.SetValue(fromDate);
                    status = true;
                }
            }

            if (status == true)
            {
                if (reportWindowName.StartsWith("Status: overview"))
                {
                    AutomationElement aeFromDateName = AUIUtilities.FindElementByName(FromDateName, aeReportWindow);
                    if (aeFromDateName == null)
                    {
                        Console.WriteLine("FromDateName not found: " + FromDateName);
                        status = false;
                    }
                    else
                    {
                        aeFromDateName.SetFocus();
                        Thread.Sleep(1000);
                        ValuePattern vp = aeFromDateName.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        vp.SetValue(fromDate);
                        status = true;
                    }
                }
            }
           

            if (status == true)
            {
                Console.WriteLine("ToDateName Set value: " + ToDateName+"    to date value:  "+toDate);
                Thread.Sleep(5000);
                AutomationElement aeToDateName = AUIUtilities.FindElementByName(ToDateName, aeReportWindow);
                if (aeToDateName == null)
                {
                    Console.WriteLine("ToDateName not found: " + ToDateName);
                    status = false;
                }
                else
                {
                    Point pt = AUIUtilities.GetElementCenterPoint(aeToDateName);
                    //aeToDateName.SetFocus();
                    Console.WriteLine("-------------------------- ToDateName found: " + ToDateName);
                    Console.WriteLine("fill value : " + toDate);
                    Input.MoveToAndDoubleClick(pt);
                    Console.WriteLine("current value aeToDateName.Current.IsKeyboardFocusable: " + aeToDateName.Current.IsKeyboardFocusable);
                    aeToDateName.SetFocus();
                    Thread.Sleep(2000);
                    System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                    Thread.Sleep(1000);
                    System.Windows.Forms.SendKeys.SendWait(toDate);
                    Thread.Sleep(1000);
                    // Check Field value
                    TextPattern tp = (TextPattern)aeToDateName.GetCurrentPattern(TextPattern.Pattern);
                    Thread.Sleep(1000);
                    string v = tp.DocumentRange.GetText(-1).Trim();
                    Console.WriteLine("filled text is : " + v);
                    Thread.Sleep(1000);
                    status = true;
                }
            }

           if (status == true)
           {
               Thread.Sleep(2000);
               AutomationElement aeVeiwReport = AUIUtilities.FindElementByID(ViewReportButtionId, aeReportWindow);
               if (aeVeiwReport == null)
               {
                   Console.WriteLine("aeVeiwReport not found: " + ViewReportButtionId);
                   status = false;
               }
               else
               {
                   Point pt = AUIUtilities.GetElementCenterPoint(aeVeiwReport);
                   Thread.Sleep(1000);
                   Input.MoveToAndClick(pt);
                   //Thread.Sleep(1000);
                   //Input.MoveToAndClick(pt);
                   //Thread.Sleep(1000);
                   //Input.MoveToAndClick(pt);
                   status = true;
               }
           }
  
            Thread.Sleep(10000);

            if (status == true)
            {
                aeLoadedReportWindow = GetReportWindow(reportWindowName, ref errorMsg);
                if (aeLoadedReportWindow == null)
                {
                    Console.WriteLine("aeLoadedReportWindow not found: " + reportWindowName);
                    status = false;
                }
                else
                {
                    Console.WriteLine("aeLoadedReportWindow found: " + reportWindowName);
                    status = true;
                }
            }

            return aeLoadedReportWindow;
        }

        static public void GetReportPerformanceVehiclesTestData(string reportName,  ref string fromDate, ref string toDate, ref string validateValue)
        {
            switch (reportName)
            {
                case ReportName.PERFORMANCE_VEHICLES_ModeOverview:
                    validateValue = "02:01:13";
                    break;
                case ReportName.PERFORMANCE_VEHICLES_StateOverview:
                    validateValue = "01:48:39";
                    break;
            }
        }

        static public void GetReportPerformanceTransportsTestData(string reportName, ref string fromDate, ref string toDate, ref string validateValue)
        {
            switch (reportName)
            {
                case ReportName.PERFORMANCE_VEHICLES_ModeOverview:
                    validateValue = "02:01:13";
                    break;
                case ReportName.PERFORMANCE_VEHICLES_StateOverview:
                    validateValue = "01:48:39";
                    break;
            }

        }

        static public void GetReportAnalysisTestData(string reportName, string currentProject, ref string fromDate, ref string toDate, ref string validateValue)
        {
            switch (reportName)
            {
                case ReportName.ANALYSIS_ProjectActivation:
                    if (currentProject.ToLower().Equals("demo"))
                        validateValue = "02:01:13";
                    else
                    {
                        fromDate = "5/9/2011";
                        toDate = "5/24/2011";
                    }
                    break;
                case ReportName.ANALYSIS_TransportLookupBySrcDstGroup:
                    if (currentProject.ToLower().Equals("demo"))
                        validateValue = "02:01:13";
                    else
                    {
                        fromDate = "5/9/2011";
                        toDate = "5/24/2011";
                    }
                    break;
            }
        }

        static public bool ValidateReportPerformanceVehiclesReport(string reportName, string fromDate, string toDate, string validateValue, ref string errorMsg)
        {
            bool status = false;

            AutomationElement aeReportWindow = null;
            AutomationElement aeLoadedReportWindow = null;
            Console.WriteLine("--------------  Get report window ---------------------------wait 5 second: ");
            Thread.Sleep(20000);
            aeReportWindow = StatUtilities.GetReportWindow(reportName, ref errorMsg);
            if (aeReportWindow == null)
            {
                Console.WriteLine(errorMsg);
                status = false;
            }
            else
            {
                Console.WriteLine("Get loaded report window---------------------------");
                // for graphical view, report name is Status: graphical view but window title is System overview
                // AND had no Param area, so no reloaded windows
                if (reportName.IndexOf("graphical view") > 0
                    || reportName.IndexOf("history") > 0)
                {
                    status = true;
                }
                else
                {
                    aeLoadedReportWindow = StatUtilities.GetLoadedReportWindow(aeReportWindow, reportName, fromDate, toDate, ref errorMsg);
                    if (aeLoadedReportWindow == null)
                    {
                        Console.WriteLine(errorMsg);
                        status = false;
                    }
                    else
                    {
                        status = true;
                        switch (reportName)
                        {
                            case ReportName.PERFORMANCE_VEHICLES_ModeOverview:
                            case ReportName.PERFORMANCE_VEHICLES_StateOverview:
                                status = ReportDataValidation(aeLoadedReportWindow, validateValue, ref errorMsg);
                                break;
                        }
                    }
                }
            }
            Console.WriteLine("ValidateReportPerformanceVehiclesReport:: status: "+status);
            Thread.Sleep(5000);
            return status;
        }

        static public bool ReportDataValidation(AutomationElement aeLoadedReportWindow, string validateValue, ref string errorMsg)
        {
            bool status = false;
            // validate field
            // find report text 
            AutomationElementCollection aeAllTexts = null;
            System.Windows.Automation.Condition cText = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Text);

            int k = 0;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            bool valueFound = false;
            while (valueFound == false && mTime.TotalSeconds <= 120)
            {
                Console.WriteLine("aeText[k]=");
                k++;
                aeAllTexts = aeLoadedReportWindow.FindAll(TreeScope.Descendants, cText);
                Thread.Sleep(3000);
                ValuePattern vp = null;
                string textValue;
                Console.WriteLine("aeAllTexts.Count=" + aeAllTexts.Count);
                for (int i = 0; i < aeAllTexts.Count; i++)
                {
                    AutomationPattern[] patterns = aeAllTexts[i].GetSupportedPatterns();
                    foreach (AutomationPattern pattern in patterns)
                    {
                        Console.WriteLine(i + " ProgramaticName=" + pattern.ProgrammaticName);
                        Console.WriteLine(i + " PatternName=" + Automation.PatternName(pattern));
                        vp = aeAllTexts[i].GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        textValue = vp.Current.Value;
                        Console.WriteLine("aeAllTexts[" + i + "]=" + textValue);
                        if (textValue.Equals(validateValue))
                        {
                            valueFound = true;
                            Console.WriteLine("text value found " + validateValue);
                            i = aeAllTexts.Count;
                            Thread.Sleep(10000);
                            break;
                        }
                    }

                    if (valueFound == true)
                        break;

                    Thread.Sleep(500);
                }
                mTime = DateTime.Now - mStartTime;
            }

            if (valueFound == false)
            {
                errorMsg = "validation value not found : " + validateValue;
                status = false;
            }
            else
                status = true;

            return status;
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
                    ErrorMSG = ErrorMSG + "\n"+ aeErrorText.Current.Name;
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

        public static void ErrorMicrosoftNETFrameWorkWindowHandling(AutomationElement element, ref string ErrorMSG)
        {
            string close = "Quit";
            string error = string.Empty;
            AutomationElement aeError1 = AUIUtilities.FindElementByType(ControlType.Text, element);
            if (aeError1 == null)
            {
                error = "Error Message Element not Fund";
                Console.WriteLine(error);
                return;
            }
            else
            {
                ErrorMSG = "Microsoft .Net Framework "+ aeError1.Current.Name;
                Console.WriteLine("aeError is found ------------:");
                //AutomationElement aeErrorText = AUIUtilities.FindElementByID("m_LblErrorText", element);
                //if (aeErrorText != null)
                //{
                //    ErrorMSG = ErrorMSG + "\n" + aeErrorText.Current.Name;
                //}
            }


            Console.WriteLine("Error Msg is ------------:" + ErrorMSG);

            AutomationElement aeClose = AUIUtilities.FindElementByName(close, element);
            if (aeClose == null)
            {
                error = "Quit button element not Found";
                Console.WriteLine(error);
                return;
            }
            else
            {
                Console.WriteLine("aeCuit is found ------------:");
            }

            Thread.Sleep(1000);
            InvokePattern ivp = (InvokePattern)aeClose.GetCurrentPattern(InvokePattern.Pattern);
            ivp.Invoke();
        }
    }
}
