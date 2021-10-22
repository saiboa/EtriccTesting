using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using TestTools;
using TFSQATestTools;
using WindowsInput;

namespace QATestEtriccStatistics
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

            AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
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

            aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
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
            aeItemFewButtons = ProjBasicUI.GetMenuItemFromElement(aeItemArrow, "Fewer buttons", "name", 120, ref errorMsg);
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
                        Console.WriteLine(menuItemName + ": = ?  aeAllMenuItems[" + i + "]=" + aeAllMenuItems[i].Current.Name);
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
                    errorMsg = "FindAndClickMenuItemOnThisNode:: MenuItem: " + menuItemName + " not found";
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

        static public AutomationElement GetElementByIdFromParserConfigurationMainWindow(string mainFormId, string elementId, int second, ref string errorMsg)
        {
            AutomationElement aeElement = null;
            AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime(mainFormId, 10);
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
                    errorMsg = "After " + second + " second element still not found";
                    Console.WriteLine(errorMsg);
                }
            }
            return aeElement;
        }
        
        // elementId = "m_TreeView"
        static public AutomationElement GetReportsTrieView(string mainFormId, string treeViewId, int second, ref string errorMsg)
        {
            AutomationElement aeReportTreeView = null;
            AutomationElement aeReportButton = null;
            AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime(mainFormId, 10);
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
                Thread.Sleep(2000);

                aeWindow = ProjBasicUI.GetMainWindowWithinTime(mainFormId, 10);
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

        static public bool FindPerformanceFinalReport(AutomationElement aeTreeView, string reportType, string reportGroup, string reportName, ref string errorMsg)
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
                    Thread.Sleep(1000);
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

                if (reportType == null)   // open system overview report from Etricc node point
                {
                    #region
                    Console.WriteLine("\n=== Find " + reportName + " node ===");
                    aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                    if (aeModeOverview != null)
                    {
                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeEtriccNode);
                        Console.WriteLine("found " + reportName + " node is displayed: " + aeModeOverview.Current.Name);
                    }
                    else
                    {
                        Console.WriteLine(reportName + " not displayed: ");
                        Thread.Sleep(200);
                        Input.MoveToAndDoubleClick(EtriccNodePt);
                    }

                    // refetch ModeOverview node from treeview
                    aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                    if (aeModeOverview == null)
                    {
                        Console.WriteLine(reportName + " NOT FOUND After " + reportType + " node is expanded");
                        result = false;
                    }
                    else
                    {
                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeModeOverview);
                        Thread.Sleep(200);
                        Input.MoveToAndDoubleClick(ModeOverviewPt);
                        result = true;
                    }
                    #endregion
                }
                else
                {
                    Console.WriteLine("\n=== Find " + reportType + " node ===");         // Performance
                    #region
                    aePerformance = RefetchNodeTreeView("MainForm", "m_TreeView", reportType, 120, ref errorMsg);
                    if (aePerformance != null)
                    {
                        PerformancePt = AUIUtilities.GetElementCenterPoint(aePerformance);
                        Console.WriteLine("found ae" + reportType + " node is displayed: " + aePerformance.Current.Name);
                    }
                    else
                    {
                        // expand Etricc Node
                        Console.WriteLine("ae" + reportType + " node not displayed: ");
                        Thread.Sleep(200);
                        Input.MoveToAndDoubleClick(EtriccNodePt);
                    }

                    // aeEtriccNode node is expanded
                    // refetch performance node from treeview
                    aePerformance = RefetchNodeTreeView("MainForm", "m_TreeView", reportType, 120, ref errorMsg);
                    if (aePerformance == null)
                    {
                        Console.WriteLine(reportType + " NOT FOUND After aeEtriccNode node is expanded");
                        result = false;
                    }
                    else
                    {
                        PerformancePt = AUIUtilities.GetElementCenterPoint(aePerformance);
                        result = true;
                    }
                    #endregion
                   
                    // find vehicles node from performance node
                    // find reportGroup node from ReportTyp node
                    Console.WriteLine("\n=== find reportGroup node from ReportTyp node ===");
                    if (result == true)
                    {
                        if (reportGroup == null)   // find report from reportType
                        {
                            Console.WriteLine("\n=== Find " + reportName + " node ===");
                            aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                            if (aeModeOverview != null)
                            {
                                ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aePerformance);
                                Console.WriteLine("found " + reportName + " node is displayed: " + aeModeOverview.Current.Name);
                            }
                            else
                            {
                                Console.WriteLine(reportName + " not displayed: ");
                                Thread.Sleep(200);
                                Input.MoveToAndDoubleClick(PerformancePt);
                            }

                            // refetch ModeOverview node from treeview
                            aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                            if (aeModeOverview == null)
                            {
                                Console.WriteLine(reportName + " NOT FOUND After " + reportType + " node is expanded");
                                result = false;
                            }
                            else
                            {
                                ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeModeOverview);
                                Thread.Sleep(200);
                                Input.MoveToAndDoubleClick(ModeOverviewPt);
                                result = true;
                            }
                        }
                        else  // for Analysis type, there is no report group
                        {
                            Console.WriteLine("\n=== Find " + reportGroup + "  node ===");
                            aeVehicles = RefetchNodeTreeView("MainForm", "m_TreeView", reportGroup, 120, ref errorMsg);
                            if (aeVehicles != null)
                            {
                                VehiclesPt = AUIUtilities.GetElementCenterPoint(aeVehicles);
                                Console.WriteLine("found " + reportGroup + " node is displayed: " + aeVehicles.Current.Name);
                            }
                            else
                            {
                                Console.WriteLine("ae" + reportGroup + " not displayed: ");
                                Thread.Sleep(200);
                                Input.MoveToAndDoubleClick(PerformancePt);
                            }

                            // performance node is expanded
                            // refetch Vehicles node from treeview
                            aeVehicles = RefetchNodeTreeView("MainForm", "m_TreeView", reportGroup, 120, ref errorMsg);
                            if (aeVehicles == null)
                            {
                                Console.WriteLine(reportGroup + " NOT FOUND After " + reportType + " node is expanded");
                                result = false;
                            }
                            else
                            {
                                VehiclesPt = AUIUtilities.GetElementCenterPoint(aeVehicles);
                                result = true;
                            }
                            // find modeOverview node from Vehicles node
                            if (result == true)
                            {
                                if (reportGroup != null)   // for Analysis type, there is no report group
                                {
                                    Console.WriteLine("\n=== Find " + reportName + " node ===");
                                    aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                                    if (aeModeOverview != null)
                                    {
                                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeVehicles);
                                        Console.WriteLine("found " + reportName + " node is displayed: " + aeModeOverview.Current.Name);
                                    }
                                    else
                                    {
                                        Console.WriteLine(reportName + " not displayed: ");
                                        Input.MoveToAndDoubleClick(VehiclesPt);
                                    }

                                    // refetch ModeOverview node from treeview
                                    aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                                    if (aeModeOverview == null)
                                    {
                                        Console.WriteLine(reportName + " NOT FOUND After " + reportGroup + " node is expanded");
                                        result = false;
                                    }
                                    else
                                    {
                                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeModeOverview);
                                        Input.MoveToAndDoubleClick(ModeOverviewPt);
                                        result = true;
                                    }
                                }
                            }
                        }
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

        static public bool FindPerformanceFinalReport(AutomationElement aeTreeView, string reportType, string reportGroup, string monthlyOrHourly, string reportName, ref string errorMsg)
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
                    Thread.Sleep(1000);
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
                AutomationElement aeHorrlyOrDailyOrMonthly = null;
                Point HorrlyOrDailyOrMonthlyPt = new Point();
                AutomationElement aeModeOverview = null;
                Point ModeOverviewPt = new Point();

                if (reportType == null)   // open system overview report from Etricc node point
                {
                    #region
                    Console.WriteLine("\n=== Find " + reportName + " node ===");
                    aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                    if (aeModeOverview != null)
                    {
                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeEtriccNode);
                        Console.WriteLine("found " + reportName + " node is displayed: " + aeModeOverview.Current.Name);
                    }
                    else
                    {
                        Console.WriteLine(reportName + " not displayed: ");
                        Thread.Sleep(200);
                        Input.MoveToAndDoubleClick(EtriccNodePt);
                    }

                    // refetch ReportOverview node from treeview
                    aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                    if (aeModeOverview == null)
                    {
                        Console.WriteLine(reportName + " NOT FOUND After " + reportType + " node is expanded");
                        result = false;
                    }
                    else
                    {
                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeModeOverview);
                        Thread.Sleep(200);
                        Input.MoveToAndDoubleClick(ModeOverviewPt);
                        result = true;
                    }
                    #endregion
                }
                else // other reports   : report type : Performance; Analysis, ...
                {
                    Console.WriteLine("\n=== Find " + reportType + " node ===");         // Performance
                    #region     find // Performance ; Analysis, ...
                    aePerformance = RefetchNodeTreeView("MainForm", "m_TreeView", reportType, 120, ref errorMsg);
                    if (aePerformance != null)
                    {
                        PerformancePt = AUIUtilities.GetElementCenterPoint(aePerformance);
                        Console.WriteLine("found ae" + reportType + " node is displayed: " + aePerformance.Current.Name);
                    }
                    else
                    {
                        // expand Etricc Node
                        Console.WriteLine("ae" + reportType + " node not displayed: ");
                        Thread.Sleep(200);
                        Input.MoveToAndDoubleClick(EtriccNodePt);
                    }

                    // aeEtriccNode node is expanded
                    // refetch performance node from treeview
                    aePerformance = RefetchNodeTreeView("MainForm", "m_TreeView", reportType, 120, ref errorMsg);
                    if (aePerformance == null)
                    {
                        Console.WriteLine(reportType + " NOT FOUND After aeEtriccNode node is expanded");
                        result = false;
                    }
                    else
                    {
                        PerformancePt = AUIUtilities.GetElementCenterPoint(aePerformance);
                        result = true;
                    }
                    #endregion

                    // find vehicles node from performance node
                    // find reportGroup node from ReportTyp node
                    Console.WriteLine("\n=== find reportGroup node from ReportTyp node ===");
                    if (result == true)
                    {
                        if (reportGroup == null)   // find report from reportType; for Analysis type, there is no report group
                        {
                            Console.WriteLine("\n=== Find " + reportName + " node ===");
                            aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                            if (aeModeOverview != null)
                            {
                                ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aePerformance);
                                Console.WriteLine("found " + reportName + " node is displayed: " + aeModeOverview.Current.Name);
                            }
                            else
                            {
                                Console.WriteLine(reportName + " not displayed: ");
                                Thread.Sleep(200);
                                Input.MoveToAndDoubleClick(PerformancePt);
                            }

                            // refetch ModeOverview node from treeview
                            aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                            if (aeModeOverview == null)
                            {
                                Console.WriteLine(reportName + " NOT FOUND After " + reportType + " node is expanded");
                                result = false;
                            }
                            else
                            {
                                ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeModeOverview);
                                Thread.Sleep(200);
                                Input.MoveToAndDoubleClick(ModeOverviewPt);
                                result = true;
                            }
                        }
                        else  // find group Vehicles, Transports, Jobs...
                        {
                            Console.WriteLine("\n=== Find " + reportGroup + "  node ===");
                            aeVehicles = RefetchNodeTreeView("MainForm", "m_TreeView", reportGroup, 120, ref errorMsg);
                            if (aeVehicles != null)
                            {
                                VehiclesPt = AUIUtilities.GetElementCenterPoint(aeVehicles);
                                Console.WriteLine("found " + reportGroup + " node is displayed: " + aeVehicles.Current.Name);
                            }
                            else
                            {
                                Console.WriteLine("ae" + reportGroup + " not displayed: ");
                                Thread.Sleep(200);
                                Input.MoveToAndDoubleClick(PerformancePt);
                            }

                            // performance node is expanded
                            // refetch Vehicles node from treeview
                            aeVehicles = RefetchNodeTreeView("MainForm", "m_TreeView", reportGroup, 120, ref errorMsg);
                            if (aeVehicles == null)
                            {
                                Console.WriteLine(reportGroup + " NOT FOUND After " + reportType + " node is expanded");
                                result = false;
                            }
                            else
                            {
                                VehiclesPt = AUIUtilities.GetElementCenterPoint(aeVehicles);
                                result = true;
                            }

                            // find Hourly or Daily or Monthly node from Jobs node
                            Console.WriteLine("find Hourly or Daily or Monthly node from Jobs node" );
                            if (result == true)
                            {
                                Console.WriteLine("\n==XXXXXXXXx = Find " + monthlyOrHourly + "  node ===");
                                //aeHorrlyOrDailyOrMonthly = RefetchNodeTreeView("MainForm", "m_TreeView", reportType, reportGroup, monthlyOrHourly, 120, ref errorMsg);
                                aeHorrlyOrDailyOrMonthly = RefetchNodeTreeView("MainForm", "m_TreeView", monthlyOrHourly, 120, ref errorMsg);
                                if (aeHorrlyOrDailyOrMonthly != null)
                                {
                                    HorrlyOrDailyOrMonthlyPt = AUIUtilities.GetElementCenterPoint(aeHorrlyOrDailyOrMonthly);
                                    Console.WriteLine("found " + monthlyOrHourly + " node is displayed: " + aeHorrlyOrDailyOrMonthly.Current.Name);
                                }
                                else
                                {
                                    Console.WriteLine("ae" + monthlyOrHourly + " not displayed: ");
                                    Thread.Sleep(200);
                                    Console.WriteLine("Click "+ reportGroup + " node point");
                                    Input.MoveToAndDoubleClick(VehiclesPt);
                                    Thread.Sleep(5000);
                                }

                                Thread.Sleep(5000);
                                // reportGroup node is expanded
                                // refetch monthlyOrHourly node from treeview
                                aeHorrlyOrDailyOrMonthly = RefetchNodeTreeView("MainForm", "m_TreeView", monthlyOrHourly, 120, ref errorMsg);
                                if (aeHorrlyOrDailyOrMonthly == null)
                                {
                                    Console.WriteLine(monthlyOrHourly + " NOT FOUND After aeGroup node is expanded");
                                    result = false;
                                }
                                else
                                {
                                    HorrlyOrDailyOrMonthlyPt = AUIUtilities.GetElementCenterPoint(aeHorrlyOrDailyOrMonthly);
                                    result = true;
                                }                               
                            }

                            // find modeOverview node from Vehicles node
                            if (result == true)
                            {
                                if (reportGroup != null)  
                                {
                                    Console.WriteLine("\n=== Find " + reportName + " node ===");
                                    aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                                    if (aeModeOverview != null)
                                    {
                                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeVehicles);
                                        Console.WriteLine("found " + reportName + " node is displayed: " + aeModeOverview.Current.Name);
                                    }
                                    else
                                    {
                                        Console.WriteLine(reportName + " not displayed: ");
                                        Input.MoveToAndDoubleClick(HorrlyOrDailyOrMonthlyPt);
                                    }

                                    // refetch ModeOverview node from treeview
                                    aeModeOverview = RefetchNodeTreeView("MainForm", "m_TreeView", reportName, 120, ref errorMsg);
                                    if (aeModeOverview == null)
                                    {
                                        Console.WriteLine(reportName + " NOT FOUND After " + monthlyOrHourly + " node is expanded");
                                        result = false;
                                    }
                                    else
                                    {
                                        ModeOverviewPt = AUIUtilities.GetElementCenterPoint(aeModeOverview);
                                        Input.MoveToAndDoubleClick(ModeOverviewPt);
                                        result = true;
                                    }
                                }
                            }
                        }
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

       // elementId = "m_TreeView"
        static public AutomationElement RefetchNodeTreeView(string mainFormId, string elementId, string node, int second, ref string errorMsg)
        {
            AutomationElement aeReportTreeView = null;
            AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime(mainFormId, 10);
            if (aeWindow == null)
            {
                errorMsg = "RefetchNodeTreeView: no MainForm found";
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

            System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
            // treeviewNode is nodeNme like Overview
            AutomationElement aeNode = AUICommon.WalkEnabledElements(aeReportTreeView, treeNode, node);
            if (aeNode != null)
            {
                errorMsg = node + " IS found";
                Console.WriteLine("found aeNodeLink node name is: " + aeNode.Current.Name);
            }
            else
            {
                errorMsg = node + " not found";
                Console.WriteLine(errorMsg);
                Thread.Sleep(3000);
            }
            return aeNode;
        }

        static public AutomationElement GetReportWindow(string reportWindowName, ref string errorMsg)
        {
            AutomationElement aeReportWindow = null;
            bool result = true;

            // for graphical view, report name is Status: graphical view but window title is System overview
            // reportWindowName should be "System overview"
            if (reportWindowName.IndexOf("graphical view") > 0)
                reportWindowName = "System overview";

            Console.WriteLine("GetReportWindow:: ");
            Console.WriteLine("GetMainWindow:: ");
            AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 10);
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

        static public AutomationElement GetLoadedReportWindow2(AutomationElement aeReportWindow, string fromDate, string toDate, ref string errorMsg)
        {
            bool status = true;

            AutomationElement aeLoadedReportWindow = null;
            string ParamAreaId = "rsParams";
            string ViewReportButtionId = "viewReport";
            string FromDateName = "FromDate"; //ControlType.Edit
            string ToDateName = "ToDate"; //ControlType.Edit
            string reportWindowName = aeReportWindow.Current.Name;
            Console.WriteLine("GetLoadedReportWindow:: current window is: " + aeReportWindow.Current.Name);
            /*
            // Set a property condition that will be used to find the control.
            System.Windows.Automation.Condition c = new PropertyCondition(
            AutomationElement.AutomationIdProperty, ParamAreaId, PropertyConditionFlags.IgnoreCase);
            AutomationElementCollection aeAlls = aeReportWindow.FindAll(TreeScope.Descendants, c);
            Console.WriteLine("aeAlls.Count: " + aeAlls.Count);
            */

            try
            {
                Console.WriteLine("- -----------------   aeReportWindow.Current.IsEnabled: " + aeReportWindow.Current.IsEnabled);

                AutomationElement aeParameters = AUIUtilities.FindElementByName("Parameters", aeReportWindow);
                AutomationElement aeParamArea = null;
                if (aeParameters == null)
                {
                    Console.WriteLine("aeParameters not found: " + "Parameters");
                    status = false;
                }
                else
                {
                    aeParamArea = AUIUtilities.FindElementByID(ParamAreaId, aeParameters);
                    if (aeParamArea == null)
                    {
                        Console.WriteLine("Param Area not found: " + ParamAreaId);
                        status = false;
                    }
                }

                if (status == true)
                {
                    AutomationElement aeFromDateName = AUIUtilities.FindElementByName(FromDateName, aeParamArea);
                    if (aeFromDateName == null)
                    {
                        Console.WriteLine("FromDateName not found: " + FromDateName);
                        status = false;
                    }
                    else
                    {
                        ValuePattern vp = aeFromDateName.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        vp.SetValue(fromDate);
                        status = true;
                    }
                }

                if (status == true)
                {
                    Console.WriteLine("ToDateName Set value: " + ToDateName + "    to date value:  " + toDate);
                    Thread.Sleep(1000);
                    AutomationElement aeToDateName = AUIUtilities.FindElementByName(ToDateName, aeParamArea);
                    if (aeToDateName == null)
                    {
                        Console.WriteLine("ToDateName not found: " + ToDateName);
                        status = false;
                    }
                    else
                    {

                        Point pt = AUIUtilities.GetElementCenterPoint(aeToDateName);
                        Console.WriteLine("-------------------------- ToDateName found: " + ToDateName);
                        Console.WriteLine("fill value : " + toDate);
                        Input.MoveToAndDoubleClick(pt);
                        Console.WriteLine("current value aeToDateName.Current.IsKeyboardFocusable: " + aeToDateName.Current.IsKeyboardFocusable);
                        Thread.Sleep(2000);
                        //ProjBasicUI.SendTextToElement(aeToDateName, toDate);
                        //System.Windows.Forms.SendKeys.SendWait("{HOME}^{DEL}"); // home ctrl del
                        //InputSimulator.SimulateModifiedKeyStroke(VirtualKeyCode.HOME,
                        //    new[] { VirtualKeyCode.CONTROL, VirtualKeyCode.DELETE } );


                        InputSimulator.SimulateKeyPress(VirtualKeyCode.HOME);
                        InputSimulator.SimulateKeyDown(VirtualKeyCode.CONTROL);
                        InputSimulator.SimulateKeyDown(VirtualKeyCode.DELETE);
                        Console.WriteLine("deleted value :------------------ ");
                        Thread.Sleep(1000);

                        InputSimulator.SimulateKeyUp(VirtualKeyCode.DELETE);
                        InputSimulator.SimulateKeyUp(VirtualKeyCode.CONTROL);
                        InputSimulator.SimulateKeyUp(VirtualKeyCode.HOME);
                        Thread.Sleep(1000);
                        InputSimulator.SimulateTextEntry(toDate);
                        Thread.Sleep(1000);
                        // Check Field value
                        /*aeToDateName = AUIUtilities.FindElementByName(ToDateName, aeParamArea);
                        TextPattern tp = (TextPattern)aeToDateName.GetCurrentPattern(TextPattern.Pattern);
                        string v = tp.DocumentRange.GetText(-1).Trim();
                        Console.WriteLine("filled text is : " + v);*/
                        Thread.Sleep(2000);
                    }
                }

                if (status == true)
                {
                    AutomationElement aeVeiwReport = AUIUtilities.FindElementByID(ViewReportButtionId, aeParamArea);
                    if (aeVeiwReport == null)
                    {
                        Console.WriteLine("aeVeiwReport not found: " + ViewReportButtionId);
                        status = false;
                    }
                    else
                    {
                        Point pt = AUIUtilities.GetElementCenterPoint(aeVeiwReport);
                        Input.MoveTo(pt);
                        Thread.Sleep(1000);
                        Input.MoveToAndClick(pt);
                    }
                }

                if (status == true)
                {
                    aeLoadedReportWindow = ProjBasicUI.GetSelectedOverviewWindow(reportWindowName, ref errorMsg);
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

            }
            catch (Exception ex)
            {
                errorMsg = "GetLoadedReportWindow2 exception: " + ex.Message + " --- " + ex.StackTrace;
                aeLoadedReportWindow = null;
            }

            return aeLoadedReportWindow;
        }
      
        static public void GetThisReportTestData(string reportName, string currentProject, ref string fromDate, ref string toDate, ref string validateValue)
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
                case ReportName.PERFORMANCE_VEHICLES_ModeOverview:
                    validateValue = "02:01:13";
                    break;
                case ReportName.PERFORMANCE_VEHICLES_StateOverview:
                    validateValue = "01:48:39";
                    break;
                case ReportName.PERFORMANCE_TRANSPORTS_DurationDistributionBySrcDstGroup:
                    validateValue = " Thursday, November 25, 2010";
                    break;
                case ReportName.PERFORMANCE_TRANSPORTS_DurationDistributionSrcDstLocationOrStation:
                    validateValue = " Thursday, November 25, 2010";
                    break;
            }
        }

        static public bool ValidateReportPerformanceVehiclesReport2(AutomationElement aeReportWindow, string fromDate, string toDate, string validateValue, ref string errorMsg)
        {
            bool status = false;

            //AutomationElement aeReportWindow = null;
            AutomationElement aeLoadedReportWindow = null;
            Console.WriteLine("--------------  Get report window ---------------------------wait 5 second: ");
            //Thread.Sleep(20000);
            Console.WriteLine("Get loaded report window--------aeReportWindow.Current.Name: " + aeReportWindow.Current.Name);
            // for graphical view, report name is Status: graphical view but window title is System overview
            // AND had no Param area, so no reloaded windows
            if (aeReportWindow.Current.Name.IndexOf("System overview") >= 0
                || aeReportWindow.Current.Name.IndexOf("history") > 0)
            {
                status = true;
            }
            else
            {
                aeLoadedReportWindow = StatUtilities.GetLoadedReportWindow2(aeReportWindow, fromDate, toDate, ref errorMsg);
                if (aeLoadedReportWindow == null)
                {
                    Console.WriteLine(errorMsg);
                    status = false;
                }
                else
                {
                    status = true;
                    switch (aeReportWindow.Current.Name)
                    {
                        case ReportName.PERFORMANCE_VEHICLES_ModeOverview:
                        case ReportName.PERFORMANCE_VEHICLES_StateOverview:
                            status = ReportDataValidation(aeLoadedReportWindow, validateValue, ref errorMsg);
                            break;
                    }
                }
            }
            
            Console.WriteLine("ValidateReportPerformanceVehiclesReport:: status: " + status);
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
                        try
                        {
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
                        catch (System.InvalidOperationException ex)
                        {
                            Console.WriteLine("aeAllTexts.name =" + aeAllTexts[i].Current.Name);
                            //errorMsg = ex.Message.ToString() + ":-> ValuePattern";
                            Console.WriteLine("errorMsg=" + errorMsg);
                            //Thread.Sleep(9000);
                            continue;
                        }
                    }

                    if (valueFound == true)
                        break;

                    Thread.Sleep(200);
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
                ErrorMSG = "Microsoft .Net Framework " + aeError1.Current.Name;
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

        public static bool SelectFileInFileBrowserWindow(AutomationElement aeOpenFileWindow, string treeViewID, List<string> FolderList, string listViewId,
            string SelectedFilename, string actionButtonId, ref string errorMsg)
        {
            bool fileSelectedOK = true;
            /*string treeViewID = "m_TreeView";
            List<string> FolderList = new List<string>();
            FolderList.Insert(0, System.Environment.MachineName);
            FolderList.Insert(1, "C: (Local Disk )");
            FolderList.Insert(2, OSVersionInfoClass.ProgramFilesx86FolderName());
            FolderList.Insert(3, "Egemin");
            FolderList.Insert(4, "Etricc Server");*/
            //-------------------------------------------
            // aeOpenFileWindow, treeViewID,  FolderList    listViewId = "m_ListView", Filename, actionButton
            #region
            Console.WriteLine(aeOpenFileWindow.Current.Name + " is opend -------------- : " + System.DateTime.Now);
            
            AutomationElement aeTreeView = null;
            AutomationElement aeRootNode = null;
            DateTime sTime = DateTime.Now;
            // find left treeview first 
            AUIUtilities.WaitUntilElementByIDFound(aeOpenFileWindow, ref aeTreeView, treeViewID, sTime, 60);
            if (aeTreeView == null)
            {
                errorMsg = "aeTreeView not found name";
                Console.WriteLine(errorMsg);
                fileSelectedOK = false;
            }
            else
            {
                aeRootNode = null;
                TreeWalker walker = TreeWalker.ControlViewWalker;
                AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
                while (elementNode != null)
                {
                    Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
                    if (elementNode.Current.Name.ToLower().Equals(FolderList[0].ToLower()))
                    {
                        aeRootNode = elementNode;
                        Console.WriteLine("Computer node name found , it is: " + aeRootNode.Current.Name);
                        break;
                    }
                    Thread.Sleep(3000);
                    elementNode = walker.GetNextSibling(elementNode);
                }

                if (aeRootNode == null)
                {
                    errorMsg = "aeRootNode not found, ";
                    Console.WriteLine(errorMsg);
                    fileSelectedOK = false;
                }
                else
                {
                    try
                    {
                        ExpandCollapsePattern ep = (ExpandCollapsePattern)aeRootNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                        ep.Expand();
                        Thread.Sleep(9000);
                    }
                    catch (Exception ex)
                    {
                        errorMsg = "RootNode can not expaned: " + aeRootNode.Current.Name + " --- " + ex.Message;
                        Console.WriteLine(errorMsg);
                        fileSelectedOK = false;
                    }
                    //Input.MoveToAndDoubleClick(aeRootNode.GetClickablePoint());
                }
            }

            AutomationElement aeNextNode = null;
            if (fileSelectedOK)
            {
                for (int i = 1; i < FolderList.Count; i++)
                {
                    System.Windows.Forms.TreeNode treeNode = new System.Windows.Forms.TreeNode();
                    aeNextNode = TestTools.AUICommon.WalkEnabledElements(aeRootNode, treeNode, FolderList[i]);
                    if (aeNextNode == null)
                    {
                        errorMsg =  FolderList[i] + " node NOT Exist ===";
                        Console.WriteLine(errorMsg);
                        fileSelectedOK = false;
                        break;
                    }
                    else
                    {
                        Console.WriteLine(FolderList[i]+ "  node Exist ===");
                        try
                        {
                            ExpandCollapsePattern ep = (ExpandCollapsePattern)aeNextNode.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                            ep.Expand();
                            Thread.Sleep(2000);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("NextNode can not expaned: " + aeNextNode.Current.Name + " --- " + ex.Message);
                        }
                        aeRootNode = aeNextNode;
                    }
                }
            }

            // select file in right side
            if (fileSelectedOK)
            {
                //string listViewId = "m_ListView";
                AutomationElement aeListView = null;
                AutomationElement aeWCSdll = null;
                sTime = DateTime.Now;
                AUIUtilities.WaitUntilElementByIDFound(aeOpenFileWindow, ref aeListView, listViewId, sTime, 60);
                if (aeListView == null)
                {
                    errorMsg = "aeListView not found name";
                    Console.WriteLine(errorMsg);
                    fileSelectedOK = false;
                }
                else
                {
                    Console.WriteLine("List view found   .........");
                    Thread.Sleep(1000);
                    // Set a property condition that will be used to find the control.
                    System.Windows.Automation.Condition c = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.DataItem);

                    AutomationElementCollection aeAllItems = aeListView.FindAll(TreeScope.Children, c);

                    Console.WriteLine("All items count ..." + aeAllItems.Count);
                    for (int i = 0; i < aeAllItems.Count; i++)
                    {
                        Console.WriteLine("item name: " + aeAllItems[i].Current.Name);
                        if (aeAllItems[i].Current.Name.StartsWith(SelectedFilename))
                            aeWCSdll = aeAllItems[i];
                    }

                    Thread.Sleep(3000);
                    if (aeWCSdll == null)
                    {
                        errorMsg = SelectedFilename + "  not found, ";
                        Console.WriteLine(errorMsg);
                        fileSelectedOK = false;
                    }
                    else
                    {
                        Console.WriteLine(SelectedFilename + " found  ......... if off screen,then scroll down");
                        ScrollPattern scrollPattern = (ScrollPattern)aeListView.GetCurrentPattern(ScrollPattern.Pattern);
                        if (scrollPattern.Current.VerticallyScrollable)
                        {
                            while (aeWCSdll.Current.IsOffscreen)
                            {
                                scrollPattern.ScrollVertical(ScrollAmount.LargeIncrement);
                                Thread.Sleep(100);
                            }
                        }

                        #region //select dll
                        //Input.MoveTo(aeWCSdll);
                        Thread.Sleep(2000);
                        bool select = false; //Utilities.SelectItemFromList("nl", aeCombo);
                        if (fileSelectedOK)
                        {
                            SelectionPattern selectPattern =
                               aeListView.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

                            AutomationElement item
                                = AUIUtilities.FindElementByName(SelectedFilename, aeListView);
                            if (item != null)
                            {
                                Console.WriteLine(SelectedFilename + " item found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                                Thread.Sleep(2000);

                                SelectionItemPattern itemPattern = item.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                itemPattern.Select();
                                select = true;
                            }
                            else
                            {
                                errorMsg = "Finding " + SelectedFilename + "  item failed, ";
                                Console.WriteLine(errorMsg);
                                fileSelectedOK = false;
                            }

                            if (!select)
                            {
                                errorMsg = "Select " + SelectedFilename + "  item failed, ";
                                Console.WriteLine(errorMsg);
                                fileSelectedOK = false;
                            }
                        }
                        #endregion
                    }
                }
            }

            // Check selected file
            if (fileSelectedOK)
            {
                // check get value Egemin.EPIA.WCS.dll exist
                // check select button is enable
                Thread.Sleep(15000);
                AutomationElement aeSelectButton = null;
                //string BtnConnectId = "m_BtnSelect";
                DateTime mAppTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - mAppTime;
                while (aeSelectButton == null && mTime.Minutes < 5)
                {
                    Console.WriteLine("Find Application aeSelectButton : " + System.DateTime.Now);
                    aeSelectButton = AUIUtilities.FindElementByID(actionButtonId, aeOpenFileWindow);
                    Console.WriteLine("Application aeSelectButton : " + System.DateTime.Now);
                    mTime = DateTime.Now - mAppTime;
                    Thread.Sleep(2000);
                    Console.WriteLine(" find time is :" + mTime.TotalSeconds);
                }

                if (aeSelectButton == null)
                {
                    errorMsg = "aeButton not found";
                    Console.WriteLine(errorMsg);
                    fileSelectedOK = false;
                }
                else
                {
                    Thread.Sleep(500);
                    TestTools.AUIUtilities.ClickElement(aeSelectButton);
                    Thread.Sleep(2000);
                }

            }
            #endregion
            //-------------------------------------------
            return fileSelectedOK;
        }

        public static bool CheckParserConfiguratorOutput(string mainformId, ref string errorMsg)
        {
            bool outputOK = true;
            #region //check output
            AutomationElement aeDocument = null;
            AutomationElement aeWindow = ProjBasicUI.GetMainWindowWithinTime("MainForm", 120);
            if (aeWindow == null)
            {
                errorMsg = "MainForm not found";
                Console.WriteLine(errorMsg);
                outputOK = false;
            }
            else
            {
                string textBoxPanelId = "m_TextBoxContainerPanel";
                AutomationElement aeTextBoxPanel = null;
                DateTime sTime = DateTime.Now;
                AUIUtilities.WaitUntilElementByIDFound(aeWindow, ref aeTextBoxPanel, textBoxPanelId, sTime, 60);
                if (aeTextBoxPanel == null)
                {
                    errorMsg = "aeTextBoxPanel not found name";
                    Console.WriteLine(errorMsg);
                    outputOK = false;
                }
                else
                {
                    AutomationElementCollection aeAllDocuments = null;
                    // find ducument text
                    System.Windows.Automation.Condition cDocs = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Document);

                    int k = 0;
                    DateTime mStartTime = DateTime.Now;
                    TimeSpan mTime = DateTime.Now - mStartTime;
                    while (aeDocument == null && mTime.TotalSeconds <= 300)
                    {
                        Console.WriteLine("run search=" + k);
                        k++;
                        aeAllDocuments = aeTextBoxPanel.FindAll(TreeScope.Children, cDocs);
                        Thread.Sleep(3000);
                        for (int i = 0; i < aeAllDocuments.Count; i++)
                        {
                            Console.WriteLine("aeDocument[" + i + "] name length=" + aeAllDocuments[i].Current.Name.Length);
                            if (aeAllDocuments[i].Current.Name.Length > 20)
                            {
                                aeDocument = aeAllDocuments[i];
                                //Console.WriteLine("--------------aeDocument[" + i + "]IsEnabled=" + aeDocument.Current.Name);
                                if (aeDocument.Current.Name.IndexOf("successfully") > 0)
                                {
                                    //Console.WriteLine("aeDocument[" + i + "]=" + aeDocument.Current.Name);
                                    break;
                                }
                                else if (aeDocument.Current.Name.IndexOf("error") > 0)
                                {
                                    errorMsg = "aeDocument not found or text not finished successfully";
                                    Console.WriteLine(errorMsg);
                                    outputOK = false;
                                    break;
                                }
                                else
                                {
                                    Console.WriteLine("=== NOT FINISHED =========:");
                                    aeDocument = null;
                                }
                            }
                            else
                            {
                                Console.WriteLine("aeDocument[" + i + "] name=" + aeAllDocuments[i].Current.Name);
                            }
                        }
                        mTime = DateTime.Now - mStartTime;
                    }

                    if (aeDocument == null)
                    {
                        errorMsg = "aeDocument not found or text not finished successfully";
                        Console.WriteLine(errorMsg);
                        outputOK = false;
                    }
                }
            }
            #endregion
            return outputOK;
        }
    }
}
