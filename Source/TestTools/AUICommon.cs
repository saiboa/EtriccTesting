using System;
using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;

namespace TestTools
{
    /// <summary>
    ///  Common AUI operation for All EPIA3 projects
    /// </summary>
    public class AUICommon
    {
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        /// <summary>
        ///  Clear Overview screens etc to avoid to many UI controls that cause the looking
        ///  element time last too long
        /// </summary>
        /// <param name="root"></param>
        public static void ClearDisplayedScreens(AutomationElement root)
        {
            AutomationElement aeTab = AUIUtilities.FindElementByType(ControlType.Tab, root);
            if (aeTab != null)
            {
                double right = aeTab.Current.BoundingRectangle.Right;
                double bottom = aeTab.Current.BoundingRectangle.Bottom;
                double top = aeTab.Current.BoundingRectangle.Top;

                double x = right - 5;
                double y = (top + bottom)/2;
                var p = new Point(x, y);

                for (int i = 0; i < 10; i++)
                {
                    Input.MoveToAndClick(p);
                }
                Thread.Sleep(1000);
            }
        }


        public static void ClearDisplayedScreens(AutomationElement root, int nrClears)
        {
            AutomationElement aeTab = AUIUtilities.FindElementByType(ControlType.Tab, root);
            if (aeTab != null)
            {
                double right = aeTab.Current.BoundingRectangle.Right;
                double bottom = aeTab.Current.BoundingRectangle.Bottom;
                double top = aeTab.Current.BoundingRectangle.Top;

                double x = right - 5;
                double y = (top + bottom)/2;
                var p = new Point(x, y);

                for (int i = 0; i < nrClears; i++)
                {
                    Input.MoveToAndClick(p);
                }
                Thread.Sleep(1000);
            }
        }

        public static void ErrorWindowHandling(AutomationElement element, ref string errorMsg)
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
            errorMsg = aeError.Current.Name;
            Console.WriteLine("Error Msg is ------------:" + errorMsg);

            AutomationElement aeClose = AUIUtilities.FindElementByID(close, element);
            if (aeClose == null)
            {
                error = "Close button element not Found";
                Console.WriteLine(error);
                return;
            }
            Console.WriteLine("aeClose is found ------------:");

            Thread.Sleep(1000);
            var ivp = (InvokePattern) aeClose.GetCurrentPattern(InvokePattern.Pattern);
            ivp.Invoke();
        }

        public static Point GetDataGridViewCellPointAt(int row, string colName, AutomationElement dgvElement)
        {
            if (null == dgvElement)
            {
                throw new ArgumentNullException("Grid Null");
            }

            Console.WriteLine("try to Find DataGridView Cell::" + dgvElement.Current.Name);

            // Construct the Grid Cell Element Name
            string name = colName + " Row " + row.ToString(CultureInfo.InvariantCulture);
            // Get the Element with the Row Col Coordinates
            //UIWindow element = FindDescendantByName(name, ControlType.Custom);
            AutomationElement element = AUIUtilities.FindElementByName(name, dgvElement);

            if (element == null)
            {
                Console.WriteLine("Find element failed:" + "element");
            }
            else
                Console.WriteLine("Grid Cell element found at time: " + DateTime.Now.ToString("HH:mm:ss"));

            Point point = AUIUtilities.GetElementCenterPoint(element);

            return point;
        }

        public static AutomationElement GetDataGridViewCellElementAt(int row, string colName,
                                                                     AutomationElement dgvElement)
        {
            if (null == dgvElement)
            {
                throw new ArgumentNullException("Grid Null");
            }

            Console.WriteLine("try to Find DataGridView Cell:element:" + dgvElement.Current.Name);

            // Construct the Grid Cell Element Name
            string name = colName + " Row " + row.ToString();
            // Get the Element with the Row Col Coordinates
            //UIWindow element = FindDescendantByName(name, ControlType.Custom);
            AutomationElement element = AUIUtilities.FindElementByName(name, dgvElement);

            if (element == null)
            {
                Console.WriteLine("Find element failed:" + "element");
            }
            else
                Console.WriteLine("Grid Cell element found at time: " + DateTime.Now.ToString("HH:mm:ss"));

            return element;
        }

        public static string GetDataGridViewCellValueAt(int row, string colName, AutomationElement dgvElement)
        {
            if (null == dgvElement)
            {
                throw new ArgumentNullException("Grid Null");
            }

            Console.WriteLine("try to Find DataGridView Cell:value:" + dgvElement.Current.Name);

            // Construct the Grid Cell Element Name
            string name = colName + " Row " + row.ToString(CultureInfo.InvariantCulture);
            // Get the Element with the Row Col Coordinates
            //UIWindow element = FindDescendantByName(name, ControlType.Custom);
            AutomationElement element = AUIUtilities.FindElementByName(name, dgvElement);

            if (element == null)
            {
                Console.WriteLine("Find element failed:" + "element");
            }
            else
                Console.WriteLine("Button element found at time: " + DateTime.Now.ToString("HH:mm:ss"));

            Point Point = AUIUtilities.GetElementCenterPoint(element);
            Input.MoveTo(Point);
            Thread.Sleep(1000);

            string result = null;
            try
            {
                var vp = element.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                result = vp.Current.Value;
                Console.WriteLine("Get element.Current Value:" + vp.Current.Value);
            }
            catch (NullReferenceException)
            {
                result = null;
            }

            return result;
        }

        /// <summary>
        /// Find TreeView Node Element of the specific menu screen
        /// </summary>
        /// <param name="testcase"></param>
        /// <param name="root"></param>
        /// <param name="category"></param>
        /// <param name="treeviewNode"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static AutomationElement FindTreeViewNodeLevel1(string testcase, AutomationElement root, string category,
                                                               string treeviewNode, ref string message)
        {
            message = "try to find :" + treeviewNode;
            AutomationElement aeNodeLink = null;
            bool findOk = true;
            Console.WriteLine(testcase + "--> start to find : " + category + "- " + treeviewNode);
            if (root == null)
            {
                message = testcase + " --> shell main form : root param is null";
                Console.WriteLine(message);
                findOk = false;
            }

            AutomationElement aeMenuArea = null;
            if (findOk)
            {
                Condition cArea = new AndCondition(
                    new PropertyCondition(AutomationElement.AutomationIdProperty, "windowDockingArea2"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                    );

                aeMenuArea = root.FindFirst(TreeScope.Element | TreeScope.Children, cArea);
                if (aeMenuArea == null)
                {
                    message = testcase + ":failed find menu list area at time: " + DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(message);
                    findOk = false;
                }
            }

            if (findOk)
            {
                Condition c1 = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "stackStrip1"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar)
                    );
                AutomationElement aeToolBar = aeMenuArea.FindFirst(TreeScope.Element | TreeScope.Descendants, c1);
                if (aeToolBar == null)
                {
                    message = testcase + ": aeToolBar not Found";
                    Console.WriteLine(message);
                    findOk = false;
                }
                else
                {
                    //Input.MoveTo(aeToolBar);
                    Console.WriteLine(testcase + " -->  Category ToolBar found ");
                    Thread.Sleep(500);
                    // Find group Button Element
                    Condition c2New = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, category),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox)
                       );

                    Condition c2Old = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, category),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                        );
                    AutomationElement aeCategory = aeToolBar.FindFirst(TreeScope.Element | TreeScope.Descendants, c2New);
                    if (aeCategory == null)
                    {
                        message = testcase + ":failed find " + category + " New at time: " +
                                  DateTime.Now.ToString("HH:mm:ss") + " , try to fild C2Old ";
                        Console.WriteLine(message);
                        aeCategory = aeToolBar.FindFirst(TreeScope.Element | TreeScope.Descendants, c2Old);
                    }
                    if (aeCategory == null)
                    {
                        message = testcase + ":failed find " + category + " at time: " +
                                  DateTime.Now.ToString("HH:mm:ss");
                        Console.WriteLine(message);
                        findOk = false;
                    }
                    else
                    {
                        Input.MoveTo(aeCategory);
                        Thread.Sleep(100);
                        Console.WriteLine(testcase + " --> " + category + ": found  ");
                        var ipX = (InvokePattern) aeCategory.GetCurrentPattern(InvokePattern.Pattern);
                        ipX.Invoke();
                        Thread.Sleep(500);
                    }
                }
            }

            const string id = "m_TreeView";
            AutomationElement aeTreeView = null;
            if (findOk)
            {
                AUIUtilities.WaitUntilElementByIDFound(aeMenuArea, ref aeTreeView, id, DateTime.Now, 60);
                if (aeTreeView == null)
                {
                    message = "aeTreeView not found name";
                    Console.WriteLine(message);
                    findOk = false;
                }
                else
                {
                    TreeWalker walker = TreeWalker.ControlViewWalker;
                    AutomationElement elementNode = walker.GetFirstChild(aeTreeView);
                    while (elementNode != null)
                    {
                        Console.WriteLine("aeTreeView node name is: -->  " + elementNode.Current.Name +
                                          " <searching node>:" + treeviewNode);
                        if (elementNode.Current.Name.Trim().Equals(treeviewNode))
                        {
                            //Input.MoveTo(elementNode);
                            aeNodeLink = elementNode;
                            message = string.Empty;
                            //Console.WriteLine(treeviewNode + " node found : " + aeNodeLink.Current.Name);
                            break;
                        }
                        elementNode = walker.GetNextSibling(elementNode);
                    }
                }
            }
            return aeNodeLink;
        }

        /// <summary>
        /// Find TreeView Node Element of the specific menu screen
        /// </summary>
        /// <param name="testcase"></param>
        /// <param name="root"></param>
        /// <param name="category"></param>
        /// <param name="treeviewNode"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static AutomationElement FindTreeViewNodeLevelAll(string testcase, AutomationElement root,
                                                                 string category,
                                                                 string treeviewNode, ref string message)
        {
            Console.WriteLine(" start FindTreeViewNodeLevelAll root.Current.Name : " + root.Current.Name);
            bool findNodeOK = true;
            AutomationElement aeNodeLink = null;

            if (root == null)
            {
                message = "FindTreeViewNodeLevelAll: main form is null";
                findNodeOK = false;
            }

            Console.WriteLine(testcase + " start to find : " + category + "- " + treeviewNode);
            AutomationElement aeToolBar = null;
            if (findNodeOK)
            {
                // Find ToolBar
                Condition c1 = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, "stackStrip1"),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar)
                    );

                // wait until aeToolBar found
                DateTime startTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - startTime;
                while (aeToolBar == null && mTime.TotalSeconds < 120)
                {
                    try
                    {
                        aeToolBar = root.FindFirst(TreeScope.Element | TreeScope.Subtree, c1);
                        Thread.Sleep(2000);
                        Console.WriteLine("stackStrip1 ToolBar  NOT FOUND : " + category);
                        mTime = DateTime.Now - startTime;
                    }
                    catch (Exception ex)
                    {
                        aeToolBar = null;
                        message = testcase + "FindTreeViewNodeLevelAll: find aeToolBar exception:" + ex.Message;
                        Console.WriteLine(testcase + "FindTreeViewNodeLevelAll: find aeToolBar exception:" + ex.Message);
                        Thread.Sleep(2000);
                        MessageBox.Show(message, "Exception", MessageBoxButtons.OK);
                    }
                }

                if (aeToolBar == null)
                {
                    message = testcase + "FindTreeViewNodeLevelAll: aeToolBar not Found";
                    Console.WriteLine(message);
                    findNodeOK = false;
                }
            }

            AutomationElement aeCategory = null;
            if (findNodeOK)
            {
                // Find group Button Element
                Condition c2Button = new AndCondition(
                 new PropertyCondition(AutomationElement.NameProperty, category),
                 new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                 );

                // new version chenged button to checkbox 
                Condition c2CheckBox = new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, category),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox)
                    );

              
                Thread.Sleep(500);
                // wait until aeToolBar found
                DateTime startTime = DateTime.Now;
                TimeSpan mTime = DateTime.Now - startTime;
                while (aeCategory == null && mTime.TotalSeconds < 120)
                {
                    try
                    {
                        aeCategory = aeToolBar.FindFirst(TreeScope.Element | TreeScope.Descendants, c2CheckBox);
                        if (aeCategory == null)
                        {
                            aeCategory = aeToolBar.FindFirst(TreeScope.Element | TreeScope.Descendants, c2Button);
                        }
                        Thread.Sleep(5000);
                        Console.WriteLine("category NOT FOUND : " + category);
                        mTime = DateTime.Now - startTime;
                    }
                    catch (Exception ex)
                    {
                        aeCategory = null;
                        message = testcase + "FindTreeViewNodeLevelAll: find " + category + " exception:" + ex.Message;
                        Console.WriteLine(testcase + "FindTreeViewNodeLevelAll: find " + category + " exception:" + ex.Message);
                        MessageBox.Show(message, "Exception", MessageBoxButtons.OK);
                    }
                }

                if (aeCategory == null)
                {
                    message = testcase + ":failed find " + category + " at time: " + DateTime.Now.ToString("HH:mm:ss");
                    Console.WriteLine(message);
                    findNodeOK = false;
                }
                else
                {
                    var ipX = (InvokePattern) aeCategory.GetCurrentPattern(InvokePattern.Pattern);
                    ipX.Invoke();
                }
            }

            string id = "m_TreeView";
            AutomationElement aeTreeView = null;
            DateTime sTime = DateTime.Now;
            if (findNodeOK)
            {
                AUIUtilities.WaitUntilElementByIDFound(root, ref aeTreeView, id, sTime, 60);
                if (aeTreeView == null)
                {
                    message = "aeTreeView not found";
                    Console.WriteLine(message);
                    findNodeOK = false;
                }
                else
                {
                    var treeNode = new TreeNode();
                    // treeviewNode is nodeNme like Overview
                    aeNodeLink = WalkEnabledElements(aeTreeView, treeNode, treeviewNode);
                    if (aeNodeLink != null)
                        Console.WriteLine("found aeNodeLink node name is: " + aeNodeLink.Current.Name);
                    else
                        message = treeviewNode + " not found";
                }
            }

            return aeNodeLink;
        }

        /// <summary>
        /// Walks the UI Automation tree and adds the control type of each enabled control 
        /// element it finds to a TreeView.
        /// </summary>
        /// <param name="rootElement">The root of the search on this iteration.</param>
        /// <param name="treeNode">The node in the TreeView for this iteration.</param>
        /// <remarks>
        /// This is a recursive function that maps out the structure of the subtree beginning at the
        /// UI Automation element passed in as rootElement on the first call. This could be, for example,
        /// an application window.
        /// CAUTION: Do not pass in AutomationElement.RootElement. Attempting to map out the entire subtree of
        /// the desktop could take a very long time and even lead to a stack overflow.
        /// </remarks>
        public static AutomationElement WalkEnabledElements(AutomationElement rootElement, TreeNode treeNode,
                                                            string nodeName)
        {
            AutomationElement ele = null;
            Condition condition1 = new PropertyCondition(AutomationElement.IsControlElementProperty, true);
            Condition condition2 = new PropertyCondition(AutomationElement.IsEnabledProperty, true);
            var walker = new TreeWalker(new AndCondition(condition1, condition2));
            AutomationElement elementNode = walker.GetFirstChild(rootElement);
            while (elementNode != null)
            {
                Console.WriteLine("aeTreeView node name is: " + elementNode.Current.Name);
                if (elementNode.Current.Name.Equals(nodeName))
                {
                    ele = elementNode;
                    break;
                }
                TreeNode childTreeNode = treeNode.Nodes.Add(elementNode.Current.ControlType.LocalizedControlType);
                ele = WalkEnabledElements(elementNode, childTreeNode, nodeName);
                elementNode = walker.GetNextSibling(elementNode);
            }

            return ele;
        }

        public static AutomationElement WalkTreeViewFirstLevelNode(AutomationElement rootElement, string nodeName)
        {
            Console.WriteLine("WalkTreeViewFirstLevelNode: " + rootElement.Current.Name);
            AutomationElement ele = null;
            Condition condition1 = new PropertyCondition(AutomationElement.IsControlElementProperty, true);
            Condition condition2 = new PropertyCondition(AutomationElement.IsEnabledProperty, true);
            var walker = new TreeWalker(new AndCondition(condition1, condition2));
            AutomationElement elementNode = walker.GetFirstChild(rootElement);
            while (elementNode != null)
            {
                Console.WriteLine("level1 node name is: " + elementNode.Current.Name);
                if (elementNode.Current.Name.Equals(nodeName))
                {
                    ele = elementNode;
                    break;
                }
                elementNode = walker.GetNextSibling(elementNode);
            }

            return ele;
        }
    }
}