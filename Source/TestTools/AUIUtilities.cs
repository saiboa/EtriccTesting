using System;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;

namespace TestTools
{
    public class AUIUtilities
    {
        /// <summary>
        /// Finds a UI Automation child element by AutomationID.
        /// </summary>
        /// <param name="automationId">AutomationID of the control, such as "button1".</param>
        /// <param name="rootElement">Parent element, such as an application window, or
        /// AutomationElement.RootElement object when searching for the application window.</param>
        /// <returns>The UI Automation element.</returns>
        public static AutomationElement FindElementByID(String automationId, AutomationElement rootElement)
        {
            if ((automationId == "") || (rootElement == null))
            {
                throw new ArgumentException("Argument cannot be null or empty.");
            }
            // Set a property condition that will be used to find the control.
            Condition c = new PropertyCondition(
                AutomationElement.AutomationIdProperty, automationId, PropertyConditionFlags.IgnoreCase);

            // Find the element.
            return rootElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
        }

        //Example 2. This example shows how to find a control by control name.
        /// <summary>
        /// Finds a UI Automation child element by name.
        /// </summary>
        /// <param name="controlName">Name of the control, such as "button1".</param>
        /// <param name="rootElement">Parent element, such as an application window, or
        /// AutomationElement.RootElement object when searching for the application window.</param>
        /// <returns>The UI Automation element.</returns>
        public static AutomationElement FindElementByName(String controlName, AutomationElement rootElement)
        {
            if ((controlName == "") || (rootElement == null))
            {
                throw new ArgumentException("Argument cannot be null or empty.");
            }
            // Set a property condition that will be used to find the control.
            Condition c = new PropertyCondition(
                AutomationElement.NameProperty, controlName,
                PropertyConditionFlags.IgnoreCase);
            // Find the element.
            return rootElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
        }

        //Example 3. This example shows how to find a control by control type.
        /// <summary>
        /// Finds a UI Automation child element by control type.
        /// </summary>
        /// <param name="controlType">Control type of the control, such as Button.</param>
        /// <param name="rootElement">Parent element, such as an application window, or
        /// AutomationElement.RootElement when searching for the application window.</param>
        /// <returns>The UI Automation element.</returns>
        public static AutomationElement FindElementByType(ControlType controlType, AutomationElement rootElement)
        {
            if ((controlType == null) || (rootElement == null))
            {
                throw new ArgumentException("Argument cannot be null.");
            }

            // Set a property condition that will be used to find the control.
            Condition c = new PropertyCondition(
                AutomationElement.ControlTypeProperty, controlType);

            // Find the element.
            return rootElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
        }

        //Example 4. This example shows how to find a control based on a control condition, such as all buttons that are enabled.
        /// <summary>
        /// Finds all enabled buttons in the specified root element.
        /// </summary>
        /// <param name=quot;rootElement">The parent element.</param>
        /// <returns>A collection of elements that meet the conditions.</returns>
        public static AutomationElementCollection FindByMultipleConditions(AutomationElement rootElement)
        {
            if (rootElement == null)
            {
                throw new ArgumentException();
            }

            Condition c = new AndCondition(
                new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                );

            // Find all children that match the specified conditions.
            return rootElement.FindAll(TreeScope.Children, c);
        }

        //Example 6. This example shows how to find an element from a list item. 
        //The example uses the FindAll method to retrieve a specified item from a list. 
        //This is faster for WPF controls than using the TreeWalker class.
        /// <summary>
        /// Retrieves an element in a list by using the FindAll method.
        /// </summary>
        /// <param name="parent">The list element.</param>
        /// <param name="index"> The index of the element to find.</param>
        /// <returns>The list item.</returns>
        public static AutomationElement FindListItemByIndex(AutomationElement parent, int index)
        {
            if (parent == null)
            {
                throw new ArgumentException();
            }

            Condition c = new AndCondition(
                //new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                new PropertyCondition(AutomationElement.IsControlElementProperty, true));

            // Find all children that match the specified conditions.
            AutomationElementCollection found = parent.FindAll(TreeScope.Descendants, c);
            return found[index];
        }

        //Example 7. This example shows how to find a UI element by class name.
        /// <summary>
        /// Finds an element by its class name starting from a specific root element.
        /// </summary>
        /// <param name="root">The root element to start from.</param>
        /// <param name="type">The class name of the control type to find.</param>
        /// <returns>The list item.</returns>
        public static AutomationElement FindElementByClassName(AutomationElement root, String type)
        {
            if ((root == null) || (type == ""))
            {
                throw new ArgumentException("Argument cannot be null or empty.");
            }

            Condition c = new AndCondition(
                new PropertyCondition(AutomationElement.ClassNameProperty, type));

            // Find all children that match the specified conditions.
            AutomationElementCollection found = root.FindAll(TreeScope.Children, c);
            return found[0];
        }

        public static bool FindElementAndToggle(String automationId, AutomationElement rootAE, ToggleState state)
        {
            Console.WriteLine("FindElementAndToggle: " + automationId);
            AutomationElement ae = FindElementByID(automationId, rootAE);
            if (ae != null)
            {
                Thread.Sleep(500);
                Console.WriteLine(automationId + " Element found");
                var tg = ae.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                Thread.Sleep(500);
                ToggleState tgTState = tg.Current.ToggleState;
                Console.WriteLine("FindElementAndToggle to: " + tgTState.ToString());
                Thread.Sleep(1000);
                switch (state)
                {
                    case ToggleState.Off:
                        if (tgTState == ToggleState.On)
                        {
                            tg.Toggle();
                        }
                        break;
                    case ToggleState.On:
                        if (tgTState == ToggleState.Off)
                        {
                            tg.Toggle();
                        }
                        break;
                }
            }
            else
            {
                Console.WriteLine("FindElementAndToggle not found:" + automationId);
                return false;
            }

            return true;
        }

        public static bool FindElementAndClick(String automationId, AutomationElement rootAE)
        {
            Console.WriteLine("FindElementAndClick: " + automationId);
            AutomationElement ae = FindElementByID(automationId, rootAE);
            if (ae != null)
            {
                Thread.Sleep(500);
                Point pnt = GetElementCenterPoint(ae);
                Input.MoveTo(pnt);
                Thread.Sleep(1000);
                var ip = (InvokePattern) ae.GetCurrentPattern(InvokePattern.Pattern);
                ip.Invoke();
                return true;
            }
            else
                return false;
        }

        public static bool FindElementAndClickPoint(String automationID, AutomationElement rootAE)
        {
            Console.WriteLine("FindElementAndClick: " + automationID);
            AutomationElement ae = FindElementByID(automationID, rootAE);
            if (ae != null)
            {
                Thread.Sleep(500);
                Point pnt = GetElementCenterPoint(ae);
                Input.MoveTo(pnt);
                Thread.Sleep(1000);
                Input.ClickAtPoint(pnt);
                return true;
            }
            else
                return false;
        }

        public static bool FindTextBoxAndChangeValue(String automationID, AutomationElement rootAE,
                                                     out string getValue, string setValue, ref string msg)
        {
            getValue = string.Empty;
            Console.WriteLine("FindTextBoxAndChangeValue: " + automationID);
            try
            {
                AutomationElement aeTextBox = FindElementByID(automationID, rootAE);
                if (aeTextBox != null)
                {
                    //Thread.Sleep(500);
                    Point pnt = GetElementCenterPoint(aeTextBox);
                    Input.MoveTo(pnt);
                    Thread.Sleep(500);
                    var vp = (ValuePattern) aeTextBox.GetCurrentPattern(ValuePattern.Pattern);
                    //Thread.Sleep(1000);
                    getValue = vp.Current.Value;
                    Thread.Sleep(500);
                    vp.SetValue(setValue);
                    Thread.Sleep(500);
                    return true;
                }
                msg = automationID + " not found";
                return false;
            }
            catch (Exception ex)
            {
                msg = ex.Message + ":" + ex.StackTrace;
                ;
                return false;
            }
        }

        public static bool FindDocumentAndSendText(String automationId, AutomationElement rootAE,
                                                   string setValue, ref string msg)
        {
            Console.WriteLine("FindTextBoxAndChangeValue: " + automationId);
            try
            {
                AutomationElement aeTextBox = FindElementByID(automationId, rootAE);
                if (aeTextBox != null)
                {
                    //Thread.Sleep(500);
                    Point pnt = GetElementCenterPoint(aeTextBox);
                    Input.MoveTo(pnt);
                    Thread.Sleep(500);

                    aeTextBox.SetFocus();
                    Thread.Sleep(500);
                    SendKeys.SendWait(setValue);
                    Thread.Sleep(1000);
                    // Check Field value
                    var tp = (TextPattern) aeTextBox.GetCurrentPattern(TextPattern.Pattern);
                    //Thread.Sleep(1000);
                    string v = tp.DocumentRange.GetText(-1).Trim();
                    Console.WriteLine("filled text is : " + v);
                    if (v.Equals(setValue))
                    {
                        return true;
                    }
                    else
                    {
                        msg = "input value  not correct" + v;
                        return false;
                    }
                }
                msg = automationId + " not found";
                return false;
            }
            catch (Exception ex)
            {
                msg = ex.Message + ":" + ex.StackTrace;
                return false;
            }
        }

        public static Point GetElementCenterPoint(AutomationElement ae)
        {
            var Point = new Point();
            //Double Bottom = ae.Current.BoundingRectangle.Bottom;
            Double left = ae.Current.BoundingRectangle.Left;
            //Double Right = ae.Current.BoundingRectangle.Right;
            Double width = ae.Current.BoundingRectangle.Width;
            Double height = ae.Current.BoundingRectangle.Height;
            //Double X = ae.Current.BoundingRectangle.X;
            //Double Y = ae.Current.BoundingRectangle.Y;
            Double top = ae.Current.BoundingRectangle.Top;

            //Console.WriteLine("Point Bottom: " + Bottom);
            //Console.WriteLine("Point Left: " + Left);
            //Console.WriteLine("Point Right: " + Right);
            Console.WriteLine("AE ( Width:" + width + ", Height:" + height + ", LeftTop(" + left + "," + top + ")" +
                              " ): aeName=" + ae.Current.Name + " <id>:" + ae.Current.AutomationId);
            //Console.WriteLine("Point X: " + X);
            //Console.WriteLine("Point Y: " + Y);
            //Console.WriteLine("Point Top: " + Top);

            Point.Y = top + (height/2);
            Point.X = left + (width/2);
            return Point;
        }

        public static void WaitUntilElementByIDFound(AutomationElement root, ref AutomationElement element, string id,
                                                     DateTime startTime, int duration)
        {
            TimeSpan mTime = DateTime.Now - startTime;
            int wtsec = 0;
            while (element == null && wtsec < duration)
            {
                Console.WriteLine("Try to Find " + id + " at : " + DateTime.Now);
                element = FindElementByID(id, root);
                Thread.Sleep(5000);
                mTime = DateTime.Now - startTime;
                wtsec = wtsec + 5;
            }

            if (element == null)
                Console.WriteLine("after " + duration + " seconds" + id + " is not found time is (sec) :" +
                                  mTime.Milliseconds);
            else
                Console.WriteLine(id + " found time is (sec) :" + mTime.TotalSeconds);
        }

        public static AutomationElement FindTreeViewNodeByName(string testcase, AutomationElement treeviewRoot,
                                                               string treeviewNode, ref string message)
        {
            Console.WriteLine(testcase + " start to find : " + treeviewNode);

            AutomationElement aeNodeLink = null;
            TreeWalker walker = TreeWalker.ControlViewWalker;
            AutomationElement elementNode = walker.GetFirstChild(treeviewRoot);
            if (elementNode != null)
                Console.WriteLine(" first child name: " + elementNode.Current.Name);

            while (elementNode != null)
            {
                if (elementNode.Current.Name.EndsWith(treeviewNode))
                {
                    Console.WriteLine(" it is equal: ");
                    aeNodeLink = elementNode;
                    break;
                }
                elementNode = walker.GetNextSibling(elementNode);
            }

            Console.WriteLine(" final: ");
            return aeNodeLink;
        }

        public static AutomationElement WindowPanelFinder(AutomationElement root, string category, string panelName,
                                                          ref string message)
        {
            AutomationElement aePanelLink;
            Console.WriteLine(panelName + " start find time: " + DateTime.Now.ToString("HH:mm:ss"));

            // Find ToolBar
            Condition c1 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "stackStrip1"),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar)
                );
            AutomationElement aeToolBar
                = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c1);

            if (aeToolBar == null)
            {
                message = panelName + ": aeToolBar not Found";
                Console.WriteLine(message);
                return null;
            }
            Input.MoveTo(aeToolBar);
            Console.WriteLine(panelName + " ToolBar found at time1: " + DateTime.Now.ToString("HH:mm:ss"));

            Thread.Sleep(2000);

            // Find group Button Element
            Condition c2 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, category),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                );

            AutomationElement aeCategory = aeToolBar.FindFirst(TreeScope.Element | TreeScope.Descendants, c2);

            if (aeCategory == null)
            {
                message = panelName + ":failed find " + category + " at time: " + DateTime.Now.ToString("HH:mm:ss");
                Console.WriteLine(message);
                return null;
            }
            Input.MoveTo(aeCategory);
            Console.WriteLine(panelName + ":" + category + ": found at time: " + DateTime.Now.ToString("HH:mm:ss"));

            Thread.Sleep(2000);
            var ipX = (InvokePattern) aeCategory.GetCurrentPattern(InvokePattern.Pattern);
            ipX.Invoke();

            Thread.Sleep(1000);

            // Find panelName HyperLink Element
            Condition c3 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, panelName),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Hyperlink)
                );

            aePanelLink = root.FindFirst(TreeScope.Element | TreeScope.Descendants, c3);
            if (aePanelLink == null)
            {
                message = panelName + ":failed to find " + panelName + " at time3: " + DateTime.Now.ToString("HH:mm:ss");
                Console.WriteLine(message);
                return null;
            }
            Input.MoveTo(aePanelLink);
            Console.WriteLine(panelName + " found at time: " + DateTime.Now.ToString("HH:mm:ss"));

            Thread.Sleep(2000);

            var ip3 = (InvokePattern) aePanelLink.GetCurrentPattern(InvokePattern.Pattern);
            ip3.Invoke();

            return aePanelLink;
        }

        public static ToggleState FindCheckBoxAndToggleState(String automationId, AutomationElement rootAE,
                                                             ref string msg)
        {
            Console.WriteLine("FindCheckBoxAndToggleState: " + automationId);
            AutomationElement ae = FindElementByID(automationId, rootAE);
            if (ae != null)
            {
                GetElementCenterPoint(ae);
                Thread.Sleep(500);
                var tg = ae.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                return tg.Current.ToggleState;
            }
            else
            {
                msg = automationId + " not found";
                return ToggleState.Indeterminate;
            }
        }

        public static string FindTextBoxAndValue(String automationID, AutomationElement rootAE, ref string msg)
        {
            Console.WriteLine("FindTextBoxAndValue: " + automationID);
            string textValue = null;
            try
            {
                AutomationElement aeTextBox = FindElementByID(automationID, rootAE);
                if (aeTextBox != null)
                {
                    Thread.Sleep(100);
                    Input.MoveTo(aeTextBox);
                    Thread.Sleep(500);
                    var vp = (ValuePattern) aeTextBox.GetCurrentPattern(ValuePattern.Pattern);
                    //Thread.Sleep(1000);
                    textValue = vp.Current.Value;
                    Console.WriteLine("TextValue: " + textValue);
                }
                else
                {
                    msg = automationID + "not found";
                }
            }
            catch (Exception ex)
            {
                msg = automationID + " exception " + ex.Message;
            }
            return textValue;
        }

        public static void TreeViewNodeExpandCollapseState(AutomationElement node, ExpandCollapseState state)
        {
            Console.WriteLine("TreeViewNodeExpandCollapseState: " + node.Current.Name);
            var ecs = node.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
            ExpandCollapseState ecsState = ecs.Current.ExpandCollapseState;
            Console.WriteLine("Current node state is: " + ecsState.ToString());
            Thread.Sleep(1000);
            switch (state)
            {
                case ExpandCollapseState.Collapsed:
                    if (ecsState != ExpandCollapseState.Collapsed)
                    {
                        ecs.Collapse();
                    }
                    break;
                case ExpandCollapseState.Expanded:
                    if (ecsState != ExpandCollapseState.Expanded)
                    {
                        ecs.Expand();
                        ;
                    }
                    break;
            }
        }

        public static void MoveUIElement(AutomationElement UiElement, double x, double y)
        {
            try
            {
                // find current location of our window
                Point targetLocation = UiElement.Current.BoundingRectangle.Location;

                // Obtain required control patterns from our automation element
                var windowPattern = UiElement.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;

                if (windowPattern == null) return;

                // Make sure our window is usable.
                // WaitForInputIdle will return before the specified time 
                // if the window is ready.
                if (false == windowPattern.WaitForInputIdle(10000))
                {
                    Console.WriteLine("Object not responding in a timely manner.");
                    return;
                }
                Console.WriteLine("Window ready for user interaction");

                // Register for required events
                //RegisterForEvents(targetWindow, WindowPattern.Pattern, TreeScope.Element);

                // Obtain required control patterns from our automation element
                var transformPattern = UiElement.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;


                if (transformPattern == null) return;

                // Is the TransformPattern object moveable?
                if (transformPattern.Current.CanMove)
                {
                    // Enable our WindowMove fields
                    //xCoordinate.IsEnabled = true;
                    //yCoordinate.IsEnabled = true;
                    //moveTarget.IsEnabled = true;

                    // Move element
                    transformPattern.Move(x, y);
                }
                else
                {
                    Console.WriteLine("Wndow is not moveable.");
                }
            }
            catch (ElementNotAvailableException)
            {
                Console.WriteLine("Client window no longer available.");
            }
            catch (InvalidOperationException e1)
            {
                Console.WriteLine("Client window cannot be moved." + e1.Message + "----" + e1.StackTrace);
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.ToString());
            }
        }

        public static void ClickElement(AutomationElement element)
        {
            if (element != null)
            {
                if (element.Current.ControlType.Equals(ControlType.Button))
                {
                    var pattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    pattern.Invoke();
                    Thread.Sleep(2000);
                }
                else if (element.Current.ControlType.Equals(ControlType.RadioButton))
                {
                    var pattern = element.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                    pattern.Select();
                    Thread.Sleep(2000);
                }
                else if (element.Current.ControlType.Equals(ControlType.CheckBox))
                {
                    var pattern = element.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;

                    ToggleState tgState = pattern.Current.ToggleState;
                    if (tgState == ToggleState.Off)
                        pattern.Toggle();

                    //SelectionItemPattern pattern = element.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                    //pattern.Toggle();
                    Thread.Sleep(2000);
                }
                else if (element.Current.ControlType.Equals(ControlType.Image))
                {
                    Input.MoveToAndClick(element);
                    Thread.Sleep(2000);
                }
                else
                {
                    Input.MoveToAndClick(element);
                    Thread.Sleep(2000);
                }
            }
        }

        public static bool SetValueInTextBox(AutomationElement rootElement, string value)
        {
            Condition textPatternAvailable = new PropertyCondition(AutomationElement.IsTextPatternAvailableProperty,
                                                                   true);
            AutomationElement txtElement = rootElement.FindFirst(TreeScope.Descendants, textPatternAvailable);
            if (txtElement != null)
            {
                try
                {
                    Console.WriteLine("Setting value in textbox");
                    var valuePattern = txtElement.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    valuePattern.SetValue(value);
                    Thread.Sleep(2000);
                    ;
                    return true;
                }
                catch
                {
                    Console.WriteLine("Error");
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public static void WaitUntilInstallationComplete(AutomationElement appElement)
        {
            AutomationElement status = null;
            Condition statusPatternAvailable
                = new PropertyCondition(AutomationElement.IsRangeValuePatternAvailableProperty, true);
            int numWaits = 0;
            do
            {
                Console.WriteLine(numWaits + " Waiting for close button");
                status = appElement.FindFirst(TreeScope.Descendants, statusPatternAvailable);
                ++numWaits;
                Thread.Sleep(5000);
            } while (status != null && numWaits < 1000);
        }

        public static AutomationElement GetElementByNameProperty(AutomationElement parentElement, string nameValue)
        {
            Condition condition = new PropertyCondition(AutomationElement.NameProperty, nameValue);
            AutomationElement element = parentElement.FindFirst(TreeScope.Descendants, condition);
            return element;
        }
    }
}