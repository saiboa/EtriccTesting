using System;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Reflection;

using System.Diagnostics;
using Microsoft.Win32;

using System.Net.Mail;

using System.Windows;
using System.Windows.Input;
using System.Windows.Automation;

using TestTools;

namespace EtriccDuurTest
{
    public class InvokeTest : BaseScenario
    {
        //Utility object I keep around for doing various random things
        Random _random = new Random();

        public delegate void StressAction();
        [STAThread]
        static void Main(string[] args)
        {
            bool showHelp = false;
            int pid = -1;
            string title = null;
            TimeSpan ttl = TimeSpan.FromHours(24);

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].ToLower())
                {
                    case "-?":
                    case "-h":
                    case "-help": { showHelp = true; i = args.Length; break; }
                    case "-pid": { pid = int.Parse(args[++i]); break; }
                    case "-title": { title = args[++i]; break; }
                    case "-ttl": { ttl = TimeSpan.Parse(args[++i]); break; }

                    default:
                        {
                            Console.Error.WriteLine("Unknown switch {0}", args[i]);
                            showHelp = true;
                            break;
                        }
                }
            }

            if (showHelp)
            {
                Usage();
                return;
            }

            if (title != null && pid != -1)
            {
                Console.Error.WriteLine("Only one of -title or -pid can be specified");
                Usage();
                return;
            }

            InvokeTest scenario = null;
            if (title != null)
            {
                scenario = new InvokeTest();
                scenario.FindScenarioByWindowTitle(title);
            }
            else
            {
                if (pid != -1)
                {
                    scenario = new InvokeTest();
                    scenario.FindScenarioByPid(pid);
                }
            }

            if (scenario == null)
            {
                Console.Error.WriteLine("One of -title or -pid must be specified");
                Usage();
                return;
            }

            Console.Out.WriteLine("Stressing scenario for {0}days {1}h {2}m {3}s", ttl.Days, ttl.Hours, ttl.Minutes, ttl.Seconds);
            Run(scenario, ttl);

            return;
        }

        public static void Usage()
        {
            Console.Out.WriteLine("{0} [-? | -h | -help] [-pid ProcessId | -title Window Title] [-ttl time-to-live]", Process.GetCurrentProcess().ProcessName);
            Console.Out.WriteLine("-ttl format hh:mm:ss");
        }

        public static void Run(InvokeTest scenario, TimeSpan timeToLive)
        {
            /*
            // find Infrastructuur
            string panelName = "Infrastructure";
            Console.WriteLine(panelName + " start find time: " + System.DateTime.Now.ToString("HH:mm:ss"));
            
            // Find ToolBar
            System.Windows.Automation.Condition c1 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "stackStrip1"),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar)
            );
            AutomationElement aeToolBar
                = scenario.RootElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c1);

            if (aeToolBar == null)
            {
                Console.WriteLine(panelName + " failed find time1: " + System.DateTime.Now.ToString("HH:mm:ss"));
                //result = Constants.TEST_FAIL;
                return;
            }
            else
                Console.WriteLine(panelName + " ToolBar Found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Thread.Sleep(2000);

            // Find "Infrastructure" Button Element
            System.Windows.Automation.Condition c2 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "Infrastructure"),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
            );

            AutomationElement aeInfrastructure
                = aeToolBar.FindFirst(TreeScope.Element | TreeScope.Children, c2);

            if (aeInfrastructure == null)
            {
                Console.WriteLine(panelName + " Infrastructure not found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                //result = Constants.TEST_FAIL;
                return;
            }
            else
                Console.WriteLine(panelName + " Infrastructure Found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Thread.Sleep(2000);
            InvokePattern ipX = (InvokePattern)aeInfrastructure.GetCurrentPattern(InvokePattern.Pattern);
            ipX.Invoke();

            Thread.Sleep(1000);

            // Find "System Overview" HyperLink Element
            System.Windows.Automation.Condition c3 = new AndCondition(
                new PropertyCondition(AutomationElement.NameProperty, "System Overview"),
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Hyperlink)
             );

            AutomationElement aeSystemOverview
               = scenario.RootElement.FindFirst(TreeScope.Element | TreeScope.Descendants, c3);

            if (aeSystemOverview == null)
            {
                Console.WriteLine(panelName + " failed to find System Overview at time3: " + System.DateTime.Now.ToString("HH:mm:ss"));
                //result = Constants.TEST_FAIL;
                return;
            }
            else
                Console.WriteLine(panelName + " System Overview Button found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            InvokePattern ip3 = (InvokePattern)aeSystemOverview.GetCurrentPattern(InvokePattern.Pattern);
            ip3.Invoke();
            */

            // Find System Overview Window
            scenario.aeOverviewWindow = AUIUtilities.FindElementByID("1", scenario.RootElement);
            if (scenario.aeOverviewWindow == null)
            {
                Console.WriteLine("FindElementByID failed:" + "aeOverviewWindow");
                return;
            }
            else
                Console.WriteLine(" System Overview window found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Thread.Sleep(3000);

            // Find Agv Overview Window
            scenario.aeAgvOverview = AUIUtilities.FindElementByID("2", scenario.RootElement);
            if (scenario.aeAgvOverview == null)
            {
                Console.WriteLine("FindElementByID failed:" + "aeAgvOverview");
                return;
            }
            else
                Console.WriteLine(" System aeAgvOverview window found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Thread.Sleep(3000);


            // Find Button Zomm to Content ----------------------------------------------------------
            AutomationElement aeBtnZoomToContent = AUIUtilities.FindElementByID("m_BtnZoomToContent", scenario.aeOverviewWindow);
            if (aeBtnZoomToContent == null)
            {
                Console.WriteLine("Find aeBtnZoomToContent failed:" + "aeBtnZoomToContent");
                return;
            }
            else
                Console.WriteLine("Button aeBtnZoomToContent found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Thread.Sleep(2000);
            Point ZoomPoint = AUIUtilities.GetElementCenterPoint(aeBtnZoomToContent);
            Input.MoveTo(ZoomPoint);
            Thread.Sleep(2000);
            Input.MoveToAndClick(ZoomPoint);
            Console.WriteLine("aeBtnZoomToContent invoked ");
            //---------------------------------------------------------------------------------------

            //A set of delegates I randomly call to perform ui actions
            //TODO ADD new test actions
            StressAction[] actions = new StressAction[] 
            {
                scenario.WheelWindow,  //OK
                scenario.ZoomToContent,
                scenario.IncreaseLarge,
                scenario.MoveThumbUp,   //OK
                scenario.MoveThumbDown, //OK
                scenario.ClickCheckBox, //OK
                scenario.LayerCheckBox,
                scenario.ZoomToContent,
                scenario.ZoomToContent,
                //scenario.AgvDataGridView,
                //scenario.ResizeWindow,  //OK
                //scenario.FocusRandomInvokableElement,
                //scenario.InvokeFocusedElement,
                //scenario.KeyNavigate
                //scenario.TypeInTextBox //TODO
            };

            Random random = new Random();
            Process myProcess = Process.GetProcessById(scenario.RootElement.Current.ProcessId);
            DateTime end = DateTime.Now + timeToLive;

            Console.Out.WriteLine("root element name: " + scenario.RootElement.Current.Name);

            // Add Open window Event Handler
            AutomationEventHandler UIAWindowEventHandler = new AutomationEventHandler(OnOpenWindowUIAEvent);
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                scenario.RootElement, TreeScope.Descendants, UIAWindowEventHandler);


            //drive stress until the target process exits or my time expires
            while (false == myProcess.HasExited && DateTime.Now < end)
            {
                try
                {
                    //wait for the target process to become responsive
                    myProcess.WaitForInputIdle(random.Next(50, 8000));
                }
                catch (InvalidOperationException)
                {
                    break;
                }

                try
                {
                    //if the the target process is ready for user interaction then execute a random stress action
                    //TODO:     if the target window opens a modal dialog then it is never ready for user interaction until the 
                    //          modal dialog is dismissed, some extra logic is needed to detect when modal dialogs are created and dismissed
                    if (scenario.WindowInteractionState == WindowInteractionState.ReadyForUserInteraction)
                    {
                        if (actions.Length > 0)
                        {
                            scenario.BringInFocus(); //some other UI may have obscured the target process's window so restore the target window
                            actions[random.Next(0, actions.Length)]();
                        }
                    }
                }
                catch (Exception e)
                { //normally an exception this broad is bad but I want the scenario to keep running until the driven app dies
                    if (false == myProcess.HasExited)
                    {
                        Console.Error.WriteLine(e);
                    }
                }
            }
        }

        #region OpenWindow Event
        public static void OnOpenWindowUIAEvent(object src, AutomationEventArgs args)
        {
            AutomationElement element;
            try
            {
                element = src as AutomationElement;
            }
            catch
            {
                return;
            }

            string name = "";
            if (element == null)
                name = "null";
            else
            {
                name = element.GetCurrentPropertyValue(
                    AutomationElement.NameProperty) as string;
            }

            if (name.Length == 0) name = "<NoName>";
            string str = string.Format("LayoutSaveAs:={0} : {1}", name, args.EventId.ProgrammaticName);
            Console.WriteLine(str);

            //if (name.Equals("Save Shell Layout"))
            //{
            string BtnCancelID = "m_BtnCancel";
            // Find and Click Save Button

            Console.WriteLine("FindElementAndClick: " + BtnCancelID);
            // Set a property condition that will be used to find the control.
            System.Windows.Automation.Condition c = new PropertyCondition(
            AutomationElement.AutomationIdProperty, BtnCancelID, PropertyConditionFlags.IgnoreCase);

            // Find the element.
            AutomationElement ae = element.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
            if (ae != null)
            {
                Thread.Sleep(3000);
                InvokePattern ip = (InvokePattern)ae.GetCurrentPattern(InvokePattern.Pattern);
                ip.Invoke();
            }

            //}
            //else
            //    Console.WriteLine("Name is ------------:" + name);
        }
        #endregion event

        #region Stress Actions
        //======================================================
        // mouse wheel forward and backward
        public void WheelWindow()
        {
            Console.WriteLine("---- WheelWindow ----");
            Point WheelPoint = AUIUtilities.GetElementCenterPoint(aeOverviewWindow);
            Input.MoveToAndClick(WheelPoint);
            Thread.Sleep(1000);

            for (int i = 1; i < 6; i++)
            {
                Input.SendMouseInput(WheelPoint.X, WheelPoint.Y, 20, SendMouseInputFlags.Wheel);
                Thread.Sleep(1000);
            }

            for (int i = 1; i < 6; i++)
            {
                Input.SendMouseInput(WheelPoint.X, WheelPoint.Y, -20, SendMouseInputFlags.Wheel);
                Thread.Sleep(1000);
            }
        }
        // To See the Whole Layout
        public void ZoomToContent()
        {
            Console.WriteLine("---- ZoomToContent ----");
            AutomationElement aeBtnZoomToContent = AUIUtilities.FindElementByID("m_BtnZoomToContent", aeOverviewWindow);
            if (aeBtnZoomToContent == null)
            {
                Console.WriteLine("Find aeBtnZoomToContent failed:" + "aeBtnZoomToContent");
                return;
            }
            else
                Console.WriteLine("Button aeBtnZoomToContent found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Thread.Sleep(1000);
            Point ZoomPoint = AUIUtilities.GetElementCenterPoint(aeBtnZoomToContent);
            Input.MoveTo(ZoomPoint);
            Thread.Sleep(1000);
            Input.MoveToAndClick(ZoomPoint);
            Console.WriteLine("aeBtnZoomToContent invoked ");
            Thread.Sleep(2000);
        }

        public void IncreaseLarge()
        {
            Console.WriteLine("---- IncreaseLarge ----");
            AutomationElement aeIncreaseLarge = AUIUtilities.FindElementByID("IncreaseLarge", aeOverviewWindow);
            if (aeIncreaseLarge == null)
            {
                Console.WriteLine("Find IncreaseLarge failed:" + "IncreaseLarge");
                return;
            }
            else
                Console.WriteLine("Button IncreaseLarge found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Point IncreasePoint = AUIUtilities.GetElementCenterPoint(aeIncreaseLarge);
            Input.MoveToAndClick(IncreasePoint);
            Thread.Sleep(1000);
            Input.MoveToAndClick(IncreasePoint);
            Thread.Sleep(1000);
            Input.MoveToAndClick(IncreasePoint);
            Thread.Sleep(1000);
            Input.MoveToAndClick(IncreasePoint);
            Thread.Sleep(1000);
            Input.MoveToAndClick(IncreasePoint);
            Thread.Sleep(1000);
            Input.MoveToAndClick(IncreasePoint);
            Thread.Sleep(1000);
            Input.MoveToAndClick(IncreasePoint);
        }

        public void MoveThumbUp()
        {
            Console.WriteLine("---- MoveThumbUp ----");
            AutomationElement aeBtnThumb = AUIUtilities.FindElementByID("Thumb", aeOverviewWindow);
            if (aeBtnThumb == null)
            {
                Console.WriteLine("Find aeBtnThumb failed:" + "aeBtnThumb");
                return;
            }
            else
                Console.WriteLine("Button aeBtnThumb found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Point point;
            if (aeBtnThumb.TryGetClickablePoint(out point))
            {
                Input.MoveTo(point);
                Console.WriteLine("Point X: " + point.X);
                Console.WriteLine("Point Y: " + point.Y);
            }

            Input.SendMouseInput(point.X, point.Y, 0, SendMouseInputFlags.LeftDown | SendMouseInputFlags.Absolute);
            Console.WriteLine("first time  up");
            Thread.Sleep(2000);

            Console.WriteLine("after 3 second ");
            Point NewPoint = point;

            for (int i = 1; i < 6; i++)
            {
                NewPoint.X = point.X;
                NewPoint.Y = point.Y - i * 10;
                Input.MoveTo(NewPoint);
                Console.WriteLine(i + " st move up ");
                Thread.Sleep(1000);
            }
            //Thread.Sleep(1000);
            //Input.MoveTo(point);
            Thread.Sleep(1000);
            Input.SendMouseInput(NewPoint.X, NewPoint.Y, 0, SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);
        }

        public void MoveThumbDown()
        {
            Console.WriteLine("---- MoveThumbDown ----");
            AutomationElement aeBtnThumb = AUIUtilities.FindElementByID("Thumb", aeOverviewWindow);
            if (aeBtnThumb == null)
            {
                Console.WriteLine("Find aeBtnThumb failed:" + "aeBtnThumb");
                return;
            }
            else
                Console.WriteLine("Button aeBtnThumb found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Point point;
            if (aeBtnThumb.TryGetClickablePoint(out point))
            {
                Input.MoveTo(point);
                Console.WriteLine("Point X: " + point.X);
                Console.WriteLine("Point Y: " + point.Y);
            }

            Input.SendMouseInput(point.X, point.Y, 0, SendMouseInputFlags.LeftDown | SendMouseInputFlags.Absolute);
            Console.WriteLine("first time  down");
            Thread.Sleep(2000);

            Console.WriteLine("after 3 second ");
            Point NewPoint = point;

            for (int i = 1; i < 6; i++)
            {
                NewPoint.X = point.X;
                NewPoint.Y = point.Y + i * 10;
                Input.MoveTo(NewPoint);
                Console.WriteLine(i + " st move  down");
                Thread.Sleep(1000);
            }
            //Thread.Sleep(1000);
            //Input.MoveTo(point);
            Thread.Sleep(1000);
            Input.SendMouseInput(NewPoint.X, NewPoint.Y, 0, SendMouseInputFlags.LeftUp | SendMouseInputFlags.Absolute);
        }

        public void ClickCheckBox()
        {
            Console.WriteLine("---- ClickCheckBox ----");
            AutomationElement ae = AUIUtilities.FindElementByID("checkBox1", aeOverviewWindow);
            if (ae == null)
            {
                Console.WriteLine("Find checkBox1 failed:" + "checkBox1");
                return;
            }
            else
                Console.WriteLine("Button checkBox1 found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Point CheckPoint = AUIUtilities.GetElementCenterPoint(ae);
            Input.MoveToAndClick(CheckPoint);
            Thread.Sleep(1000);
            TogglePattern tg = ae.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
            ToggleState tgTState = tg.Current.ToggleState;
            tg.Toggle();

            Thread.Sleep(2000);
        }

        public void LayerCheckBox()
        {
            Console.WriteLine("---- LayerCheckBox ----");
            try
            {
                AutomationElement ae = AUIUtilities.FindElementByType(ControlType.ComboBox, aeOverviewWindow);

                if (ae == null)
                {
                    Console.WriteLine("Find LayerCheckBox failed:" + "LayerCheckBox");
                    return;
                }
                else
                    Console.WriteLine("LayerCheckBox found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

                Point CheckPoint = AUIUtilities.GetElementCenterPoint(ae);
                Input.MoveToAndClick(CheckPoint);
                Thread.Sleep(1000);

                int index = _random.Next(0, 3);
                Console.WriteLine("index is :" + index);
                index = _random.Next(0, 3);
                Console.WriteLine("index is :" + index);
                index = _random.Next(0, 3);
                Console.WriteLine("index is :" + index);
                index = _random.Next(0, 3);
                Console.WriteLine("index is :" + index);
                AutomationElement aeLayer = AUIUtilities.FindListItemByIndex(ae, index);
                if (aeLayer == null)
                {
                    Console.WriteLine("Find aeLayer failed:" + "aeLayer");
                    return;
                }
                else
                    Console.WriteLine("LayerCheckBox aeLayer found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));


                AutomationElement aeCheckBox = AUIUtilities.FindElementByType(ControlType.CheckBox, aeLayer);
                if (aeCheckBox == null)
                {
                    Console.WriteLine("Find aeCheckBox failed:" + "aeLayer");
                    return;
                }
                else
                    Console.WriteLine("LayerCheckBox aeCheckBox found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

                Point BoxPoint = AUIUtilities.GetElementCenterPoint(aeLayer);
                Input.MoveTo(BoxPoint);
                Thread.Sleep(1000);

                TogglePattern tg = aeCheckBox.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                ToggleState tgTState = tg.Current.ToggleState;
                tg.Toggle();
                Thread.Sleep(2000);

                Input.MoveToAndClick(ae);
                Thread.Sleep(1000);

            }
            catch (Exception ex)
            {
                Console.WriteLine("LayerCheckBoxt exception:" + ex.Message + "----" + ex.StackTrace);
            }


            Thread.Sleep(3000);
        }

        public void ResizeWindow()
        {
            Console.WriteLine("---- ResizeWindow ----");
            TransformPattern tp;
            object tpo;
            if (RootElement.TryGetCurrentPattern(TransformPattern.Pattern, out tpo))
            {
                tp = (TransformPattern)tpo;
                if (tp.Current.CanResize)
                {
                    System.Drawing.Size size = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Size;
                    tp.Resize(_random.Next(0, size.Width), _random.Next(0, size.Height));
                }
            }
        }

        //======================================================
        public void AgvDataGridView()
        {
            AutomationElement ae = AUIUtilities.FindElementByID("m_DataGridView", aeAgvOverview);
            if (ae == null)
            {
                Console.WriteLine("Find AgvDataGridView failed:" + "AgvDataGridView");
                return;
            }
            else
                Console.WriteLine("Button AgvDataGridView found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));

            Point point = AUICommon.GetDataGridViewCellPointAt(13, "Id", ae);
            Input.MoveToAndDoubleClick(point);
            Thread.Sleep(2000);
            string cellValue = AUICommon.GetDataGridViewCellValueAt(13, "Id", ae);
            Console.WriteLine("End id: " + cellValue);

            Thread.Sleep(2000000);
        }
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++
      
        //Try and randomly get a potentially keyboard focusable element and set focus on that element
        public void FocusRandomInvokableElement()
        {
            AutomationElement focusableElement = GetRandomElement(IsFocusable);
            if (focusableElement != null)
            {
                if (true == (bool)focusableElement.GetCurrentPropertyValue(AutomationElement.IsKeyboardFocusableProperty))
                {
                    //Keep away from the win32 MainMenu it doesnt like to be focused
                    if (false == MatchesCriteria(focusableElement, "MainMenu", ControlType.Menu))
                    {
                        Console.Out.WriteLine("Setting keyboard focus on {0}", GetName(focusableElement));
                        Point point;
                        if (focusableElement.TryGetClickablePoint(out point))
                        {
                            Input.MoveTo(point);
                        }
                        focusableElement.SetFocus();
                    }
                }
            }
        }

        //Use a random method (mouseclick\keypress\hypnosis) to invoke the focused element
        public void InvokeFocusedElement()
        {
            AutomationElement focusedElement = AutomationElement.FocusedElement;

            //sometimes the focused element is part of the system menu and we dont want to invoke that do we?                
            while (IsInTitleBar(focusedElement))
            {
                PressKey(System.Windows.Input.Key.Escape);
                focusedElement = AutomationElement.FocusedElement;
            }

            if (focusedElement != null)
            {
                InvokeElement(focusedElement);
            }
        }

        //Todo find a scheme that allows me to combine keys and key modifiers
        static System.Windows.Input.Key[] FocusChangingKeys = 
        {
            System.Windows.Input.Key.Tab //, //tab
            /*            
            Key.LeftShift | Key.Tab, //shift + tab
            Key.RightShift | Key.Tab, 
            
            Key.LeftCtrl  | Key.Tab, //ctrl + tab 
            Key.RightCtrl | Key.Tab, 
            
            Key.LeftCtrl  | Key.LeftShift | Key.Tab, //ctrl + shift + tab
            Key.LeftCtrl  | Key.RightShift | Key.Tab,
            Key.RightCtrl | Key.LeftShift | Key.Tab,
            Key.RightCtrl | Key.RightShift | Key.Tab,
            */
        };

        static System.Windows.Input.Key[] NavigationKeys = 
        {
            System.Windows.Input.Key.Tab,
            System.Windows.Input.Key.Left,
            System.Windows.Input.Key.Right,
            System.Windows.Input.Key.Up,
            System.Windows.Input.Key.Down,
            System.Windows.Input.Key.PageUp,
            System.Windows.Input.Key.PageDown,
            System.Windows.Input.Key.Home,
            System.Windows.Input.Key.End
        };

        public void KeyNavigate()
        {
            Key key = GetRandomKey(NavigationKeys);
            Console.Out.WriteLine("pressing {0}", key);
            PressKey(key);
        }
        #endregion //Stress Actions

        //Keys commonly used by windows to invoke ui
        //TODO: investigate querying the access keys so that I can invoke shortcuts to menus etc
        static System.Windows.Input.Key[] InvokingKeys = 
        {
            System.Windows.Input.Key.Enter, 
            System.Windows.Input.Key.Space, 
            System.Windows.Input.Key.Escape, 
        };

        //The different ways an AutomationElement may be invoked
        public enum InvokeMethods
        {
            UIAutomation,
            MouseClick,
            KeyPress,
            Count // This value exists as a cheap trick to keep track of the number of elements in the enum
        }

        void InvokeElement(AutomationElement e)
        {
            if (e == null)
            {
                throw new ArgumentNullException("e");
            }

            string elementName = GetName(e);

            switch ((InvokeMethods)_random.Next(0, (int)InvokeMethods.Count))
            {
                case InvokeMethods.UIAutomation:
                    {
#if false          
                    //I may not want to do this because I might invoke items that were hidden or out of view which is not true of real user interaction
                    object invokableObject = null;
                    if(true == e.TryGetCurrentPattern(InvokePattern.Pattern, out invokableObject)) {
                        if(null != invokableObject) {
                            Console.Out.WriteLine("Invoking on {0}", elementName);
                            ((InvokePattern)invokableObject).Invoke();
                        }
                    }
#endif
                        break;
                    }

                case InvokeMethods.MouseClick:
                    {
                        Console.Out.WriteLine("Clicking on {0}", elementName);
                        AutomationHelper.ClickElement(e);
                        break;
                    }

                //TODO: Be mean
                //Im being nice here, I could hold down a key for a while 
                //or just press but do not release
                //or vice versa
                case InvokeMethods.KeyPress:
                    {
                        Key key = GetRandomKey(InvokingKeys);
                        Console.Out.WriteLine("pressing {0} on {1} at {2}", key, elementName, System.DateTime.Now);
                        PressKey(key);
                        break;
                    }

                default:
                    {
                        throw new ApplicationException("No case label for switch statement value");
                    }
            }
        }

        Key GetRandomKey(Key[] keys)
        {
            if (keys.Length > 0)
            {
                return keys[_random.Next(0, keys.Length)];
            }
            return Key.None;
        }

        void PressKey(Key key)
        {
            Input.SendKeyboardInput(key, true);
            Input.SendKeyboardInput(key, false);
        }

        /*TODO:
            For certain applications there could be hundreds or thousands of invokeable elements (think of a page in IE and all its hyperlinks)
            It could take several seconds to enumerate them all and then randomly pick one to return
            I currently clamp my search to a fixed number of elements, which means some UI is never exercised
            A smarter algo would be to start my walk at the currently focused element and traverse parents as well as children
        */
        const int MaxElementsToSearch = 55;
        AutomationElement GetRandomElement(Predicate<AutomationElement> condition /* a filter that rejects elements that cause the predicate to return false */)
        {
            int elementCount = 0;
            int elementLimit = MaxElementsToSearch;
            List<AutomationElement> elements = new List<AutomationElement>();
            FindAllElements(TreeWalker.RawViewWalker, RootElement, elements, condition, ref elementCount, elementLimit);
            if (elements.Count < 1)
            {
                return null;
            }
            else
            {
                return elements[_random.Next(0, elements.Count)];
            }
        }

        static string GetName(AutomationElement e)
        {
            return
            String.Format
            (
                "{0}:[{1}]",
                e.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty),
                e.GetCurrentPropertyValue(AutomationElement.ClassNameProperty)
            );
        }

        static AutomationElement FindFirstElement(TreeWalker walker, AutomationElement root, Predicate<AutomationElement> criteria)
        {
            int elementCount = 0;
            return FindFirstElement(walker, root, criteria, ref elementCount, int.MaxValue);
        }

        static AutomationElement FindFirstElement(TreeWalker walker, AutomationElement root, Predicate<AutomationElement> criteria, ref int elementCount, int elementLimit)
        {
            if (root == null)
            {
                throw new ArgumentNullException("Root");
            }

            if (criteria == null)
            {
                throw new ArgumentNullException("criteria");
            }

            if (elementCount >= elementLimit)
            {
                return null;
            }

            AutomationElement result;

            for (AutomationElement cur = walker.GetFirstChild(root); cur != null; cur = walker.GetNextSibling(cur))
            {
                elementCount++;

                if (elementCount >= elementLimit)
                {
                    return null;
                }

                if (true == IsIgnorable(cur))
                {
                    continue;
                }

                if (true == criteria(cur))
                {
                    return cur;
                }

                result = FindFirstElement(walker, cur, criteria, ref elementCount, elementLimit);
                if (result != null)
                {
                    return result;
                }
            }

            return null;
        }

        static void FindAllElements(TreeWalker walker, AutomationElement root, List<AutomationElement> results, Predicate<AutomationElement> criteria)
        {
            int elementCount = 0;
            FindAllElements(walker, root, results, criteria, ref elementCount, int.MaxValue);
        }

        static void FindAllElements(TreeWalker walker, AutomationElement root, List<AutomationElement> results, Predicate<AutomationElement> criteria, ref int elementCount, int elementLimit)
        {
            if (root == null)
            {
                throw new ArgumentNullException("Root");
            }

            if (results == null)
            {
                throw new ArgumentNullException("results");
            }

            if (criteria == null)
            {
                throw new ArgumentNullException("criteria");
            }

            if (elementCount >= elementLimit)
            {
                return;
            }

            for (AutomationElement cur = walker.GetFirstChild(root); cur != null; cur = walker.GetNextSibling(cur))
            {
                elementCount++;

                if (elementCount >= elementLimit)
                {
                    return;
                }

                if (true == IsIgnorable(cur))
                {
                    continue;
                }

                if (true == criteria(cur))
                {
                    results.Add(cur);
                }

                FindAllElements(walker, cur, results, criteria, ref elementCount, elementLimit);
            }
        }

        ///Returns true for automation elements to ignore
        ///e.g. The title bar and its system menu are good candidates to ignore because they are likely to close your app
        ///     disabled elements are also ignored
        static bool IsIgnorable(AutomationElement e)
        {
            if (MatchesCriteria(e, "TitleBar", ControlType.TitleBar))
            {
                return true;
            }

            if (MatchesCriteria(e, "AppControlToolBar", ControlType.ToolBar))
            {
                return true;
            }

            if (false == (bool)e.GetCurrentPropertyValue(AutomationElement.IsEnabledProperty))
            {
                return true;
            }

            return false;
        }

        //Useful filter for identifying very specific elements in your UI
        //returns true if element e has the given automation id and control type 
        //eg your system menu always has automation id TitleBar and ControlType TitleBar
        //UISpy that comes with the platform sdk is a great tool for inspecting UI and finding such properties
        static bool MatchesCriteria(AutomationElement e, string automationId, ControlType controlType)
        {
            ControlType type = (ControlType)e.GetCurrentPropertyValue(AutomationElement.ControlTypeProperty);
            if (type.Equals(controlType))
            {
                string id = (string)e.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty);
                if (0 == string.CompareOrdinal(id, automationId))
                {
                    return true;
                }
            }
            return false;
        }

        static bool IsFocusable(AutomationElement e)
        {
            return (bool)e.GetCurrentPropertyValue(AutomationElement.IsKeyboardFocusableProperty)
                    && true == (bool)e.GetCurrentPropertyValue(AutomationElement.IsEnabledProperty);
        }

        static bool IsClickable(AutomationElement e)
        {
            Point point = new Point(0, 0);
            return (bool)e.GetCurrentPropertyValue(AutomationElement.IsInvokePatternAvailableProperty)
                   && false == (bool)e.GetCurrentPropertyValue(AutomationElement.IsOffscreenProperty)
                   && true == e.TryGetClickablePoint(out point);
        }

        static bool IsInvokable(AutomationElement e)
        {
            return (bool)e.GetCurrentPropertyValue(AutomationElement.IsInvokePatternAvailableProperty);
        }

        static bool IsInTitleBar(AutomationElement e)
        {
            while (e != null)
            {
                if (e == AutomationElement.RootElement)
                {
                    return false;
                }

                ControlType controlType = (ControlType)e.GetCurrentPropertyValue(AutomationElement.ControlTypeProperty);
                if (controlType.Equals(ControlType.TitleBar))
                {
                    string automationId = (string)e.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty);
                    if (0 == string.CompareOrdinal("TitleBar", automationId))
                    {
                        return true;
                    }
                }

                e = TreeWalker.ControlViewWalker.GetParent(e);
            }
            return false;
        }
    }
}
