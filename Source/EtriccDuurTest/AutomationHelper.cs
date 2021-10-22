using System;
using System.Collections;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Navigation;
using System.Windows.Shapes;
using TestTools;

namespace EtriccDuurTest
{
    ///<Summary>
    /// Convenience wrappers for UIAutomation
    ///</Summary>
    public static class AutomationHelper
    {
        /// <summary>
        /// Generic Function to Click a Button with Mouse moving over to button and clicking
        /// </summary>
        /// <param name="element">Automation element representing the button</param>
        public static void ClickElement(AutomationElement element)
        {
            if (element == null)
            {
                throw new ArgumentNullException("element");
            }

            Point pt;
#if false
            //some bad automation clients return true and set pt to 0,0
            if(element.TryGetClickablePoint(out pt)) {
                Input.MoveToAndClick(pt);
            }
#else
            //TODO: In the future I may experiment with randomly clicking within the bounds of an element
            Rect rect;
            rect = (Rect)(element.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty));
            pt = new Point(rect.TopLeft.X + 5, rect.TopLeft.Y + 5);
            Input.MoveToAndClick(pt);
#endif

            Thread.Sleep(1000);
        }


        public static AutomationElement FindWindowByTitle(string title)
        {
            AutomationElement returnLE = null;
            PropertyCondition conds = new PropertyCondition(AutomationElement.NameProperty, title);
            returnLE = AutomationElement.RootElement.FindFirst(TreeScope.Children, conds);
            return returnLE;
        }

        public static AutomationElement FindElementByProcessId(int pid)
        {
            int threadSleepTime = 2000;
            AutomationElement element = null;
            PropertyCondition conds = new PropertyCondition(AutomationElement.ProcessIdProperty, pid);
            while (element == null)
            {
                Thread.Sleep(threadSleepTime);
                element = AutomationElement.RootElement.FindFirst(TreeScope.Children, conds);
            }
            Thread.Sleep(threadSleepTime);
            return element;
        }

        public static AutomationElement FindElementByName(AutomationElement Root, String Name)
        {
            AutomationElement returnLE = null;
            PropertyCondition conds = new PropertyCondition(AutomationElement.NameProperty, Name);
            returnLE = Root.FindFirst(TreeScope.Element | TreeScope.Descendants, conds);
            return returnLE;
        }


        public static AutomationElement FindElementByID(AutomationElement Root, String Name)
        {
            AutomationElement returnLE = null;
            PropertyCondition conds = new PropertyCondition(AutomationElement.AutomationIdProperty, Name);
            returnLE = Root.FindFirst(TreeScope.Element | TreeScope.Descendants, conds);
            return returnLE;
        }

        public static AutomationElement FindElementByType(AutomationElement Root, String Type)
        {
            AutomationElement returnLE = null;
            PropertyCondition conds = new PropertyCondition(AutomationElement.ClassNameProperty, Type);
            returnLE = Root.FindFirst(TreeScope.Element | TreeScope.Descendants, conds);
            return returnLE;
        }

        public static AutomationElementCollection FindAllElementsOfType(AutomationElement Root, String Type)
        {
            AutomationElementCollection returnLE;
            PropertyCondition conds = new PropertyCondition(AutomationElement.ClassNameProperty, Type);
            returnLE = Root.FindAll(TreeScope.Element | TreeScope.Descendants, conds);
            return returnLE;
        }

        public static AutomationElementCollection FindAllElementsOfName(AutomationElement Root, String Name)
        {
            AutomationElementCollection returnLE;
            PropertyCondition conds = new PropertyCondition(AutomationElement.NameProperty, Name);
            returnLE = Root.FindAll(TreeScope.Element | TreeScope.Descendants, conds);
            return returnLE;
        }

        public static AutomationElement FindRawElementByID(AutomationElement Root, String Name)
        {
            AutomationElement cur = null;
            AutomationElement result;

            for (cur = TreeWalker.RawViewWalker.GetFirstChild(Root); cur != null; cur = TreeWalker.RawViewWalker.GetNextSibling(cur))
            {
                if (TreeWalker.RawViewWalker.GetFirstChild(cur) != null)
                {
                    result = FindRawElementByID(cur, Name);
                    if (result != null)
                    {
                        return result;
                    }
                }

                Object t = cur.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty);
                if (t.ToString() == Name)
                {
                    return cur;
                }
            }
            return null;
        }
    }
}
