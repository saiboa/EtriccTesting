using System;
using System.Collections;
using System.Threading;
//using System.Windows.Threading;
using System.Windows;
using System.Windows.Automation;
using TestTools;

namespace EtriccDuurTest
{
    ///<Summary>
    /// Base class for writing automation scenarios eg leak test or stress tests
    ///</Summary>
    public class BaseScenario : MarshalByRefObject
    {
        protected const int clickInterval = 6000;
        protected const int borderPadding = 3;
        AutomationElement _rootElement = null;

        public BaseScenario()
        {
        }

        public BaseScenario(AutomationElement rootElement)
        {
            if (rootElement == null)
            {
                throw new ArgumentNullException("rootElement");
            }

            Initialize(rootElement);
        }

        protected AutomationElement RootElement { get { return _rootElement; } }
        protected AutomationElement aeOverviewWindow = null;
        protected AutomationElement aeAgvOverview = null;

        public WindowInteractionState WindowInteractionState
        {
            get
            {
                if (RootElement != null)
                {
                    object pattern;
                    if (RootElement.TryGetCurrentPattern(WindowPattern.Pattern, out pattern))
                    {
                        return ((WindowPattern)pattern).Current.WindowInteractionState;
                    }
                }
                return WindowInteractionState.NotResponding; //TODO: is this really a good cop out?
            }
        }

        protected virtual void Initialize(AutomationElement rootElement)
        {
            if (rootElement == null)
            {
                throw new ArgumentNullException("rootElement");
            }
            _rootElement = rootElement;
        }

        public void FindScenarioByWindowTitle(string title)
        {
            Console.WriteLine("Trying to find scenario with window title {0}", title);
            AutomationElement element = null;
            while (element == null)
            {
                element = AutomationHelper.FindWindowByTitle(title);
            }
            Thread.Sleep(clickInterval * 2);
            Initialize(element);
        }

        public void FindScenarioByPid(int pid)
        {
            Console.WriteLine("Trying to find scenario with process id {0}", pid);
            AutomationElement element = AutomationHelper.FindElementByProcessId(pid);
            Initialize(element);
        }

        public void BringInFocus()
        {
            if (null != RootElement)
            {
                RootElement.SetFocus();
            }
        }

    }
}
