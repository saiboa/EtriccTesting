using System;
using System.Threading;
using System.Windows.Automation;


namespace QATestProInstallationProcudure
{
    internal class ProMonitorUtilities
    {
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

        public static AutomationElement GetMainWindowFromName(string mainFormName)
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
                        if (aeAllWindows[i].Current.Name.Equals(mainFormName))
                        {
                            aeWindow = aeAllWindows[i];
                            Console.WriteLine("aeWindow[" + i + "]=" + aeWindow.Current.Name);
                            Thread.Sleep(2000);
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
    }
}
