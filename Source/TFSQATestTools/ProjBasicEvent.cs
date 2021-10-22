using System;
using System.Threading;
using System.Windows.Automation;
using System.IO;
using TestTools;

namespace TFSQATestTools
{
    public class ProjBasicEvent
    {
        private static string sErrorMessage;
        private static bool sEventEnd;

        public static bool IsEventEnd
        {
            get
            {
                return sEventEnd;
            }
            private set
            {
                // Can only be called in this class.
                sEventEnd = value;
            }
        }

        public static string HasErrorMessage
        {
            get
            {
                return sErrorMessage;
            }
            set
            {
                sErrorMessage = value;
            }
        }

        public static void OnInstallApplicationEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("Begin OnInstallApplicationEvent");
            AutomationElement element;
            string name = "";
            try
            {
                element = src as AutomationElement;
                if (element == null)
                {
                    name = "null";
                    return;
                }
                else
                {
                    name = element.GetCurrentPropertyValue(
                        AutomationElement.NameProperty) as string;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Excep OnInstallApplicationEvent:" + ex.Message + " ---- " + ex.StackTrace);
                return;
            }

            try
            {    
                #region
                if (name.Length == 0) name = "<NoName>";
                string str = string.Format("OnInstallApplicationEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
                Console.WriteLine(str);

                Thread.Sleep(2000);
                if (name.Equals("Windows Installer"))   // Another Version installed window
                {
                    Console.WriteLine("<Windows Installer> <------------>:" + name);
                    bool anotherVersion = false;
                    AutomationElementCollection aeAllTexts = null;
                    // find windows text
                    System.Windows.Automation.Condition cText = new PropertyCondition(
                        AutomationElement.ControlTypeProperty, ControlType.Text);
                    aeAllTexts = element.FindAll(TreeScope.Descendants, cText);
                    Console.WriteLine("aeAllTexts.Count:   " + aeAllTexts.Count);
                    for (int i = 0; i < aeAllTexts.Count; i++)
                    {
                        if (aeAllTexts[i].Current.Name.StartsWith("Another version of this product is already installed"))
                        {
                            anotherVersion = true;
                            AutomationElement aeOKBtn = TestTools.AUIUtilities.FindElementByName("OK", element);
                            if (aeOKBtn != null)
                            {
                                TestTools.Input.MoveToAndClick(aeOKBtn);
                                Console.WriteLine("WriteLine:anotherversion of this product is already installed");
                                StreamWriter write = File.CreateText(Path.Combine(@"C:\EtriccTests", "event.txt"));
                                write.WriteLine("anotherversion of this product is already installed");
                                write.Close();
                                Thread.Sleep(5000);
                            }
                        }
                    }

                    if (anotherVersion)
                    {
                        sErrorMessage = "AnotherVersionInstalled";
                        sEventEnd = true;
                        return;
                    }
                }
                else if (name.Equals("Open File - Security Warning"))
                {
                    //Epia3Common.WriteTestLogMsg(slogFilePath, "open window name: " + name, sOnlyUITest);
                    Console.WriteLine("<Click Run Button> <------------>:" + name);
                    System.Windows.Automation.Condition c = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, "Run"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                    );

                    // Find the element.
                    AutomationElement aeRun = element.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeRun != null)
                    {
                        System.Windows.Point pt = TestTools.AUIUtilities.GetElementCenterPoint(aeRun);
                        TestTools.Input.MoveTo(pt);
                        Thread.Sleep(1000);
                        TestTools.Input.ClickAtPoint(pt);
                    }
                    else
                    {
                        Console.WriteLine("<Run Button NOT Found> <------------>:" + name);
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("<NOT Do Other Name> -----------:" + name);
                    //Epia3Common.WriteTestLogMsg(slogFilePath, "SERVER open other window name: " + name, sOnlyUITest);
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine("OnInstallAppEvent-Exception:" + ex.Message);
                //throw;
            }

            Console.WriteLine("OnInstallAppEvent-End");
            sEventEnd = true;
        }

        #region OnUninstallAppEvent
        public static void OnUninstallAppEvent(object src, AutomationEventArgs args)
        {
            Console.WriteLine("OnUninstallAppEvent-Begin");
            AutomationElement element;
            string name = "";
            try
            {
                element = src as AutomationElement;
                if (element == null)
                {
                    name = "null";
                    return;
                }
                else
                {
                    name = element.GetCurrentPropertyValue(
                        AutomationElement.NameProperty) as string;
                }
            }
            catch
            {
                return;
            }

            try
            {
                #region

                if (name.Length == 0) name = "<NoName>";
                string str = string.Format("OnUninstallAppEvent:={0} : {1}", name, args.EventId.ProgrammaticName);
                Console.WriteLine(str);

                Thread.Sleep(5000);
                if (name.StartsWith("E'tricc Statistics Parser") || name.StartsWith("E'pia Framework") || name.StartsWith("E'tricc Shell"))
                {
                    Console.WriteLine("Name is ------------:" + name);
                    System.Windows.Automation.Condition c = new AndCondition(
                       new PropertyCondition(AutomationElement.NameProperty, "OK"),
                       new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                   );

                    // Find the element.
                    AutomationElement aeRun = element.FindFirst(TreeScope.Element | TreeScope.Descendants, c);
                    if (aeRun != null)
                    {
                        System.Windows.Point pt = AUIUtilities.GetElementCenterPoint(aeRun);
                        Input.MoveTo(pt);
                        Thread.Sleep(100);
                        Input.ClickAtPoint(pt);
                    }
                    else
                    {
                        Console.WriteLine("OK Button not Found ------------:" + name);
                        return;
                    }
                }
                else if (name.ToLower().Equals("user account control"))
                {   // will investigate later how to bypass this window
                    Console.WriteLine("----  user account control window, try to find Yes button");
                    // Remove Etricc Core
                    System.Windows.Automation.Condition cNt = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                        new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                        );

                    AutomationElement aeBtnYes = element.FindFirst(TreeScope.Element | TreeScope.Descendants, cNt);
                    Console.WriteLine("Yes button found: " + aeBtnYes.Current.Name);
                    Console.WriteLine("Yes button found... ---> Click Yes button");
                    System.Windows.Point pt = AUIUtilities.GetElementCenterPoint(aeBtnYes);
                    Input.MoveTo(pt);
                    Console.WriteLine("Input.MoveTo(aeBtnYes); " + aeBtnYes.Current.Name);
                    Thread.Sleep(2000);
                    Input.ClickAtPoint(pt);
                    //AUIUtilities.ClickElement(aeBtnNext);
                    Console.WriteLine("AUIUtilities.ClickElement(aeBtnYes): " + aeBtnYes.Current.Name);
                    Thread.Sleep(2000);
                }
                else if (name.ToLower().Equals("epia"))
                {
                    Console.WriteLine("try to find OK button");
                    // Remove Etricc Core
                    System.Windows.Automation.Condition cNt = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, "OK"),
                        new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                        );

                    AutomationElement aeBtnNext = element.FindFirst(TreeScope.Element | TreeScope.Descendants, cNt);
                    Console.WriteLine("OK button found: " + aeBtnNext.Current.Name);
                    Console.WriteLine("OK button found... ---> Click OK button");
                    AUIUtilities.ClickElement(aeBtnNext);
                }
                else if (element.Current.AutomationId.Equals("FrmRemoveRegistryKeysDialog"))
                {   // Remove Etricc Core
                    System.Windows.Automation.Condition cNt = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                        new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                        );

                    AutomationElement aeBtnNext = element.FindFirst(TreeScope.Element | TreeScope.Descendants, cNt);
                    Console.WriteLine("Yes button found: " + aeBtnNext.Current.Name);
                    Console.WriteLine("Yes button found... ---> Click Yes button");
                    AUIUtilities.ClickElement(aeBtnNext);
                }
                else if (name.StartsWith("E'tricc")
                        && name.IndexOf("Shell") < 0
                        && name.IndexOf("Playback") < 0
                        && name.IndexOf("Statistics") < 0
                        && name.IndexOf("HostTest") < 0)
                {
                    TransformPattern tranform =
                     element.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                    if (tranform != null)
                    {
                        tranform.Move(10, 10);
                    }
                }
                else if (name.Equals("<NoName>"))
                {
                    Console.WriteLine("Do nothing Name is ------------:" + name);
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine("OnUninstallAppEvent-Exception:" + ex.Message + "---"+ex.StackTrace);
                //throw;
            }

            Console.WriteLine("OnUninstallAppEvent-End");
        }
        #endregion
    }
}
