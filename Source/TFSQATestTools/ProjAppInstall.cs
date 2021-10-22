using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using System.Windows;
using System.Windows.Automation;
using System.IO;
using TestTools;

namespace TFSQATestTools
{
    public class ProjAppInstall
    {
        static public void StartAppMsiInstallExecution(string InstallerSource)
        {
            System.Diagnostics.Process Proc = new System.Diagnostics.Process();
            Proc.StartInfo.FileName = InstallerSource;
            Proc.StartInfo.CreateNoWindow = false;
            Proc.Start();
        }

        static public bool InstallApplication(string path, string App, EgeminApplication.SetupType setupType, ref string errorMsg, Logger logger, bool isDemo)
        {
            bool installed = true;
            string msiName = string.Empty;
            string msiNamePattern = string.Empty;
            if (App.Equals(EgeminApplication.ETRICC_SHELL))
                msiNamePattern = "Etricc Shell.msi";
            else if (App.Equals(EgeminApplication.EPIA))
                msiNamePattern = "Epia.msi";
            else if (App.Equals(EgeminApplication.EPIA_RESOURCEFILEEDITOR))
                msiNamePattern = EgeminApplication.EPIA_RESOURCEFILEEDITOR + ".msi";
            else if (App.Equals(EgeminApplication.ETRICC_HOSTTEST))
                msiNamePattern = "Etricc HostTest.msi";
            else if (App.Equals(EgeminApplication.ETRICC_PLAYBACK))
                msiNamePattern = "Etricc Playback.msi";
            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_PARSER))
                msiNamePattern = "Etricc.Statistics.Parser.msi";
            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR))
                msiNamePattern = "Etricc.Statistics.ParserConfigurator.msi";
            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_UI))
                msiNamePattern = "Etricc.Statistics.UI.msi";

            //find the msi in the filepath
            DirectoryInfo DirInfo = new DirectoryInfo(path);
            FileInfo[] files = DirInfo.GetFiles(msiNamePattern);
            try
            {
                if (files[0] != null)
                    msiName = files[0].Name;
            }
            catch (Exception ex)
            {
                installed = false;
                errorMsg = App + " Installation exception : " + ex.Message + " -- " + ex.StackTrace;
                System.Windows.Forms.MessageBox.Show(errorMsg + " --- \\n path is : " + path, "InstallApplication :" + App);
            }

            if (installed == true)
            {
                #region install application
                AutomationEventHandler UIInstallAppEventHandler = new AutomationEventHandler(ProjBasicEvent.OnInstallApplicationEvent);
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIInstallAppEventHandler);

                string filePath = System.IO.Path.Combine(path, msiName);
                Console.WriteLine("Start the MSI: " + App + "  --- at : "+filePath);
                if (logger != null) logger.LogMessageToFile("Start the MSI" + filePath, 0, 0);

                if (!System.IO.File.Exists(filePath))
                {
                    installed = false;
                    errorMsg = "---- msi file not exist in folder:" + filePath;
                    if (logger != null) logger.LogMessageToFile(errorMsg, 0, 0);
                    return installed;
                }

                StreamWriter write = File.CreateText(Path.Combine(@"C:\EtriccTests", "event.txt"));
                write.WriteLine("newversion");
                write.Close();

                Task t = Task.Factory.StartNew(() => { ProjAppInstall.StartAppMsiInstallExecution(filePath); });
                Thread.Sleep(5000);
                string versionInfo = string.Empty;
                try
                {
                    int kkk = 0;
                    int kcheckCnt  = 10;
                    if (isDemo)
                        kcheckCnt = 2;
                    while (kkk++ < kcheckCnt)
                    {
                        Thread.Sleep(4000);
                        try
                        {
                            StreamReader reader = File.OpenText(Path.Combine(@"C:\EtriccTests", "event.txt"));
                            versionInfo = reader.ReadLine();
                            reader.Close();

                            Console.WriteLine(kkk + "  version info: " + versionInfo);
                        }
                        catch (Exception ex)
                        {
                            Thread.Sleep(5000);
                            Console.WriteLine("Read event.txt exception: " + ex.Message);
                        }

                        if (versionInfo.StartsWith("anotherversion"))
                        {
                            try
                            {
                                Console.WriteLine("  Start " + App + " unInstallation -------------->");
                                ProjAppInstall.UninstallApplication(App, ref errorMsg);

                                StreamWriter write2 = File.CreateText(Path.Combine(@"C:\EtriccTests", "event.txt"));
                                write2.WriteLine("newversion");
                                write2.Close();
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("YYYYYYYYYYYYYYYYYYYY  exception " + ex.Message + " ----  "+ ex.StackTrace);
                            }
                            Console.WriteLine("Restart the MSI" + App);
                            t = Task.Factory.StartNew(() => { ProjAppInstall.StartAppMsiInstallExecution(filePath); });
                            Thread.Sleep(10000);
                            Console.WriteLine("Start " + App + " Installation -------------->");
                            Wait(5);
                            break;
                        }
                    }

                    DateTime installStartTime = DateTime.Now;
                    TimeSpan installTime = DateTime.Now - installStartTime;
                    int ix = 2;
                    bool hasFinishButton = false;
                    installed = true;
                    while (ProjAppInstall.InstallApplicationSetupByStep(App, setupType, ix, string.Empty, ref hasFinishButton, logger) == false)
                    {
                        if (hasFinishButton)
                        {
                            if (logger != null) logger.LogMessageToFile(" has fini <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  Start the MSI again:" + App, 0, 0);
                            Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  Start the MSI again:" + App);
                            hasFinishButton = false;
                            Console.WriteLine("Again Start the MSI" + App);
                            Task.Factory.StartNew(() => { ProjAppInstall.StartAppMsiInstallExecution(filePath); });
                            Thread.Sleep(10000);
                            Wait(5);
                        }
                        ix++;

                        installTime = DateTime.Now - installStartTime;
                        if (installTime.TotalMinutes > 30)   // if after half hour still not finished, install application can be consider as failed
                        {
                            errorMsg = "After half hour still not finished, install application can be consider as failed";
                            if (logger != null) logger.LogMessageToFile(" after half hour still not finished, install application can be consider as failed:", 0, 0);
                            ProcessUtilities.CloseProcess("MSIEXEC");
                            installed = false;
                            break;
                        }
                    }
                    
                }
                catch (Exception ex)
                {
                    installed = false;
                    errorMsg = ex.ToString() + System.Environment.NewLine + ex.StackTrace;
                    if (logger != null) logger.LogMessageToFile(App+ " InstallApplication exception::" + errorMsg, 0, 0);
                    //System.Windows.Forms.MessageBox.Show(errorMsg, App + "Install  sqq ");

                }
                finally
                {
                    //logger.LogMessageToFile(" finally<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  finally" + App, 0, 0);
                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                        AutomationElement.RootElement, UIInstallAppEventHandler);
                }
                #endregion End install application
            }
            return installed;
        }

        static public bool InstallApplication(string path, string App, EgeminApplication.SetupType setupType, ref string errorMsg, Logger logger)
        {
            bool installed = true;
            string msiName = string.Empty;
            string msiNamePattern = string.Empty;
            if (App.Equals(EgeminApplication.ETRICC_SHELL))
                msiNamePattern = "Etricc Shell.msi";
            else if (App.Equals(EgeminApplication.EPIA))
                msiNamePattern = "Epia.msi";
            else if (App.Equals(EgeminApplication.EPIA_RESOURCEFILEEDITOR))
                msiNamePattern = EgeminApplication.EPIA_RESOURCEFILEEDITOR + ".msi";
            else if (App.Equals(EgeminApplication.ETRICC_HOSTTEST))
                msiNamePattern = "Etricc HostTest.msi";
            else if (App.Equals(EgeminApplication.ETRICC_PLAYBACK))
                msiNamePattern = "Etricc Playback.msi";
            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_PARSER))
                msiNamePattern = "Etricc.Statistics.Parser.msi";
            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR))
                msiNamePattern = "Etricc.Statistics.ParserConfigurator.msi";
            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_UI))
                msiNamePattern = "Etricc.Statistics.UI.msi";

            //find the msi in the filepath
            DirectoryInfo DirInfo = new DirectoryInfo(path);
            FileInfo[] files = DirInfo.GetFiles(msiNamePattern);
            try
            {
                if (files[0] != null)
                    msiName = files[0].Name;
            }
            catch (Exception ex)
            {
                installed = false;
                errorMsg = App + " Installation exception : " + ex.Message + " -- " + ex.StackTrace;
                System.Windows.Forms.MessageBox.Show(errorMsg + " --- \\n path is : " + path, "InstallApplication :" + App);
            }

            if (installed == true)
            {
                #region install application
                AutomationEventHandler UIInstallAppEventHandler = new AutomationEventHandler(ProjBasicEvent.OnInstallApplicationEvent);
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIInstallAppEventHandler);

                string filePath = System.IO.Path.Combine(path, msiName);
                Console.WriteLine("Start the MSI: " + App + "  --- at : " + filePath);
                if (logger != null) logger.LogMessageToFile("Start the MSI" + filePath, 0, 0);

                if (!System.IO.File.Exists(filePath))
                {
                    installed = false;
                    errorMsg = "---- msi file not exist in folder:" + filePath;
                    if (logger != null) logger.LogMessageToFile(errorMsg, 0, 0);
                    return installed;
                }

                StreamWriter write = File.CreateText(Path.Combine(@"C:\EtriccTests", "event.txt"));
                write.WriteLine("newversion");
                write.Close();

                Task t = Task.Factory.StartNew(() => { ProjAppInstall.StartAppMsiInstallExecution(filePath); });
                Thread.Sleep(5000);
                string versionInfo = string.Empty;
                try
                {
                    int kkk = 0;
                    while (kkk++ < 10)
                    {
                        Thread.Sleep(4000);
                        try
                        {
                            StreamReader reader = File.OpenText(Path.Combine(@"C:\EtriccTests", "event.txt"));
                            versionInfo = reader.ReadLine();
                            reader.Close();

                            Console.WriteLine(kkk + "  version info: " + versionInfo);
                        }
                        catch (Exception ex)
                        {
                            Thread.Sleep(5000);
                            Console.WriteLine("Read event.txt exception: " + ex.Message);
                        }

                        if (versionInfo.StartsWith("anotherversion"))
                        {
                            try
                            {
                                Console.WriteLine("  Start " + App + " unInstallation -------------->");
                                ProjAppInstall.UninstallApplication(App, ref errorMsg);

                                StreamWriter write2 = File.CreateText(Path.Combine(@"C:\EtriccTests", "event.txt"));
                                write2.WriteLine("newversion");
                                write2.Close();
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("YYYYYYYYYYYYYYYYYYYY  exception " + ex.Message + " ----  " + ex.StackTrace);
                            }
                            Console.WriteLine("Restart the MSI" + App);
                            t = Task.Factory.StartNew(() => { ProjAppInstall.StartAppMsiInstallExecution(filePath); });
                            Thread.Sleep(10000);
                            Console.WriteLine("Start " + App + " Installation -------------->");
                            Wait(5);
                            break;
                        }
                    }

                    DateTime installStartTime = DateTime.Now;
                    TimeSpan installTime = DateTime.Now - installStartTime;
                    int ix = 2;
                    bool hasFinishButton = false;
                    installed = true;
                    while (ProjAppInstall.InstallApplicationSetupByStep(App, setupType, ix, string.Empty, ref hasFinishButton, logger) == false)
                    {
                        if (hasFinishButton)
                        {
                            if (logger != null) logger.LogMessageToFile(" has fini <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  Start the MSI again:" + App, 0, 0);
                            Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  Start the MSI again:" + App);
                            hasFinishButton = false;
                            Console.WriteLine("Again Start the MSI" + App);
                            Task.Factory.StartNew(() => { ProjAppInstall.StartAppMsiInstallExecution(filePath); });
                            Thread.Sleep(10000);
                            Wait(5);
                        }
                        ix++;

                        installTime = DateTime.Now - installStartTime;
                        if (installTime.TotalMinutes > 30)   // if after half hour still not finished, install application can be consider as failed
                        {
                            errorMsg = "After half hour still not finished, install application can be consider as failed";
                            if (logger != null) logger.LogMessageToFile(" after half hour still not finished, install application can be consider as failed:", 0, 0);
                            ProcessUtilities.CloseProcess("MSIEXEC");
                            installed = false;
                            break;
                        }
                    }

                }
                catch (Exception ex)
                {
                    installed = false;
                    errorMsg = ex.ToString() + System.Environment.NewLine + ex.StackTrace;
                    if (logger != null) logger.LogMessageToFile(App + " InstallApplication exception::" + errorMsg, 0, 0);
                    //System.Windows.Forms.MessageBox.Show(errorMsg, App + "Install  sqq ");

                }
                finally
                {
                    //logger.LogMessageToFile(" finally<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  finally" + App, 0, 0);
                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                        AutomationElement.RootElement, UIInstallAppEventHandler);
                }
                #endregion End install application
            }
            return installed;
        }

        static public bool InstallApplicationNet45(string path, string App, EgeminApplication.SetupType setupType, ref string errorMsg, Logger logger)
        {
            bool installed = true;
            string msiName = string.Empty;
            string msiNamePattern = string.Empty;
            if (App.Equals(EgeminApplication.ETRICC_SHELL))
                msiNamePattern = "Etricc Shell.msi";
            else if (App.Equals(EgeminApplication.EPIA))
                msiNamePattern = "Epia.msi";
            else if (App.Equals(EgeminApplication.EPIA_RESOURCEFILEEDITOR))
                msiNamePattern = EgeminApplication.EPIA_RESOURCEFILEEDITOR+".msi";
            else if (App.Equals(EgeminApplication.ETRICC_HOSTTEST))
                msiNamePattern = "Etricc HostTest.msi";
            else if (App.Equals(EgeminApplication.ETRICC_PLAYBACK))
                msiNamePattern = "Etricc Playback.msi";
            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_PARSER))
                msiNamePattern = "Etricc.Statistics.Parser.msi";
            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR))
                msiNamePattern = "Etricc.Statistics.ParserConfigurator.msi";
            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_UI))
                msiNamePattern = "Etricc.Statistics.UI.msi";

            //find the msi in the filepath
            DirectoryInfo DirInfo = new DirectoryInfo(path);
            FileInfo[] files = DirInfo.GetFiles(msiNamePattern);
            try
            {
                if (files[0] != null)
                    msiName = files[0].Name;
            }
            catch (Exception ex)
            {
                installed = false;
                errorMsg = App + " Installation exception : " + ex.Message + " -- " + ex.StackTrace;
                System.Windows.Forms.MessageBox.Show(errorMsg, "InstallApplication :" + path);
            }

            if (installed == true)
            {
                #region install application
                AutomationEventHandler UIInstallAppEventHandler = new AutomationEventHandler(ProjBasicEvent.OnInstallApplicationEvent);
                // Add Open window Event Handler
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    AutomationElement.RootElement, TreeScope.Descendants, UIInstallAppEventHandler);

                Console.WriteLine("Start the MSI: " + App);
                string filePath = System.IO.Path.Combine(path, msiName);
                if (logger != null) logger.LogMessageToFile("Start the MSI" + filePath, 0, 0);

                if (!System.IO.File.Exists(filePath))
                {
                    installed = false;
                    errorMsg = "---- msi file not exist in folder:" + filePath;
                    if (logger != null) logger.LogMessageToFile(errorMsg, 0, 0);
                    return installed;
                }

                Task.Factory.StartNew(() => { ProjAppInstall.StartAppMsiInstallExecution(filePath); });
                Thread.Sleep(5000);
                try
                {
                    int ix = 2;
                    bool hasFinishButton = false;
                    while (ProjAppInstall.InstallApplicationSetupByStepNet45(App, setupType, ix, string.Empty, ref hasFinishButton, logger) == false)
                    {
                        if (hasFinishButton)
                        {
                            if (logger != null) logger.LogMessageToFile(" has fini <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  Start the MSI again:" + App, 0, 0);
                            Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  Start the MSI again:" + App);
                            hasFinishButton = false;
                            Console.WriteLine("Again Start the MSI" + App);
                            Task.Factory.StartNew(() => { ProjAppInstall.StartAppMsiInstallExecution(path); });
                            Thread.Sleep(10000);
                            Wait(5);
                        }
                        ix++;
                    }
                    installed = true;
                }
                catch (Exception ex)
                {
                    installed = false;
                    errorMsg = ex.ToString() + System.Environment.NewLine + ex.StackTrace;
                    if (logger != null) logger.LogMessageToFile(App + " InstallApplication exception::" + errorMsg, 0, 0);
                    //System.Windows.Forms.MessageBox.Show(errorMsg, App + "Install  sqq ");

                }
                finally
                {
                    //logger.LogMessageToFile(" finally<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  finally" + App, 0, 0);
                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                        AutomationElement.RootElement, UIInstallAppEventHandler);
                }
                #endregion End install application
            }
            return installed;
        }

        static public bool InstallApplicationSetupByStepNet45(string App, EgeminApplication.SetupType setupType, int step, string errorMsg, ref bool hasFinishButton, Logger logger)
        {
            bool clickCloseButton = false;
            AutomationElement aeMsiDialogCloseClassWindow = null;
            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
            AutomationElement aeClickButton = null;
            Point ClickButtonPt = new Point(0, 0);

            DateTime StartTime = DateTime.Now;
            TimeSpan Time = DateTime.Now - StartTime;
            if (logger != null) logger.LogMessageToFile("<-----> Start Install : " + App + "  Step by step, :" + step, 0, 0);
            Console.WriteLine("<-----> Start Install : " + App + "  Step by step, :" + step);
            try
            {
                // //find install application Window
                while (aeMsiDialogCloseClassWindow == null && Time.TotalMinutes <= 2)
                {
                    if (logger != null) logger.LogMessageToFile("<-----> Install : " + App + "  Step by step, step: " + step, 0, 0);
                    Console.WriteLine("<-----> Install : " + App + "  Step by step, step: " + step);
                    try
                    {
                        aeMsiDialogCloseClassWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                        if (aeMsiDialogCloseClassWindow != null)
                        {
                            Console.WriteLine("<-----> aeMsiDialogCloseClassWindow != null: " );
                            WindowPattern wp = aeMsiDialogCloseClassWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                            if (wp.Current.WindowVisualState == WindowVisualState.Minimized)
                            {
                                //System.Windows.Forms.MessageBox.Show("wp.Current.WindowVisualState == WindowVisualState.Minimized", stepMsg);
                                wp.SetWindowVisualState(WindowVisualState.Normal);
                                Thread.Sleep(1000);
                            }

                            #region // process Text   Setup type
                            if (setupType != EgeminApplication.SetupType.Default)
                            {
                                if (logger != null) logger.LogMessageToFile("<-----> find MsiDialogCloseClassWindow name is:" + aeMsiDialogCloseClassWindow.Current.Name, 0, 0);
                                //find all install Window screen texts
                                System.Windows.Automation.Condition cText = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                                AutomationElementCollection aeAllTexts = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Children, cText);
                                Thread.Sleep(1000);
                                for (int i = 0; i < aeAllTexts.Count; i++)
                                {
                                    if (logger != null) logger.LogMessageToFile("<----->MsiDialogCloseClassWindow text(" + i + ")  : " + aeAllTexts[i].Current.Name, 0, 0);
                                    if (aeAllTexts[i].Current.Name.StartsWith("Setup Type"))
                                    {
                                        switch (App)
                                        {
                                            case EgeminApplication.EPIA:
                                                #region
                                                AutomationElement aeCustomRadioButton
                                                    = AUIUtilities.FindElementByName("Custom", aeMsiDialogCloseClassWindow);
                                                if (aeCustomRadioButton != null)
                                                {
                                                    SelectionItemPattern sPattern = aeCustomRadioButton.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                                    sPattern.Select();
                                                    Thread.Sleep(2000);
                                                }
                                            #endregion
                                                break;
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region // process Text   "Custom Setup"
                            if (setupType != EgeminApplication.SetupType.Default)
                            {
                                if (logger != null) logger.LogMessageToFile("<-----> find MsiDialogCloseClassWindow name is:" + aeMsiDialogCloseClassWindow.Current.Name, 0, 0);
                                //find all install Window screen texts
                                System.Windows.Automation.Condition cText = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                                AutomationElementCollection aeAllTexts = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Children, cText);
                                Thread.Sleep(1000);
                                for (int i = 0; i < aeAllTexts.Count; i++)
                                {
                                    if (logger != null) logger.LogMessageToFile("<----->MsiDialogCloseClassWindow text(" + i + ")  : " + aeAllTexts[i].Current.Name, 0, 0);
                                    if (aeAllTexts[i].Current.Name.StartsWith("Custom Setup"))
                                    {
                                        switch (App)
                                        {
                                            case EgeminApplication.EPIA:
                                                #region
                                                string typeName = "Epia Server";
                                                if (setupType == EgeminApplication.SetupType.EpiaServerOnly)
                                                    typeName = "Epia Shell";
                                                else if (setupType == EgeminApplication.SetupType.EpiaShellOnly)
                                                    typeName = "Epia Server";
                    
                                                System.Windows.Automation.Condition cTreeItem = new AndCondition(
                                                        new PropertyCondition(AutomationElement.NameProperty, typeName),
                                                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem)
                                                    );

                                                AutomationElement aeType = aeMsiDialogCloseClassWindow.FindFirst(TreeScope.Descendants, cTreeItem);
                                                if (aeType != null)
                                                {
                                                    double left = aeType.Current.BoundingRectangle.Left;
                                                    double Ymid = (aeType.Current.BoundingRectangle.Top + aeType.Current.BoundingRectangle.Bottom) / 2;
                                                    Point typeCheckBoxPoint = new Point(left - 10, Ymid);
                                                    Input.MoveToAndClick(typeCheckBoxPoint);
                                                    Thread.Sleep(3000);

                                                    Console.WriteLine("---------------------------aeMsiDialogCloseClassWindow.Current.IsEnabled=" + aeMsiDialogCloseClassWindow.Current.IsEnabled);
                                                    System.Windows.Automation.Condition cMenu = new AndCondition(
                                                            new PropertyCondition(AutomationElement.NameProperty, "Menu"),
                                                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Menu)
                                                        );

                                                    AutomationElement aeMenu = AutomationElement.RootElement.FindFirst(TreeScope.Children, cMenu);
                                                    if (aeMenu == null)
                                                    {
                                                        errorMsg = "Find aeMenu failed:" + "Meuu";
                                                        Console.WriteLine(errorMsg);
                                                    }
                                                    else
                                                    {
                                                        AutomationElement aeMenuItem = ProjBasicUI.GetMenuItemFromElement(aeMenu, "Item 4", "id", 120, ref errorMsg);
                                                        if (aeMenuItem == null)
                                                        {
                                                            errorMsg = aeMenuItem.Current.Name + " aeMenuItem not found -->  " + "Item 4";
                                                            Console.WriteLine(errorMsg);
                                                        }
                                                        else
                                                        {
                                                            Input.MoveToAndClick(aeMenuItem);
                                                            Thread.Sleep(3000);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    errorMsg = aeType.Current.Name + " aeType not found -->  " + typeName;
                                                    Console.WriteLine(errorMsg);
                                                    i = aeAllTexts.Count;
                                                }
                                                #endregion
                                                break;
                                        }
                                    }
                                }
                            }
                            #endregion
                            
                            System.Windows.Automation.Condition cButton = new AndCondition(
                                new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                                new PropertyCondition(AutomationElement.IsContentElementProperty, true),
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                            );

                            AutomationElementCollection aeButtons = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Descendants, cButton);
                            if ( aeButtons != null)
                                Console.WriteLine("--------------------------- aeButtons.Count=" + aeButtons.Count);
                            Thread.Sleep(1000);
                            bool clickButtonFound = false;
                            for (int i = 0; i < aeButtons.Count; i++)
                            {
                                Console.WriteLine("---------------------------aeButtons[i]=" + aeButtons[i].Current.Name);
                                //System.Windows.Forms.MessageBox.Show("aeButtons[i]=" + aeButtons[i].Current.Name);
                                if (logger != null) logger.LogMessageToFile("<----->MsiDialogCloseClassWindow enabled button(" + i + ")  : " + aeButtons[i].Current.Name, 0, 0);
                                if (aeButtons[i].Current.Name.StartsWith("Next"))
                                {
                                    aeClickButton = aeButtons[i];
                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    clickButtonFound = true;
                                    break;
                                }
                                else if (aeButtons[i].Current.Name.StartsWith("Install"))
                                {
                                    aeClickButton = aeButtons[i];
                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    clickButtonFound = true;
                                    break;
                                }
                                else if (aeButtons[i].Current.Name.StartsWith("OK"))
                                {
                                    aeClickButton = aeButtons[i];
                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    clickButtonFound = true;
                                    break;
                                }
                                else if (aeButtons[i].Current.Name.StartsWith("Finish"))
                                {
                                    // if launcher checkbox exist, first uncheck it. 
                                    System.Windows.Automation.Condition cText = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                                    AutomationElementCollection aeAllTexts = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Children, cText);
                                    Thread.Sleep(1000);
                                    for (int ix = 0; ix < aeAllTexts.Count; ix++)
                                    {
                                        if (logger != null) logger.LogMessageToFile("<----->MsiDialogCloseClassWindow text(" + ix + ")  : " + aeAllTexts[ix].Current.Name, 0, 0);
                                        Console.WriteLine("--------------------------- (aeAllTexts[i].Current.Name=" + aeAllTexts[i].Current.Name);
                                        //Thread.Sleep(5000);
                                        if (aeAllTexts[i].Current.Name.StartsWith("The InstallShield Wizard has successfully installed"))
                                        {
                                            switch (App)
                                            {
                                                case EgeminApplication.EPIA_RESOURCEFILEEDITOR:
                                                    #region
                                                    AutomationElement aeLauncherCheckBox
                                                        = AUIUtilities.FindElementByType(ControlType.CheckBox, aeMsiDialogCloseClassWindow);
                                                    if (aeLauncherCheckBox != null)
                                                    {
                                                        TogglePattern sPattern = aeLauncherCheckBox.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                                        ToggleState tgState = sPattern.Current.ToggleState;
                                                        if (tgState == ToggleState.On)
                                                            sPattern.Toggle();
                                                        Thread.Sleep(2000);
                                                    }
                                                    #endregion
                                                    break;
                                            }
                                        }
                                    }

                                    aeClickButton = aeButtons[i];
                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    clickButtonFound = true;
                                    clickCloseButton = true;
                                    break;
                                }
                                else if (aeButtons[i].Current.Name.StartsWith("Close"))   // --> wait until Close button
                                {
                                    hasFinishButton = true;
                                    #region process Finish
                                    aeClickButton = aeButtons[i];
                                    Point FinishButtonPoint = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    Console.WriteLine("Click Button is <Finish>");
                                    // find radio button
                                    // Set a property condition that will be used to find the control.
                                    System.Windows.Automation.Condition c2 = new PropertyCondition(
                                        AutomationElement.ControlTypeProperty, ControlType.RadioButton);

                                    AutomationElementCollection aeAllRadioButtons = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Element | TreeScope.Descendants, c2);
                                    Thread.Sleep(1000);
                                    foreach (AutomationElement s in aeAllRadioButtons)
                                    {
                                        if (s.Current.Name.StartsWith("Remove"))
                                        {
                                            SelectionItemPattern itemRadioPattern = s.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                            itemRadioPattern.Select();
                                            Thread.Sleep(2000);
                                        }
                                    }

                                    Input.MoveToAndClick(FinishButtonPoint);
                                    Thread.Sleep(2000);
                                    //Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
                                    bool hasMsiDialogCloseClassWindow = true;
                                    DateTime xStartTime = DateTime.Now;
                                    TimeSpan xTime = DateTime.Now - xStartTime;
                                    while (hasMsiDialogCloseClassWindow && xTime.TotalSeconds <= 600)
                                    {
                                        aeMsiDialogCloseClassWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                        if (aeMsiDialogCloseClassWindow == null)
                                        {
                                            hasMsiDialogCloseClassWindow = false;
                                            Console.WriteLine("Remove " + App + " finished");
                                        }
                                        else
                                        {
                                            aeButtons = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Children, cButton);
                                            Thread.Sleep(1000);
                                            for (int ib = 0; ib < aeButtons.Count; ib++)
                                            {
                                                if (aeButtons[ib].Current.Name.StartsWith("Close"))
                                                {
                                                    Console.WriteLine("Close button displayed " + App + " ......");
                                                    clickButtonFound = true;
                                                    aeClickButton = aeButtons[ib];
                                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                                    hasMsiDialogCloseClassWindow = false;
                                                    break;
                                                }
                                            }

                                            if (clickButtonFound)     // Close button found
                                                Console.WriteLine("Stop Continue Remove " + App + " ......");
                                            else
                                            {
                                                Console.WriteLine("Continue Remove " + App + " ......");
                                                hasMsiDialogCloseClassWindow = true;
                                                Thread.Sleep(2000);
                                                xTime = DateTime.Now - xStartTime;
                                            }
                                        }
                                    }
                                    #endregion end process Finish
                                    break;
                                }
                            }

                            if (clickButtonFound)
                            {
                                if (logger != null) logger.LogMessageToFile("<---> clickButtonFound... " + aeClickButton.Current.Name, 0, 0);
                                //Point OptionPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                Input.MoveTo(ClickButtonPt);
                                Thread.Sleep(1000);
                                if (logger != null) logger.LogMessageToFile("<---> " + aeClickButton.Current.Name + " button clicking ... ", 0, 0);
                                Input.ClickAtPoint(ClickButtonPt);
                                Thread.Sleep(500);
                            }
                            else
                                Thread.Sleep(2000);
                        }
                        else
                        {
                            errorMsg = "Error: install window not found:" + App;
                            Console.WriteLine(errorMsg);
                        }
                    }
                    catch (System.Windows.Automation.ElementNotAvailableException ex)
                    {
                        string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                        if (logger != null) logger.LogMessageToFile("<---> " + msg, 0, 0);
                        aeMsiDialogCloseClassWindow = null;
                        Console.WriteLine("++++++  "+msg);
                    }

                    Thread.Sleep(2000);
                    Time = DateTime.Now - StartTime;
                }
            }
            catch (System.Windows.Automation.ElementNotAvailableException ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                if (logger != null) logger.LogMessageToFile("<---> " + msg, 0, 0);
                clickCloseButton = false;
                Console.WriteLine("------  "+msg);
            }

            return clickCloseButton;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="stepMsg"></param>
        /// <param name="App">Epia4, EtriccShell</param>
        /// <param name="logger"></param>
        /// <returns></returns>
        static public bool InstallApplicationSetupByStep(string App, EgeminApplication.SetupType setupType, int step, string errorMsg, ref bool hasFinishButton, Logger logger)
        {
            bool clickCloseButton = false;
            AutomationElement aeMsiDialogCloseClassWindow = null;
            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
            AutomationElement aeClickButton = null;
            Point ClickButtonPt = new Point(0, 0);

            DateTime StartTime = DateTime.Now;
            TimeSpan Time = DateTime.Now - StartTime;
            if (logger != null) logger.LogMessageToFile("<-----> Start Install : " + App + "  Step by step, :" + step, 0, 0);
            Console.WriteLine("<-----> Start Install : " + App + "  Step by step, :" + step);
            try
            {
                // //find install application Window
                while (aeMsiDialogCloseClassWindow == null && Time.TotalMinutes <= 2)
                {
                    if (logger != null) logger.LogMessageToFile("<-----> Install : " + App + "  Step by step, step: " + step, 0, 0);
                    try
                    {
                        aeMsiDialogCloseClassWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                        if (aeMsiDialogCloseClassWindow != null)
                        {
                            WindowPattern wp = aeMsiDialogCloseClassWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                            if (wp.Current.WindowVisualState == WindowVisualState.Minimized)
                            {
                                //System.Windows.Forms.MessageBox.Show("wp.Current.WindowVisualState == WindowVisualState.Minimized", stepMsg);
                                wp.SetWindowVisualState(WindowVisualState.Normal);
                                Thread.Sleep(1000);
                            }

                            #region // process Text
                            if (logger != null) logger.LogMessageToFile("<-----> find MsiDialogCloseClassWindow name is:" + aeMsiDialogCloseClassWindow.Current.Name, 0, 0);
                            //find all install Window screen texts
                            System.Windows.Automation.Condition cText = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                            AutomationElementCollection aeAllTexts = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Children, cText);
                            Thread.Sleep(1000);
                            for (int i = 0; i < aeAllTexts.Count; i++)
                            {
                                if (logger != null) logger.LogMessageToFile("<----->MsiDialogCloseClassWindow text(" + i + ")  : " + aeAllTexts[i].Current.Name, 0, 0);
                                if (aeAllTexts[i].Current.Name.StartsWith("Components"))
                                {
                                    switch (App)
                                    {
                                        case EgeminApplication.EPIA:
                                            #region
                                            if (setupType == EgeminApplication.SetupType.Default || setupType == EgeminApplication.SetupType.EpiaServerOnly)
                                            {
                                                AutomationElement aeIAgreeRadioButton
                                                    = AUIUtilities.FindElementByName("E'pia Server", aeMsiDialogCloseClassWindow);
                                                TogglePattern tg = aeIAgreeRadioButton.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                                ToggleState tgState = tg.Current.ToggleState;
                                                if (tgState == ToggleState.Off)
                                                    tg.Toggle();
                                            }

                                            if (setupType == EgeminApplication.SetupType.Default || setupType == EgeminApplication.SetupType.EpiaShellOnly)
                                            {
                                                AutomationElement aeShellckb
                                                    = AUIUtilities.FindElementByName("E'pia Shell", aeMsiDialogCloseClassWindow);
                                                TogglePattern tg2 = aeShellckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                                ToggleState tg2State = tg2.Current.ToggleState;
                                                if (tg2State == ToggleState.Off)
                                                    tg2.Toggle();

                                                if (logger != null) logger.LogMessageToFile("<---> Components Sever and shell checkbox state set on  ... ", 0, 0);
                                                Thread.Sleep(3000);
                                            }
                                            #endregion
                                            break;
                                        case EgeminApplication.ETRICC_SHELL:
                                            #region
                                            AutomationElement aeServerckb
                                                = AUIUtilities.FindElementByName("E'pia Server Extensions (Resource & Security Files)", aeMsiDialogCloseClassWindow);
                                            TogglePattern tgShell = aeServerckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                            ToggleState tgShellState = tgShell.Current.ToggleState;
                                            if (tgShellState == ToggleState.Off)
                                                tgShell.Toggle();

                                            AutomationElement aeEtriccShellckb
                                                = AUIUtilities.FindElementByName("E'pia Shell Extensions (Shell Module & Config)", aeMsiDialogCloseClassWindow);
                                            TogglePattern tgEtricc = aeEtriccShellckb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                            ToggleState tg2ShellState = tgEtricc.Current.ToggleState;
                                            if (tg2ShellState == ToggleState.Off)
                                                tgEtricc.Toggle();

                                            AutomationElement aeEtricccCorekb
                                                = AUIUtilities.FindElementByName("E'tricc Core Extensions (Wrappers)", aeMsiDialogCloseClassWindow);
                                            TogglePattern tg3 = aeEtricccCorekb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                            ToggleState tg3State = tg3.Current.ToggleState;
                                            if (tg3State == ToggleState.Off)
                                                tg3.Toggle();

                                            if (logger != null) logger.LogMessageToFile("<---> Components Sever and shell checkbox state set on  ... ", 0, 0);
                                            Thread.Sleep(3000);
                                            #endregion
                                            break;
                                        case EgeminApplication.ETRICC_STATISTICS_UI:
                                            #region
                                            AutomationElement aeServerExtentionCkb
                                                = AUIUtilities.FindElementByName("E'pia Server Extensions (Resource & Security Files && Server Components)", aeMsiDialogCloseClassWindow);
                                            TogglePattern tgServerExtentionCkb = aeServerExtentionCkb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                            ToggleState tgServerExtentionCkbState = tgServerExtentionCkb.Current.ToggleState;
                                            if (tgServerExtentionCkbState == ToggleState.Off)
                                                tgServerExtentionCkb.Toggle();

                                            AutomationElement aeShellExtentionCkb
                                                = AUIUtilities.FindElementByName("E'pia Shell Extensions (Shell Module & Config)", aeMsiDialogCloseClassWindow);
                                            TogglePattern tgShellExtentionCkb = aeShellExtentionCkb.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                            ToggleState tgShellExtentionCkbState = tgShellExtentionCkb.Current.ToggleState;
                                            if (tgShellExtentionCkbState == ToggleState.Off)
                                                tgShellExtentionCkb.Toggle();

                                            if (logger != null) logger.LogMessageToFile("<---> Components Sever and shell extention checkbox state set on  ... ", 0, 0);
                                            Thread.Sleep(3000);
                                            #endregion
                                            break;
                                    }
                                }
                            }
                            #endregion

                            System.Windows.Automation.Condition cButton = new AndCondition(
                                new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                                new PropertyCondition(AutomationElement.IsContentElementProperty, true),
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                            );

                            AutomationElementCollection aeButtons = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Children, cButton);
                            Thread.Sleep(1000);
                            bool clickButtonFound = false;
                            for (int i = 0; i < aeButtons.Count; i++)
                            {
                                if (logger != null) logger.LogMessageToFile("<----->MsiDialogCloseClassWindow enabled button(" + i + ")  : " + aeButtons[i].Current.Name, 0, 0);
                                if (aeButtons[i].Current.Name.StartsWith("Next"))
                                {
                                    aeClickButton = aeButtons[i];
                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    clickButtonFound = true;
                                    break;
                                }
                                else if (aeButtons[i].Current.Name.StartsWith("Close"))
                                {
                                    aeClickButton = aeButtons[i];
                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    clickButtonFound = true;
                                    clickCloseButton = true;
                                    break;
                                }
                                else if (aeButtons[i].Current.Name.StartsWith("Finish"))   // --> wait until Close button
                                {
                                    hasFinishButton = true;
                                    #region process Finish
                                    aeClickButton = aeButtons[i];
                                    Point FinishButtonPoint = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    Console.WriteLine("Click Button is <Finish>");
                                    // find radio button
                                    // Set a property condition that will be used to find the control.
                                    System.Windows.Automation.Condition c2 = new PropertyCondition(
                                        AutomationElement.ControlTypeProperty, ControlType.RadioButton);

                                    AutomationElementCollection aeAllRadioButtons = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Element | TreeScope.Descendants, c2);
                                    Thread.Sleep(1000);
                                    foreach (AutomationElement s in aeAllRadioButtons)
                                    {
                                        if (s.Current.Name.StartsWith("Remove"))
                                        {
                                            SelectionItemPattern itemRadioPattern = s.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                            itemRadioPattern.Select();
                                            Thread.Sleep(2000);
                                        }
                                    }

                                    Input.MoveToAndClick(FinishButtonPoint);
                                    Thread.Sleep(2000);
                                    //Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
                                    bool hasMsiDialogCloseClassWindow = true;
                                    DateTime xStartTime = DateTime.Now;
                                    TimeSpan xTime = DateTime.Now - xStartTime;
                                    while (hasMsiDialogCloseClassWindow && xTime.TotalSeconds <= 600)
                                    {
                                        aeMsiDialogCloseClassWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                        if (aeMsiDialogCloseClassWindow == null)
                                        {
                                            hasMsiDialogCloseClassWindow = false;
                                            Console.WriteLine("Remove " + App + " finished");
                                        }
                                        else
                                        {
                                            aeButtons = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Children, cButton);
                                            Thread.Sleep(1000);
                                            for (int ib = 0; ib < aeButtons.Count; ib++)
                                            {
                                                if (aeButtons[ib].Current.Name.StartsWith("Close"))
                                                {
                                                    Console.WriteLine("Close button displayed " + App + " ......");
                                                    clickButtonFound = true;
                                                    aeClickButton = aeButtons[ib];
                                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                                    hasMsiDialogCloseClassWindow = false;
                                                    break;
                                                }
                                            }

                                            if (clickButtonFound)     // Close button found
                                                Console.WriteLine("Stop Continue Remove " + App + " ......");
                                            else
                                            {
                                                Console.WriteLine("Continue Remove " + App + " ......");
                                                hasMsiDialogCloseClassWindow = true;
                                                Thread.Sleep(2000);
                                                xTime = DateTime.Now - xStartTime;
                                            }
                                        }
                                    }
                                    #endregion end process Finish
                                    break;
                                }
                            }

                            // aeMsiDialogCloseClassWindow.Current.IsEnabled == false maybe install error screen displayed
                            if (aeMsiDialogCloseClassWindow.Current.IsEnabled == false)
                            {
                                Thread.Sleep(15000);
                                if (logger != null) logger.LogMessageToFile("<----->  : aeMsiDialogCloseClassWindow.Current.IsEnabled == false", 0, 0);
                                Console.WriteLine("<----->  : aeMsiDialogCloseClassWindow.Current.IsEnabled == false");
                                return false;
                            }

                            if (clickButtonFound)
                            {
                                if (logger != null) logger.LogMessageToFile("<---> clickButtonFound... " + aeClickButton.Current.Name, 0, 0);
                                //Point OptionPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                Input.MoveTo(ClickButtonPt);
                                Thread.Sleep(1000);
                                if (logger != null) logger.LogMessageToFile("<---> " + aeClickButton.Current.Name + " button clicking ... ", 0, 0);
                                Input.ClickAtPoint(ClickButtonPt);
                                Thread.Sleep(500);
                            }
                            else
                                Thread.Sleep(2000);
                        }
                        else
                        {
                            errorMsg = "Error: install window not found:" + App;
                            if (logger != null) logger.LogMessageToFile("<---> " + errorMsg, 0, 0);
                            Thread.Sleep(2000);
                        }
                    }
                    catch (System.Windows.Automation.ElementNotAvailableException ex)
                    {
                        string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                        if (logger != null) logger.LogMessageToFile("<---> " + msg, 0, 0);
                        aeMsiDialogCloseClassWindow = null;
                    }

                    Thread.Sleep(2000);
                    Time = DateTime.Now - StartTime;
                }
            }
            catch (System.Windows.Automation.ElementNotAvailableException ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                if (logger != null) logger.LogMessageToFile("<---> " + msg, 0, 0);
                clickCloseButton = false;
            }

            return clickCloseButton;
        }

        public static bool InstallEtriccCoreSetup(string FilePath, string sEtriccMsiName, ref string errorMsg,  Logger logger,
            bool mInstallEtriccLauncher, bool mInstallOldEtricc5Service)
        {
            logger.LogMessageToFile(EgeminApplication.ETRICC + "::: start installation : " + FilePath, 0, 0);
            bool installed = true;
            string App = EgeminApplication.ETRICC;
            AutomationEventHandler UIInstallAppEventHandler = new AutomationEventHandler(ProjBasicEvent.OnInstallApplicationEvent);
            //find the msi in the filepath
            string msiName = string.Empty;

            try
            {
                DirectoryInfo DirInfo = new DirectoryInfo(FilePath);
                FileInfo[] files = DirInfo.GetFiles(sEtriccMsiName);
                if (files[0] != null)
                    msiName = files[0].Name;
            }
            catch (Exception ex)
            {
                installed = false;
                errorMsg = App + " Installation exception : " + ex.Message + " -- " + ex.StackTrace;
                System.Windows.Forms.MessageBox.Show("exception: " + ex.Message + " -- " + ex.StackTrace, "InstallEtricc5Setup2 xwe");
            }

            if (installed == true)
            {
                try
                {
                    #region install application
                    Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                        AutomationElement.RootElement, TreeScope.Descendants, UIInstallAppEventHandler);
                    Console.WriteLine("start the MSI" + App);
                    Thread.Sleep(5000);
                    Console.WriteLine("Start " + App + " Installation -------------->" + msiName);
                    string path = System.IO.Path.Combine(FilePath, msiName);
                    Task.Factory.StartNew(() => { ProjAppInstall.StartAppMsiInstallExecution(path); });
                    Thread.Sleep(10000);

                    int ix = 2;
                    bool hasFinishButton = false;
                    while (ProjAppInstall.InstallEtriccCoreSetupByStep(EgeminApplication.ETRICC, mInstallEtriccLauncher, true, mInstallOldEtricc5Service, ix++, string.Empty, ref hasFinishButton, logger) == false)
                    {
                        Thread.Sleep(3000);
                        if (hasFinishButton)
                        {
                            logger.LogMessageToFile(" has fini <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  Start the MSI again:" + App, 0, 0);
                            Console.WriteLine("<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  Start the MSI again:" + App);
                            hasFinishButton = false;
                            Console.WriteLine("Again Start the MSI" + App);
                            Task.Factory.StartNew(() => { ProjAppInstall.StartAppMsiInstallExecution(path); });
                            Thread.Sleep(10000);
                        }
                        ix++;
                    }

                    installed = true;
                    #endregion End install Etricc Core
                }
                catch (Exception ex)
                {
                    installed = false;
                    errorMsg = ex.ToString() + System.Environment.NewLine + ex.StackTrace;
                    System.Windows.Forms.MessageBox.Show(errorMsg, App + "InstallSetup2 sqq ");
                }
                finally
                {
                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                        AutomationElement.RootElement, UIInstallAppEventHandler);
                }
            }
            return installed;
        }

        static public bool InstallEtriccCoreSetupByStep(string App, bool isLauncher, bool isLogView, bool isService, int step, string errorMsg, ref bool hasFinishButton, Logger logger)
        {
            bool clickCloseButton = false;
            AutomationElement aeMsiDialogCloseClassWindow = null;
            System.Windows.Automation.Condition condition 
                = new System.Windows.Automation.PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
            AutomationElement aeClickButton = null;
            Point ClickButtonPt = new Point(0, 0);

            DateTime StartTime = DateTime.Now;
            TimeSpan Time = DateTime.Now - StartTime;
            logger.LogMessageToFile("<-----> Start Install : " + App + "  Step by step, :" + step, 0, 0);
            try
            {
                // //find install application Window
                while (aeMsiDialogCloseClassWindow == null && Time.TotalMinutes <= 2)
                {
                    logger.LogMessageToFile("<-----> Install : " + App + "  Step by step, step: " + step, 0, 0);
                    try
                    {
                        aeMsiDialogCloseClassWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                        if (aeMsiDialogCloseClassWindow != null)
                        {
                            WindowPattern wp = aeMsiDialogCloseClassWindow.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                            if (wp.Current.WindowVisualState == WindowVisualState.Minimized)
                            {
                                //System.Windows.Forms.MessageBox.Show("wp.Current.WindowVisualState == WindowVisualState.Minimized", stepMsg);
                                wp.SetWindowVisualState(WindowVisualState.Normal);
                                Thread.Sleep(1000);
                            }

                            #region // process Text
                            logger.LogMessageToFile("<-----> find MsiDialogCloseClassWindow name is:" + aeMsiDialogCloseClassWindow.Current.Name, 0, 0);
                            //find all install Window screen texts
                            System.Windows.Automation.Condition cText =
                                new System.Windows.Automation.PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text);
                            AutomationElementCollection aeAllTexts = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Children, cText);
                            Thread.Sleep(1000);
                            for (int i = 0; i < aeAllTexts.Count; i++)
                            {
                                logger.LogMessageToFile("<----->MsiDialogCloseClassWindow text(" + i + ")  : " + aeAllTexts[i].Current.Name, 0, 0);
                                if (aeAllTexts[i].Current.Name.StartsWith("License Agreement"))
                                {
                                    AutomationElement aeBtnAgree = AUIUtilities.GetElementByNameProperty(aeMsiDialogCloseClassWindow, "I Agree");
                                    Point pt = AUIUtilities.GetElementCenterPoint(aeBtnAgree);
                                    Input.MoveTo(pt);
                                    Thread.Sleep(500);
                                    AUIUtilities.ClickElement(aeBtnAgree);
                                }
                                else if (aeAllTexts[i].Current.Name.StartsWith("Select components to install."))
                                {
                                    #region select components
                                    AutomationElement aeEtriccLauncher
                                            = AUIUtilities.FindElementByName("Launcher", aeMsiDialogCloseClassWindow);
                                    TogglePattern tgpLaun = aeEtriccLauncher.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                    ToggleState tgLauncher = tgpLaun.Current.ToggleState;
                                    if (isLauncher)
                                    {
                                        if (tgLauncher == ToggleState.Off)
                                            tgpLaun.Toggle();
                                    }
                                    else
                                    {
                                        if (tgLauncher == ToggleState.On)
                                            tgpLaun.Toggle();
                                    }

                                    AutomationElement aeEtriccExplorer
                                                = AUIUtilities.FindElementByName("Explorer", aeMsiDialogCloseClassWindow);
                                    TogglePattern tgpExp = aeEtriccExplorer.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                    ToggleState tgExplorer = tgpExp.Current.ToggleState;
                                    if (tgExplorer == ToggleState.Off)
                                        tgpExp.Toggle();

                                    AutomationElement aeLogView
                                            = AUIUtilities.FindElementByName("Log Tools", aeMsiDialogCloseClassWindow);
                                    TogglePattern tgpLog = aeLogView.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                    ToggleState tgLogView = tgpLog.Current.ToggleState;
                                    if (isLogView)
                                    {
                                        if (tgLogView == ToggleState.Off)
                                            tgpLog.Toggle();
                                    }
                                    else
                                    {
                                        if (tgLogView == ToggleState.On)
                                            tgpLog.Toggle();
                                    }

                                    AutomationElement aeService
                                            = AUIUtilities.FindElementByName("Service", aeMsiDialogCloseClassWindow);
                                    TogglePattern tgpServ = aeService.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                                    ToggleState tgService = tgpServ.Current.ToggleState;
                                    if (isService)
                                    {
                                        if (tgService == ToggleState.Off)
                                            tgpServ.Toggle();
                                    }
                                    else
                                    {
                                        if (tgService == ToggleState.On)
                                            tgpServ.Toggle();
                                    }
                                    #endregion
                                }
                                else if (aeAllTexts[i].Current.Name.StartsWith("Installing E'tricc"))
                                {
                                    #region installing
                                    if (isService || isLauncher)
                                    {
                                        TransformPattern tranform = aeMsiDialogCloseClassWindow.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                                        tranform.Move(10, 10);

                                        System.Windows.Automation.Condition conditionFrm = new PropertyCondition(AutomationElement.AutomationIdProperty, "FrmLauncherFunctionality");
                                        System.Windows.Automation.Condition conditionFrm2 = new PropertyCondition(AutomationElement.AutomationIdProperty, "FrmSecurityFunctionality");
                                        AutomationElement frmElement = null;
                                        if (isLauncher)
                                        {
                                            Console.WriteLine("try to find  FrmLauncherFunctionality window...");
                                            while (frmElement == null)
                                            {
                                                Thread.Sleep(1000);
                                                Console.WriteLine("Wait until Next button found...");
                                                frmElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, conditionFrm);
                                                if (frmElement == null)
                                                    Console.WriteLine("Frm Window  not found");
                                                else
                                                {
                                                    System.Windows.Automation.Condition cNt = new AndCondition(
                                                    new PropertyCondition(AutomationElement.NameProperty, "Next >"),
                                                    new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                                                    );

                                                    AutomationElement aeBtnNext = frmElement.FindFirst(TreeScope.Element | TreeScope.Descendants, cNt);
                                                    Console.WriteLine("Next button found: " + aeBtnNext.Current.Name);
                                                    Console.WriteLine("Nex button found... ---> Click NExt button");
                                                    AUIUtilities.ClickElement(aeBtnNext);
                                                }
                                            }
                                        }

                                        if (isService)
                                        {
                                            AutomationElement frmElement2 = null;
                                            while (frmElement2 == null)
                                            {
                                                Thread.Sleep(1000);
                                                Console.WriteLine("Wait until Next button found...");
                                                frmElement2 = AutomationElement.RootElement.FindFirst(TreeScope.Children, conditionFrm2);
                                                if (frmElement2 == null)
                                                    Console.WriteLine("Frm2 Window  not found");
                                                else
                                                {
                                                    System.Windows.Automation.Condition cNt = new AndCondition(
                                                    new PropertyCondition(AutomationElement.NameProperty, "Next >"),
                                                    new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                                                    );

                                                    AutomationElement aeBtnNext = frmElement2.FindFirst(TreeScope.Element | TreeScope.Descendants, cNt);
                                                    Console.WriteLine("Next button found: " + aeBtnNext.Current.Name);
                                                    Console.WriteLine("Nex button found... ---> Click NExt button");
                                                    AUIUtilities.ClickElement(aeBtnNext);
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                }
                            }
                            #endregion

                            System.Windows.Automation.Condition cButton = new AndCondition(
                                new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                                new PropertyCondition(AutomationElement.IsContentElementProperty, true),
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                            );

                            AutomationElementCollection aeButtons = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Children, cButton);
                            Thread.Sleep(1000);
                            bool clickButtonFound = false;
                            for (int i = 0; i < aeButtons.Count; i++)
                            {
                                logger.LogMessageToFile("<----->MsiDialogCloseClassWindow enabled button(" + i + ")  : " + aeButtons[i].Current.Name, 0, 0);
                                if (aeButtons[i].Current.Name.StartsWith("Next"))
                                {
                                    aeClickButton = aeButtons[i];
                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    clickButtonFound = true;
                                    break;
                                }
                                else if (aeButtons[i].Current.Name.StartsWith("Close"))
                                {
                                    aeClickButton = aeButtons[i];
                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    clickButtonFound = true;
                                    clickCloseButton = true;
                                    break;
                                }
                                else if (aeButtons[i].Current.Name.StartsWith("Finish"))   // --> wait until Close button
                                {
                                    hasFinishButton = true;
                                    #region process Finish
                                    aeClickButton = aeButtons[i];
                                    Point FinishButtonPoint = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                    Console.WriteLine("Click Button is <Finish>");
                                    // find radio button
                                    // Set a property condition that will be used to find the control.
                                    System.Windows.Automation.Condition c2 = new PropertyCondition(
                                        AutomationElement.ControlTypeProperty, ControlType.RadioButton);

                                    AutomationElementCollection aeAllRadioButtons = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Element | TreeScope.Descendants, c2);
                                    Thread.Sleep(1000);
                                    foreach (AutomationElement s in aeAllRadioButtons)
                                    {
                                        if (s.Current.Name.StartsWith("Remove"))
                                        {
                                            SelectionItemPattern itemRadioPattern = s.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                                            itemRadioPattern.Select();
                                            Thread.Sleep(2000);
                                        }
                                    }

                                    Input.MoveToAndClick(FinishButtonPoint);
                                    Thread.Sleep(2000);

                                    //TransformPattern tranform = aeMsiDialogCloseClassWindow.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                                    //tranform.Move(10, 10);
                                    System.Windows.Automation.Condition conditionFrmRemove 
                                        = new PropertyCondition(AutomationElement.AutomationIdProperty, "FrmRemoveRegistryKeysDialog");
                                    AutomationElement frmElement2 = null;
                                    while (frmElement2 == null)
                                    {
                                        Thread.Sleep(1000);
                                        Console.WriteLine("Wait until yes button found...");
                                        frmElement2 = AutomationElement.RootElement.FindFirst(TreeScope.Children, conditionFrmRemove);
                                        if (frmElement2 == null)
                                            Console.WriteLine("FrmRemoveRegistryKeysDialog Window  not found");
                                        else
                                        {
                                            System.Windows.Automation.Condition cNt = new AndCondition(
                                            new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                                            new PropertyCondition(AutomationElement.IsContentElementProperty, true)
                                            );

                                            AutomationElement aeBtnNext = frmElement2.FindFirst(TreeScope.Element | TreeScope.Descendants, cNt);
                                            Console.WriteLine("Yes button found: " + aeBtnNext.Current.Name);
                                            Console.WriteLine("Yes button found... ---> Click Yes button");
                                            AUIUtilities.ClickElement(aeBtnNext);
                                        }
                                    }

                                    //Condition condition = new PropertyCondition(AutomationElement.ClassNameProperty, "MsiDialogCloseClass");
                                    bool hasMsiDialogCloseClassWindow = true;
                                    DateTime xStartTime = DateTime.Now;
                                    TimeSpan xTime = DateTime.Now - xStartTime;
                                    while (hasMsiDialogCloseClassWindow && xTime.TotalSeconds <= 600)
                                    {
                                        aeMsiDialogCloseClassWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                        if (aeMsiDialogCloseClassWindow == null)
                                        {
                                            hasMsiDialogCloseClassWindow = false;
                                            Console.WriteLine("Remove " + App + " finished");
                                        }
                                        else
                                        {
                                            aeButtons = aeMsiDialogCloseClassWindow.FindAll(TreeScope.Children, cButton);
                                            Thread.Sleep(1000);
                                            for (int ib = 0; ib < aeButtons.Count; ib++)
                                            {
                                                if (aeButtons[ib].Current.Name.StartsWith("Close"))
                                                {
                                                    Console.WriteLine("Close button displayed " + App + " ......");
                                                    clickButtonFound = true;
                                                    aeClickButton = aeButtons[ib];
                                                    ClickButtonPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                                    hasMsiDialogCloseClassWindow = false;
                                                    break;
                                                }
                                            }

                                            if (clickButtonFound)     // Close button found
                                                Console.WriteLine("Stop Continue Remove " + App + " ......");
                                            else
                                            {
                                                Console.WriteLine("Continue Remove " + App + " ......");
                                                hasMsiDialogCloseClassWindow = true;
                                                Thread.Sleep(2000);
                                                xTime = DateTime.Now - xStartTime;
                                            }
                                        }
                                    }
                                    #endregion end process Finish
                                    break;
                                }
                            }

                            if (clickButtonFound)
                            {
                                logger.LogMessageToFile("<---> clickButtonFound... " + aeClickButton.Current.Name, 0, 0);
                                //Point OptionPt = AUIUtilities.GetElementCenterPoint(aeClickButton);
                                Input.MoveTo(ClickButtonPt);
                                Thread.Sleep(1000);
                                logger.LogMessageToFile("<---> " + aeClickButton.Current.Name + " button clicking ... ", 0, 0);
                                Input.ClickAtPoint(ClickButtonPt);
                                Thread.Sleep(500);
                            }
                            else
                                Thread.Sleep(2000);
                        }
                        else
                        {
                            errorMsg = "Error: install window not found:" + App;
                            logger.LogMessageToFile("<---> " + errorMsg, 0, 0);
                            Thread.Sleep(10000);
                        }
                    }
                    catch (System.Windows.Automation.ElementNotAvailableException ex)
                    {
                        string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                        logger.LogMessageToFile("Exception 1........  <---> " + msg, 0, 0);
                        aeMsiDialogCloseClassWindow = null;
                    }
                    Time = DateTime.Now - StartTime;
                }
            }
            catch (System.Windows.Automation.ElementNotAvailableException ex)
            {
                string msg = ex.ToString() + "----" + System.Environment.NewLine + ex.StackTrace;
                logger.LogMessageToFile("Exception 1........ <---> " + msg, 0, 0);
                clickCloseButton = false;
            }

            return clickCloseButton;
        }

        //   Uninstallapplication 
        //
        //
        static public bool UninstallApplication(string App, ref string errorMsg)
        {
            bool uninstallOK = true;
            AutomationElement aeApp = null;
         
            string uninstallWindowName = "Programs and Features";

            #region // open Programs and Features Window
            string uninstallWindowNameAddOrRemoveProgram = "Add or Remove Programs";    // XP
            string uninstallWindowNameProgramsAndFeatures = "Programs and Features";    //
            string uninstallWindowNameControlPanelPrograms = "Control Panel\\Programs\\Programs and Features";

            System.OperatingSystem os = System.Environment.OSVersion;
            int OSVersionMajor = os.Version.Major;
            if (OSVersionMajor >= 6)
            {
                uninstallWindowName = uninstallWindowNameProgramsAndFeatures;
            }
            else
                uninstallWindowName = uninstallWindowNameAddOrRemoveProgram;

            // start uninstall programs feature windows 
            // uninstall windows names are different for different platform
            Console.WriteLine("<> Start uninstall programs feature window:" + App + "   and machine: " + System.Environment.MachineName.ToUpper());
            Task.Factory.StartNew(() => { StartAppMsiInstallExecution(@"C:\Windows\System32\appwiz.cpl"); });
            Thread.Sleep(5000);
            #endregion
            Console.WriteLine(App+ " uninstall ==> now Searching for Programs and Features main window:" + uninstallWindowName);
            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName);
            System.Windows.Automation.Condition condition1 = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowNameAddOrRemoveProgram);
            System.Windows.Automation.Condition condition2 = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowNameProgramsAndFeatures);
            System.Windows.Automation.Condition condition3 = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowNameControlPanelPrograms);
            AutomationElement aeProgramsFeaturesScreen = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
            int kx = 0;
            string CloseWindowNameCondition = uninstallWindowName;
            while (aeProgramsFeaturesScreen == null && kx++ < 20 )
            {
                Thread.Sleep(5000);
                if (kx % 3 == 0)
                {
                    CloseWindowNameCondition = uninstallWindowNameAddOrRemoveProgram;
                    aeProgramsFeaturesScreen = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition1);
                    Console.WriteLine(App + "   " + kx + "de Searching for Programs and Features main window:" + "Add or Remove Programs");
                }
                else if (kx % 3 == 1)
                {
                    CloseWindowNameCondition = uninstallWindowNameProgramsAndFeatures;
                    aeProgramsFeaturesScreen = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition2);
                    Console.WriteLine(App + "   " + kx + "de Searching for Programs and Features main window:" + "Programs and Features");
                }
                else if (kx % 3 == 2)
                {
                    CloseWindowNameCondition = uninstallWindowNameControlPanelPrograms;
                    aeProgramsFeaturesScreen = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition3);
                    Console.WriteLine(App + "   " + kx + "de Searching for Programs and Features main window:" + "Control Panel\\Programs\\Programs and Features");
                }
            }

            if (aeProgramsFeaturesScreen != null)
            {
                aeProgramsFeaturesScreen.SetFocus();
                WindowPattern windowPattern = aeProgramsFeaturesScreen.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                if (windowPattern != null)
                {
                    // Make sure our window is usable.
                    // WaitForInputIdle will return before the specified time if the window is ready.
                    while (windowPattern.WaitForInputIdle(100) == false)
                    {
                        Console.WriteLine("HasKeyboardFocus false");
                        Console.WriteLine("windowPattern.WaitForInputIdle(100):" + windowPattern.WaitForInputIdle(100));
                        Thread.Sleep(100);
                    }
                }
                else
                { 
                    uninstallOK = false;
                    errorMsg = "aeProgramsFeaturesScreen windowPattern == null";
                    Console.WriteLine(errorMsg);
                }

                Console.WriteLine("Programs and Features  Window ready for user interaction");
                Console.WriteLine("Searching programs item button...");
                if (uninstallOK == true)
                {
                    AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(aeProgramsFeaturesScreen, "Folder View"); //  "文件夹视图"
                    if (aeGridView != null)
                    {
                        bool appOffScreen = true;
                        while (appOffScreen == true)
                        {
                            Console.WriteLine("Gridview found...");
                            
                            #region find App menu item
                            int itemsCnt = 0;
                            // Set a property condition that will be used to find the control.
                            AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children,
                                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem));
                            Console.WriteLine("Programs count ..." + aeProgram.Count);

                            while (aeProgram.Count > itemsCnt )
                            {
                                itemsCnt = aeProgram.Count;
                                Console.WriteLine("itemsCnt ..." + itemsCnt);
                                Thread.Sleep(5000);
                                aeProgram = aeGridView.FindAll(TreeScope.Children,
                                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem));
                                Console.WriteLine("Programs count ..." + aeProgram.Count);
                            }
                            for (int i = 0; i < aeProgram.Count; i++)
                            {
                                Console.WriteLine("programs name: " + aeProgram[i].Current.Name);
                                if (App.Equals(EgeminApplication.EPIA))
                                {
                                    if (aeProgram[i].Current.Name.StartsWith("E'pia Fr") 
                                        || ( aeProgram[i].Current.Name.StartsWith("Epia") && aeProgram[i].Current.Name.IndexOf("Epia.ResourceFileEditor") < 0 ))
                                        aeApp = aeProgram[i];
                                }
                                else if (App.Equals(EgeminApplication.EPIA_RESOURCEFILEEDITOR))
                                {
                                    if (aeProgram[i].Current.Name.StartsWith("E'pia Resource File Editor ") 
                                        || ( aeProgram[i].Current.Name.StartsWith("Epia.ResourceFileEditor") ) )
                                        aeApp = aeProgram[i];
                                }
                                else if (App.Equals(EgeminApplication.ETRICC))
                                {
                                    if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                            && aeProgram[i].Current.Name.IndexOf("Statistics") < 0
                                            && aeProgram[i].Current.Name.IndexOf("Shell") < 0
                                            && aeProgram[i].Current.Name.IndexOf("Playback") < 0
                                            && aeProgram[i].Current.Name.IndexOf("HostTest") < 0)
                                        aeApp = aeProgram[i];
                                }
                                else if (App.Equals(EgeminApplication.ETRICC_SHELL))
                                {
                                    if (aeProgram[i].Current.Name.StartsWith("E'tricc Shell "))
                                        aeApp = aeProgram[i];
                                }
                                else if (App.Equals(EgeminApplication.ETRICC_PLAYBACK))
                                {
                                    if (aeProgram[i].Current.Name.StartsWith("E'tricc Playback "))
                                        aeApp = aeProgram[i];
                                }
                                else if (App.Equals(EgeminApplication.ETRICC_HOSTTEST))
                                {
                                    if (aeProgram[i].Current.Name.StartsWith("E'tricc HostTest "))
                                        aeApp = aeProgram[i];
                                }
                                else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_PARSER))
                                {
                                    if (aeProgram[i].Current.Name.StartsWith("E'tricc Statistics Parser "))
                                        aeApp = aeProgram[i];
                                }
                                else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR))
                                {
                                    if (aeProgram[i].Current.Name.StartsWith("E'tricc Statistics ParserConfigurator "))
                                        aeApp = aeProgram[i];
                                }
                                else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_UI))
                                {
                                    if (aeProgram[i].Current.Name.StartsWith("E'tricc Statistics UI "))
                                        aeApp = aeProgram[i];
                                }
                                else if (App.Equals(EgeminApplication.AUTOMATICTESTING))
                                {
                                    if (aeProgram[i].Current.Name.StartsWith(EgeminApplication.AUTOMATICTESTING))
                                        aeApp = aeProgram[i];
                                }
                            }
                            #endregion

                            if (aeApp == null)
                            {
                                appOffScreen = false;
                            }
                            else
                            {
                                if (aeApp.Current.IsOffscreen == true)
                                {
                                    Console.WriteLine("aeApp.Current.IsOffscreen == true ");
                                    Thread.Sleep(2000);
                                    //ScrollPattern scrollPattern = GetScrollPattern(element);
                                    ScrollPattern scrollPattern = (ScrollPattern)aeGridView.GetCurrentPattern(ScrollPattern.Pattern);
                                    if (scrollPattern.Current.VerticallyScrollable)
                                    {
                                        scrollPattern.ScrollVertical(ScrollAmount.LargeIncrement);
                                    }
                                    Thread.Sleep(2000);
                                }
                                else
                                {
                                    appOffScreen = false;
                                }
                            }
                        }
                    }
                    else
                    {
                        uninstallOK = false;
                        errorMsg = "Gridview NOT found...";
                        Console.WriteLine(errorMsg);
                    }
                }

                if (uninstallOK == true)
                {
                    if (aeApp == null)
                    {
                        Console.WriteLine("No " + App + " name: ");
                        AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(aeProgramsFeaturesScreen, "Close");
                        if (btnClose != null)
                        {
                            AUIUtilities.ClickElement(btnClose);
                        }
                    }
                    else
                    {
                        AutomationEventHandler UIUninstallAppEventHandler = new AutomationEventHandler(ProjBasicEvent.OnUninstallAppEvent);
                        // Add Open window Event Handler
                        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                            AutomationElement.RootElement, TreeScope.Descendants, UIUninstallAppEventHandler);

                        Console.WriteLine(App + " name: " + aeApp.Current.Name);
                        Input.MoveTo(aeApp);
                        string x = aeApp.Current.Name;
                        Thread.Sleep(2000);
                        // click on aeApp item 
                        InvokePattern pattern = aeApp.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        pattern.Invoke();
                        Thread.Sleep(2000);
                        #region Program and Features dialog
                        // find Program and Features dialog (in the future, do not show me this dialog box possible)
                        System.Windows.Automation.Condition cWindows = new AndCondition(
                            new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName),
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                            );
                        AutomationElement dialogElement = aeProgramsFeaturesScreen.FindFirst(TreeScope.Children, cWindows);
                        if (dialogElement != null)
                        {
                            Thread.Sleep(3000);
                            AUIUtilities.MoveUIElement(dialogElement, 0, 0);
                            Thread.Sleep(3000);
                            AutomationElement btnYes = AUIUtilities.GetElementByNameProperty(aeProgramsFeaturesScreen, "Yes");
                            if (btnYes != null)
                            {
                                AUIUtilities.ClickElement(btnYes);
                            }
                        }
                        #endregion

                        #region // Window Installer section
                        // wait until application uninstalled
                        DateTime startTime = DateTime.Now;
                        TimeSpan mTime = DateTime.Now - startTime;
                        bool hasApplication = IsApplicationInstalled(App, uninstallWindowName);
                        while (hasApplication == true && mTime.TotalSeconds < 360)
                        {
                            Thread.Sleep(8000);
                            mTime = DateTime.Now - startTime;
                            if (mTime.TotalMilliseconds > 300000)
                            {
                                System.Windows.Forms.MessageBox.Show(App + " Uninstall run timeout (sec) " + mTime.TotalSeconds);
                                break;
                            }
                            hasApplication = IsApplicationInstalled(App, uninstallWindowName);
                        }
                        #endregion
                        // close Features and Programs window    
                        Console.WriteLine("close Features and Programs window----------------------");
                        System.Windows.Automation.Condition conditionClose = new PropertyCondition(AutomationElement.NameProperty, CloseWindowNameCondition);
                        aeProgramsFeaturesScreen = AutomationElement.RootElement.FindFirst(TreeScope.Children, conditionClose);
                        Thread.Sleep(2000);
                        AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(aeProgramsFeaturesScreen, "Close");
                        if (btnClose != null)
                            AUIUtilities.ClickElement(btnClose);

                        Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                AutomationElement.RootElement, UIUninstallAppEventHandler);
                    }
                }    
            }
            else
            {
                uninstallOK = false;
                errorMsg = "Programs And Features Window not opened after 2 min";
                Console.WriteLine(errorMsg);
            }

            if (uninstallOK == true)
                 Console.WriteLine(App + " ---------- Uninstalled ---------- OK");
            else
                 Console.WriteLine(App + " ---------- Uninstalled ---------- FAILED");

            return uninstallOK;
        }

        public static bool IsApplicationInstalled(string ApplicationType, string uninstallWindowName)
        {
            bool applicationInstalled = false;

            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName);
           
            ////------------------
            #region // open Programs and Features Window
            string uninstallWindowNameAddOrRemoveProgram = "Add or Remove Programs";    // XP
            string uninstallWindowNameProgramsAndFeatures = "Programs and Features";    //
            string uninstallWindowNameControlPanelPrograms = "Control Panel\\Programs\\Programs and Features";
            #endregion
            Console.WriteLine(ApplicationType + " installed ?  ==> now Searching for Programs and Features main window:" + uninstallWindowName);
            System.Windows.Automation.Condition condition1 = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowNameAddOrRemoveProgram);
            System.Windows.Automation.Condition condition2 = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowNameProgramsAndFeatures);
            System.Windows.Automation.Condition condition3 = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowNameControlPanelPrograms);
            AutomationElement appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
            int kx = 0;
            while (appElement == null && kx++ < 20)
            {
                Thread.Sleep(5000);
                if (kx % 3 == 0)
                {
                    appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition1);
                    Console.WriteLine(ApplicationType + "   " + kx + "de Searching for Programs and Features main window:" + "Add or Remove Programs");
                }
                else if (kx % 3 == 1)
                {
                    appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition2);
                    Console.WriteLine(ApplicationType + "   " + kx + "de Searching for Programs and Features main window:" + "Programs and Features");
                }
                else if (kx % 3 == 2)
                {
                    appElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition3);
                    Console.WriteLine(ApplicationType + "   " + kx + "de Searching for Programs and Features main window:" + "Control Panel\\Programs\\Programs and Features");
                }
            }
            //-------------------
            if (appElement != null)
            {   // (1) Programs and Features main window
                Console.WriteLine("Programs and Features main window opend");
                Wait(1);
                Console.WriteLine("Searching programs item button...");
                AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(appElement, "Folder View");
                if (aeGridView != null)
                    Console.WriteLine("Gridview found...");
                Wait(1);
                // Set a property condition that will be used to find the control.
                System.Windows.Automation.Condition c = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.DataItem);

                AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children, c);
                Console.WriteLine("Programs count ..." + aeProgram.Count);
                for (int i = 0; i < aeProgram.Count; i++)
                {
                    switch (ApplicationType)
                    {
                        case EgeminApplication.EPIA:
                            if (aeProgram[i].Current.Name.StartsWith("E'pia Fr")
                                 || (aeProgram[i].Current.Name.StartsWith("Epia") && aeProgram[i].Current.Name.IndexOf("Epia.ResourceFileEditor") < 0))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case EgeminApplication.EPIA_RESOURCEFILEEDITOR:
                            if (aeProgram[i].Current.Name.StartsWith("E'pia Resource File Editor ")
                                || (aeProgram[i].Current.Name.StartsWith("Epia.ResourceFileEditor"))  )
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case EgeminApplication.ETRICC:
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                    && aeProgram[i].Current.Name.IndexOf("Shell") < 0
                                 && aeProgram[i].Current.Name.IndexOf("Playback") < 0
                                   && aeProgram[i].Current.Name.IndexOf("Statistics") < 0
                                && aeProgram[i].Current.Name.IndexOf("HostTest") < 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case EgeminApplication.ETRICC_SHELL:
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc Shell "))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case EgeminApplication.ETRICC_HOSTTEST:
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc HostTest "))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case EgeminApplication.ETRICC_PLAYBACK:
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc Playback "))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case EgeminApplication.ETRICC_STATISTICS_PARSER:
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc Statistics Parser "))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR:
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc Statistics ParserConfigurator "))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case EgeminApplication.ETRICC_STATISTICS_UI:
                            if (aeProgram[i].Current.Name.StartsWith("E'tricc Statistics UI "))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case "Ewcs":
                            if (aeProgram[i].Current.Name.StartsWith("E'wcs")
                                && aeProgram[i].Current.Name.IndexOf("TestPrograms") < 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case "EwcsTestProgram":
                            if (aeProgram[i].Current.Name.StartsWith("E'wcs")
                                && aeProgram[i].Current.Name.IndexOf("TestPrograms") > 0)
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                        case "AutomaticTesting":
                            if (aeProgram[i].Current.Name.StartsWith("Automatic"))
                            {
                                applicationInstalled = true;
                                Console.WriteLine(aeProgram[i].Current.Name + " is installed");
                                Wait(5);
                                break;
                            }
                            break;
                    }
                }
            }

            return applicationInstalled;
        }

        public static void Wait(int second)
        {
            Thread.Sleep(second * 1000);
        }

        // Uninstall applications in Windows XP 
        /// <summary>
        ///             uninstall application in Windows XP
        /// </summary>
        /// <param name="App"></param>
        /// <param name="errorMsg"></param>
        /// <returns></returns>
        static public bool UninstallApplicationXP(string App, ref string errorMsg)
        {
            bool uninstallOK = true;
            AutomationElement aeProgramsFeaturesScreen = null;
            AutomationElement aeApp = null;
            string MachineName = System.Environment.MachineName;
            string uninstallWindowName = "Add or Remove Programs";

            #region // open Programs and Features Window
            // start uninstall programs feature windowa
            Console.WriteLine("Start uninstall programs feature window:" + App);
            Task.Factory.StartNew(() => { StartAppMsiInstallExecution(@"C:\Windows\System32\appwiz.cpl"); });
            Thread.Sleep(5000);
            #endregion

            Console.WriteLine("--------------  Searching for Programs and Features main window:" + uninstallWindowName);
            Thread.Sleep(2000);
            AutomationElement aeRemoveBtn = null;
            Console.WriteLine("---------------  SearchingApplicationInGridViewXP:" + uninstallWindowName);
            
            if (SearchingApplicationInGridViewXP(uninstallWindowName, App, ref aeApp, ref aeProgramsFeaturesScreen, ref errorMsg) == false)
            {
                uninstallOK = false;
                errorMsg = "SearchingApplicationInGridViewXP failed:" + errorMsg;
                Console.WriteLine(errorMsg);
            }
            else
            {
                if (aeApp == null)
                {
                    Console.WriteLine("No " + App + " name: ");
                    AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(aeProgramsFeaturesScreen, "Close");
                    if (btnClose != null)
                    {
                        AUIUtilities.ClickElement(btnClose);
                        Thread.Sleep(2000);
                        return true;
                    }
                }
                else
                {
                    Console.WriteLine("Yes  " + App + " name found -------------------------------------: " + aeApp.Current.Name);
                    Point appPt = AUIUtilities.GetElementCenterPoint(aeApp);
                    Input.MoveToAndClick(appPt);
                    //Input.MoveToAndDoubleClick(TestTools.AUIUtilities.GetElementCenterPoint(aeApp));
                    Thread.Sleep(3000);
                    // find main form again
                    aeApp = null;
                    if (SearchingApplicationInGridViewXP(uninstallWindowName, App, ref aeApp, ref aeProgramsFeaturesScreen, ref errorMsg) == false)
                    {
                        uninstallOK = false;
                        errorMsg = "SearchingApplicationInGridViewXP failed:" + errorMsg;
                        Console.WriteLine(errorMsg);
                    }
                    else
                    {
                        if (aeApp == null)
                        {
                            Console.WriteLine("2 No " + App + " name: ");
                            AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(aeProgramsFeaturesScreen, "Close");
                            if (btnClose != null)
                            {
                                AUIUtilities.ClickElement(btnClose);
                                Thread.Sleep(2000);
                                return true;
                            }
                        }
                        else
                        {
                            Console.WriteLine("================================== Yes  " + App + " name: ");
                            AutomationElementCollection aeBtns= aeApp.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
                            if (aeBtns.Count > 0)
                            {
                                for (int i = 0; i < aeBtns.Count; i++)
                                {
                                    Console.WriteLine("aeBtns[i].Current.Name : " + aeBtns[i].Current.Name);
                                    if ( aeBtns[i].Current.Name.Equals("Remove"))
                                    {
                                        aeRemoveBtn = aeBtns[i];
                                        break;
                                    }
                                }
                            }
                            
                            
                            //aeRemoveBtn = aeApp.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Remove"));
                            if (aeRemoveBtn != null)
                            {
                                //AutomationEventHandler UIUninstallAppEventHandler = new AutomationEventHandler(ProjBasicEvent.OnUninstallAppEvent);
                                // Add Open window Event Handler
                                //Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                //    AutomationElement.RootElement, TreeScope.Descendants, UIUninstallAppEventHandler);
                                Input.MoveToAndClick(TestTools.AUIUtilities.GetElementCenterPoint(aeRemoveBtn));
                                Thread.Sleep(2000);
                                // click on aeApp item 
                                //InvokePattern pattern = aeApp.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                                // pattern.Invoke();
                                Thread.Sleep(2000);
                                #region Program and Features dialog
                                // find Program and Features dialog (in the future, do not show me this dialog box possible)
                                System.Windows.Automation.Condition cWindows = new AndCondition(
                                    new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName),
                                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
                                    );
                                AutomationElement dialogElement = aeProgramsFeaturesScreen.FindFirst(TreeScope.Children, cWindows);
                                if (dialogElement != null)
                                {
                                    Thread.Sleep(3000);
                                    AUIUtilities.MoveUIElement(dialogElement, 0, 0);
                                    Thread.Sleep(3000);
                                    AutomationElement btnYes = AUIUtilities.GetElementByNameProperty(aeProgramsFeaturesScreen, "Yes");
                                    if (btnYes != null)
                                    {
                                        AUIUtilities.ClickElement(btnYes);
                                    }
                                }
                                #endregion

                                #region // Window Installer section
                                // wait until application uninstalled
                                DateTime startTime = DateTime.Now;
                                TimeSpan mTime = DateTime.Now - startTime;
                                bool hasApplication = IsApplicationInstalledXP(App, uninstallWindowName);
                                while (hasApplication == true && mTime.TotalSeconds < 360)
                                {
                                    Console.WriteLine("xxxxxxxxxxxxxxxxxx sApplicationInstalledXP---------------------aeProgramsFeaturesScreen.Current.IsEnable:" + aeProgramsFeaturesScreen.Current.IsEnabled);
                                    Thread.Sleep(8000);
                                    mTime = DateTime.Now - startTime;
                                    if (mTime.TotalMilliseconds > 300000)
                                    {
                                        System.Windows.Forms.MessageBox.Show(App + " Uninstall run timeout (se.c) " + mTime.TotalSeconds);
                                        break;
                                    }
                                    hasApplication = IsApplicationInstalledXP(App, uninstallWindowName);
                                    Console.WriteLine("has application " + App + " -----> " + hasApplication);
                                }
                                #endregion
                                // close Features and Programs window    
                                aeProgramsFeaturesScreen.SetFocus();
                                Console.WriteLine("close Features and Programs window----------------------");
                                //System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName);
                                //aeProgramsFeaturesScreen = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                                aeProgramsFeaturesScreen = ProjBasicUI.GetMainWindowByNameWithinTime(uninstallWindowName, 10);
                                if (aeProgramsFeaturesScreen != null)
                                {
                                    System.Windows.Automation.Condition cButtonClose = new AndCondition(
                                           new PropertyCondition(AutomationElement.IsEnabledProperty, true),
                                           new PropertyCondition(AutomationElement.IsContentElementProperty, false),
                                            new PropertyCondition(AutomationElement.NameProperty, "Close"),
                                           new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                                       );

                                    AutomationElement btnClose = aeProgramsFeaturesScreen.FindFirst(TreeScope.Descendants, cButtonClose);
                                    //AutomationElement btnClose = AUIUtilities.GetElementByNameProperty(aeProgramsFeaturesScreen, "Close");
                                    Console.WriteLine("btnClose----------------------" + btnClose.Current.Name);
                                    if (btnClose != null)
                                    {
                                        AUIUtilities.ClickElement(btnClose);
                                        Thread.Sleep(2000);
                                    }
                                }
                                //Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                                //        AutomationElement.RootElement, UIUninstallAppEventHandler);
                            }
                            else
                            {
                                uninstallOK = false;
                                errorMsg = "Remove button not found:";
                                Console.WriteLine(errorMsg);
                                Thread.Sleep(2000);
                            }
                        }
                    }
                }
            }

            if (uninstallOK == true)
                Console.WriteLine(App + " ---------- Uninstalled ---------- OK");
            else
                Console.WriteLine(App + " ---------- Uninstalled ---------- FAILED");

            //TestTools.ProcessUtilities.CloseProcess("rundll32");
            return uninstallOK;
        }

        public static bool IsApplicationInstalledXP(string ApplicationType, string uninstallWindowName)
        {
            bool applicationInstalled = false;
            AutomationElement aeProgramsFeaturesScreen = null;
            AutomationElement aeApp = null;
            string errorMsg = string.Empty;

            if (SearchingApplicationInGridViewXP(uninstallWindowName, ApplicationType, ref aeApp, ref aeProgramsFeaturesScreen, ref errorMsg))
            {
                Console.WriteLine( " ---------- ApplicationType ---------- " + ApplicationType);
                if (aeApp != null)
                    applicationInstalled = true;

            }
         
            return applicationInstalled;
        }

        public static bool SearchingApplicationInGridViewXP(string uninstallWindowName, string App, ref AutomationElement aeApp, 
            ref AutomationElement aeProgramsFeaturesScreen, ref string errorMsg)
        { 
            bool statusOK = true;
            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.NameProperty, uninstallWindowName);
            aeProgramsFeaturesScreen = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
            if (aeProgramsFeaturesScreen != null)
            {
                /*aeProgramsFeaturesScreen.SetFocus();
                while (ProjBasicUI.GetThisWindowIsTopMost(aeProgramsFeaturesScreen) == false)
                {
                    Console.WriteLine("aeProgramsFeaturesScreen is not TopMost ... wait 5 second");
                    Thread.Sleep(5000);
                    aeProgramsFeaturesScreen = AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
                }*/

                Console.WriteLine("Programs and Features  Window ready for user interaction");
                Console.WriteLine("Searching programs item button...");

                Console.WriteLine("---------------  statusOK:" + statusOK);
                if (statusOK == true)
                {
                    //AutomationElement aeGridView = AUIUtilities.GetElementByNameProperty(aeProgramsFeaturesScreen, "Folder View"); //  "文件夹视图"
                    AutomationElement aeGridView = aeProgramsFeaturesScreen.FindFirst(TreeScope.Children,
                        new PropertyCondition(AutomationElement.NameProperty, "Add or Remove Programs")); //  "文件夹视图"
                    if (aeGridView != null)
                    {
                        Console.WriteLine("Gridview found...");
                        #region find App menu item
                        // Set a property condition that will be used to find the control.
                        AutomationElementCollection aeProgram = aeGridView.FindAll(TreeScope.Children,
                                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
                        Console.WriteLine("Programs count ..." + aeProgram.Count);
                        // Make sure our window is usable. if the window is NOT ready. aeProgram.Count == 0
                        // if the window is ready. aeProgram.Count> 0
                        while (aeProgram.Count == 0)
                        {
                            Console.WriteLine("--------------------------------------- Programs count ..." + aeProgram.Count);
                            Thread.Sleep(5000);
                            aeProgram = aeGridView.FindAll(TreeScope.Children,
                                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
                            Thread.Sleep(5000);
                            
                        }
                        for (int i = 0; i < aeProgram.Count; i++)
                        {
                            Console.WriteLine("programs name: " + aeProgram[i].Current.Name);
                            if (App.Equals(EgeminApplication.EPIA))
                            {
                                if (aeProgram[i].Current.Name.StartsWith("E'pia Fr") || aeProgram[i].Current.Name.StartsWith("Epia"))
                                    aeApp = aeProgram[i];
                            }
                            else if (App.Equals(EgeminApplication.EPIA_RESOURCEFILEEDITOR))
                            {
                                if (aeProgram[i].Current.Name.StartsWith("E'pia Resource File Editor "))
                                {
                                    aeApp = aeProgram[i];
                                    Console.WriteLine(aeApp.Current.Name +"    ffffffffffffffffffffffound ---------------");
                                    Thread.Sleep(5000);
                                }
                            }
                            else if (App.Equals(EgeminApplication.ETRICC))
                            {
                                if (aeProgram[i].Current.Name.StartsWith("E'tricc")
                                        && aeProgram[i].Current.Name.IndexOf("Statistics") < 0
                                        && aeProgram[i].Current.Name.IndexOf("Shell") < 0
                                        && aeProgram[i].Current.Name.IndexOf("Playback") < 0
                                        && aeProgram[i].Current.Name.IndexOf("HostTest") < 0)
                                    aeApp = aeProgram[i];
                            }
                            else if (App.Equals(EgeminApplication.ETRICC_SHELL))
                            {
                                if (aeProgram[i].Current.Name.StartsWith("E'tricc Shell "))
                                    aeApp = aeProgram[i];
                            }
                            else if (App.Equals(EgeminApplication.ETRICC_PLAYBACK))
                            {
                                if (aeProgram[i].Current.Name.StartsWith("E'tricc Playback "))
                                    aeApp = aeProgram[i];
                            }
                            else if (App.Equals(EgeminApplication.ETRICC_HOSTTEST))
                            {
                                if (aeProgram[i].Current.Name.StartsWith("E'tricc HostTest "))
                                    aeApp = aeProgram[i];
                            }
                            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_PARSER))
                            {
                                if (aeProgram[i].Current.Name.StartsWith("E'tricc Statistics Parser "))
                                    aeApp = aeProgram[i];
                            }
                            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_PARSERCONFIGURATOR))
                            {
                                if (aeProgram[i].Current.Name.StartsWith("E'tricc Statistics ParserConfigurator "))
                                    aeApp = aeProgram[i];
                            }
                            else if (App.Equals(EgeminApplication.ETRICC_STATISTICS_UI))
                            {
                                if (aeProgram[i].Current.Name.StartsWith("E'tricc Statistics UI "))
                                    aeApp = aeProgram[i];
                            }
                            else if (App.Equals(EgeminApplication.AUTOMATICTESTING))
                            {
                                if (aeProgram[i].Current.Name.StartsWith(EgeminApplication.AUTOMATICTESTING))
                                    aeApp = aeProgram[i];
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        statusOK = false;
                        errorMsg = "Gridview NOT found...";
                        Console.WriteLine(errorMsg);
                        Thread.Sleep(5000);
                    }
                }
            }
            else
            {
                statusOK = false;
                errorMsg = "aeProgramsFeaturesScreen NOT found........";
                Console.WriteLine(errorMsg);
                Thread.Sleep(5000);
            }
            return statusOK;
        }

       
    }
}
