using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;

using TestTools;
using TFSQATestTools;

namespace QATestProInstallationProcudure
{
    class ProMonitorTestsProgram
    {

        static AutomationElement aeForm;
        static string sErrorMessage;
        static DateTime sStartTime = DateTime.Now;
        static TimeSpan sTime;
        static bool sEventEnd;
        static void Main(string[] args)
        {
            int sResult = ConstCommon.TEST_UNDEFINED;
            ProMonitiorCheck("ProMonitor", aeForm, out sResult);

        }

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region ProMonitiorCheck
        public static void ProMonitiorCheck(string testname, AutomationElement root, out int result)
        {
            Console.WriteLine("\n=== Test " + testname + " ===");
            //Epia3Common.WriteTestLogTitle(slogFilePath, testname, Counter, sOnlyUITest); ;
            result = ConstCommon.TEST_UNDEFINED;
            //TestCheck = ConstCommon.TEST_PASS;

            string EpiaServerFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Server";
            string EpiaShellFolder = "C:\\" + TestTools.OSVersionInfoClass.ProgramFilesx86FolderName() + "\\Egemin\\Epia Shell";
            Thread.Sleep(5000);
            try
            {
                //Process[] procs = Process.GetProcessesByName(processName);
                Process[] procs = Process.GetProcesses();
                //Console.WriteLine(processName + " procs.Length:" + procs.Length + " at " + DateTime.Now.ToString("HH:mm:ss"));
                try
                {
                    for (int i = 0; i < procs.Length; i++)
                    {
                        if (procs[i].Responding)
                        {
                            //pID = procs[i].Id;
                            //proc = procs[i];
                            //Console.WriteLine("Proc name:" + procs[i].ProcessName);
                        }
                        else
                        {
                            procs[i].Kill();
                            Console.WriteLine("Kill Proc at:" + DateTime.Now.ToString("HH:mm:ss"));
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(" check proc exception:" + ex.Message, ex.StackTrace);
                    //pID = -1;
                }
                Thread.Sleep(2000);
                //ProcessUtilities.CloseProcess("ProMonitor");
                
                sEventEnd = false;
                #region  ProMonitor Window
                AutomationElement aeWindow = null;
                AutomationElement aeSecurityRoles = null;
                #region //xx
                try
                {
                   
                    aeWindow = ProMonitorUtilities.GetMainWindowFromName("Process Monitor");
                    if (aeWindow != null)
                    {
                        //EpiaUtilities.ClearDisplayedScreens(aeWindow);
                        //aeSystemService = AUICommon.FindTreeViewNodeLevel1(testname, aeWindow, "System", "Services", ref sErrorMessage);
                        //if (aeSystemService == null)
                        //{
                            sErrorMessage = "Process Monitor found " + " === " + sErrorMessage;
                            Console.WriteLine(sErrorMessage);

                        ControlType controlType = ControlType.Pane;
                        // Set a property condition that will be used to find the control.
                        Condition c = new PropertyCondition(
                            AutomationElement.ControlTypeProperty, controlType);

                        string[] ProcessName = new string[13];
                        string[] ProcessAction = new string[13];
                        string[] ProcessStatus = new string[13];
                        int ix = 0;
                        int iy = 0;
                        int iz = 0;
                        // Find the element.
                        AutomationElementCollection aePanes = aeWindow.FindAll(TreeScope.Children, c);

                        //AutomationElement aeButtonAdd = AUIUtilities.FindElementByType(ControlType.Edit, aeSelectedWindow);
                        if (aePanes == null)
                        {
                            Console.WriteLine("aePanes not find :" );
                            //Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            //TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine("aePanes.Count=" + aePanes.Count);
                            Thread.Sleep(1000);
                            for (int i = 0; i < aePanes.Count; i++)
                            {
                                AutomationElement aePane = aePanes[i];
                                string paneName = aePane.Current.Name;
                                Console.WriteLine("----   aePanes[" + i + "].Name=" + paneName);

                                if (i % 4 == 0)
                                {
                                    Console.WriteLine("000000----   aePanes[" + i + "].Name=" + paneName);
                                    ProcessStatus[ix] = paneName;
                                    ix++;
                                }

                                if (i % 4 == 1)
                                {
                                    Console.WriteLine("111111----   aePanes[" + i + "].Name=" + paneName);
                                    ProcessAction[iy] = paneName;
                                    iy++;
                                }

                                if (i % 4 == 3)
                                {
                                    Console.WriteLine("333333----   aePanes[" + i + "].Name=" + paneName);
                                    ProcessName[iz] = paneName;
                                    iz++;
                                }

                                Thread.Sleep(500);
                                if (i >= 51)
                                    break;
                                
                            }

                            for (int i = 0; i < 13; i++)
                            {
                                Console.WriteLine("ProcessName[" + i + "]=" + ProcessName[i]);
                                Console.WriteLine("ProcessStatus[" + i + "]=" + ProcessStatus[i]);
                                Console.WriteLine("ProcessAction[" + i + "]=" + ProcessAction[i]);
                                

                            }
                            Thread.Sleep(100000);
                        }
                        //Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        //TestCheck = ConstCommon.TEST_FAIL;
                        //}
                        //else
                        //{
                        //Point Pnt = AUIUtilities.GetElementCenterPoint(aeSystemService);
                        //Input.MoveToAndClick(Pnt);
                        //Thread.Sleep(5000);
                        //}
                    }
                    else
                    {
                        sErrorMessage = "Process Monitor not found";
                        //Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        //TestCheck = ConstCommon.TEST_FAIL;
                    }

                    int k = 0;
                    /*if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        AutomationElement aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow("Services", ref sErrorMessage);
                        while (aeSelectedWindow == null && k < 5)
                        {
                            Console.WriteLine("wait until selected window open :" + k++);
                            aeSelectedWindow = ProjBasicUI.GetSelectedOverviewWindow("Services", ref sErrorMessage);
                            Thread.Sleep(5000);
                        }

                        if (aeSelectedWindow == null)
                        {
                            Console.WriteLine("Window not opened :" + sErrorMessage);
                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            // find Add...button
                            //System.Windows.Automation.Condition cButtonAdd = new AndCondition(
                            //    new PropertyCondition(AutomationElement.NameProperty, "m_TxtFilter"),
                            //    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)
                            //);
                            //AutomationElement aeButtonAdd = aeSelectedWindow.FindFirst(TreeScope.Element | TreeScope.Descendants, cButtonAdd);
                            AutomationElement aeButtonAdd = AUIUtilities.FindElementByType(ControlType.Edit, aeSelectedWindow);
                            if (aeButtonAdd == null)
                            {
                                Console.WriteLine("aeButtonAdd not find :" + aeSelectedWindow.Current.Name);
                                Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                TestCheck = ConstCommon.TEST_FAIL;
                            }
                            else
                            {
                                double x = aeButtonAdd.Current.BoundingRectangle.Right + 70.0;
                                double y = (aeButtonAdd.Current.BoundingRectangle.Bottom + aeButtonAdd.Current.BoundingRectangle.Top) / 2.0;
                                Point pt = new Point(x, y);
                                for (int iservice = 1; iservice < 2; iservice++)
                                {
                                    Input.MoveTo(pt);
                                    Input.ClickAtPoint(pt);
                                    Thread.Sleep(3000);
                                    if (iservice == 1 && TestCheck == ConstCommon.TEST_PASS)
                                    {
                                        if (EpiaUtilities.AddService(slogFilePath, "Egemin Epia Server", PCName, sOnlyUITest, ref sErrorMessage) == false)
                                        {
                                            Console.WriteLine(sErrorMessage);
                                            Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                                            TestCheck = ConstCommon.TEST_FAIL;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }*/

                    // validate result
                    /*if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        #region // Find GridView
                        aeWindow = EpiaUtilities.GetMainWindow("MainForm");
                        AutomationElement aeGrid = AUIUtilities.FindElementByID("m_GridData", aeWindow);
                        if (aeGrid == null)
                        {
                            sErrorMessage = aeWindow.Current.Name + " GridData not found";
                            Console.WriteLine(sErrorMessage);
                            TestCheck = ConstCommon.TEST_FAIL;
                        }
                        else
                        {
                            Console.WriteLine(aeWindow.Current.Name + " GridData found at time: " + System.DateTime.Now.ToString("HH:mm:ss"));
                            Thread.Sleep(3000);


                            string ServiceName = "Egemin Epia Server";
                            // Construct the Grid Cell Element Name
                            for (int iservicename = 0; iservicename < 1; iservicename++)
                            {
                                if (iservicename == 0)
                                    ServiceName = "Egemin Epia Server";

                                string cellname = "Service" + " Row " + iservicename;
                                // Get the Element with the Row Col Coordinates
                                AutomationElement aeCell = AUIUtilities.FindElementByName(cellname, aeGrid);
                                if (aeCell == null)
                                {
                                    sErrorMessage = "Find aeCell failed:" + cellname;
                                    Console.WriteLine(sErrorMessage);
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    break;
                                }
                                else
                                {
                                    string cellValue = string.Empty;
                                    try
                                    {
                                        ValuePattern vp = aeCell.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                                        cellValue = vp.Current.Value;
                                        Console.WriteLine("Get element.Current Value:" + cellValue);
                                    }
                                    catch (System.NullReferenceException)
                                    {
                                        cellValue = string.Empty;
                                    }

                                    if (cellValue == null || cellValue == string.Empty)
                                    {
                                        sErrorMessage = "aeCell Value not found:" + cellname;
                                        Console.WriteLine(sErrorMessage);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                        break;
                                    }
                                    else if (!cellValue.Equals(ServiceName))
                                    {
                                        sErrorMessage = "aeCell Value not equal " + ServiceName + " , but :" + cellValue;
                                        Console.WriteLine(sErrorMessage);
                                        TestCheck = ConstCommon.TEST_FAIL;
                                        break;
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    */
                    /*
                    if (TestCheck == ConstCommon.TEST_FAIL)
                    {
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, sErrorMessage, sOnlyUITest);
                        result = ConstCommon.TEST_FAIL;
                    }

                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        sErrorMessage = string.Empty;
                        Console.WriteLine("\nTest scenario" + testname + " : Pass");
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                        result = ConstCommon.TEST_PASS;
                    }
                    */
                }
                catch (Exception ex)
                {
                    result = ConstCommon.TEST_EXCEPTION;
                    sErrorMessage = ex.Message;
                    Console.WriteLine(testname + " === " + sErrorMessage);
                    //Epia3Common.WriteTestLogFail(slogFilePath, testname + " === " + sErrorMessage, sOnlyUITest);
                }
                finally
                {
                    Thread.Sleep(3000);
                    //Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                    //	   AutomationElement.RootElement,
                    //	  UIALayoutCancelButtonEventHandler);
                }


                #endregion
                /*
                sErrorMessage = string.Empty;
                ProcessUtilities.StartProcessNoWait(OSVersionInfoClass.ProgramFilesx86() + @"\Egemin\EtriccGL\Server\bin",
                    "ProMonitor.exe", string.Empty);
                //ConstCommon.EGEMIN_EPIA_SHELL_EXE, string.Empty);

                sStartTime = DateTime.Now;
                sTime = DateTime.Now - sStartTime;
                int wt = 0;
                Console.WriteLine(" time is :" + sTime.TotalSeconds);
                while (sEventEnd == false && wt < 60)
                {
                    Thread.Sleep(2000);
                    //sTime = DateTime.Now - sStartTime;
                    wt = wt + 2;
                    Console.WriteLine("wait shell start up time is (sec) : " + wt);
                }

                Console.WriteLine("Shell started after (sec) : " + 2 * wt);
                */
                /*Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent,
                       AutomationElement.RootElement, UIAShellEventHandler);

                Thread.Sleep(4000);
                Console.WriteLine("TestCkeck : " + TestCheck.ToString());
                if (TestCheck == ConstCommon.TEST_FAIL)
                {
                    throw new Exception("shell start up failed:" + sErrorMessage);
                }*/
                Thread.Sleep(4000);
                #endregion
                // find Window PRO MONITOR
                
                //ProcessUtilities.CloseProcess("Egemin.Epia.Server");
                Console.WriteLine("--------------- " + DeployUtilities.getThisPCOS());

                #region ProMonitior 1
                // uninstall playback if already installed:
                /*if (DeployUtilities.getThisPCOS().StartsWith("WindowsXP.32"))
                {
                    Console.WriteLine("---------------  UninstallApplicationXP");
                    //Thread.Sleep(20000);
                    if (ProjAppInstall.UninstallApplicationXP(EgeminApplication.EPIA, ref sErrorMessage) == false)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = EgeminApplication.EPIA + " Uninstall failed:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                    }
                }
                else
                {
                    if (ProjAppInstall.UninstallApplication(EgeminApplication.EPIA, ref sErrorMessage) == false)
                    {
                        TestCheck = ConstCommon.TEST_FAIL;
                        sErrorMessage = EgeminApplication.EPIA + " Uninstall failed:" + sErrorMessage;
                        Console.WriteLine(sErrorMessage);
                    }
                }

                if (TestCheck == ConstCommon.TEST_PASS)
                {
                    if (System.IO.Directory.Exists(EpiaServerFolder)
                        || System.IO.Directory.Exists(EpiaShellFolder))
                    {
                        if (System.IO.Directory.Exists(EpiaServerFolder))
                        {
                            // get files in ServerFolder
                            DirectoryInfo DirInfo = new DirectoryInfo(EpiaServerFolder);
                            FileInfo[] serverFolderFiles = DirInfo.GetFiles("*.*");
                            if (serverFolderFiles.Length > 1)
                            {
                                TestCheck = ConstCommon.TEST_FAIL;
                                sErrorMessage = EpiaServerFolder + " still has some files:" + serverFolderFiles[0].FullName;
                                Console.WriteLine(sErrorMessage);
                                //System.Windows.Forms.MessageBox.Show(sErrorMessage);
                            }
                        }

                        if (TestCheck == ConstCommon.TEST_PASS)
                        {
                            if (System.IO.Directory.Exists(EpiaShellFolder))
                            {
                                // get files in ShellFolder
                                DirectoryInfo DirInfo = new DirectoryInfo(EpiaShellFolder);
                                FileInfo[] shellFolderFiles = DirInfo.GetFiles("*.*");
                                if (shellFolderFiles.Length > 0)
                                {
                                    TestCheck = ConstCommon.TEST_FAIL;
                                    sErrorMessage = EpiaShellFolder + " still has some files:" + shellFolderFiles[0].FullName;
                                    Console.WriteLine(sErrorMessage);
                                    //System.Windows.Forms.MessageBox.Show(sErrorMessage);
                                }
                            }
                        }
                    }
                    else
                    {
                        TestCheck = ConstCommon.TEST_PASS;
                    }
                }

                if (testname.ToLower().StartsWith("no"))
                {
                    Console.WriteLine("do nothing, not test case:" + testname);
                }
                else
                {
                    if (TestCheck == ConstCommon.TEST_FAIL)
                    {
                        result = ConstCommon.TEST_FAIL;
                        Console.WriteLine(sErrorMessage);
                        Epia3Common.WriteTestLogFail(slogFilePath, testname + ":" + sErrorMessage, sOnlyUITest);
                    }

                    if (TestCheck == ConstCommon.TEST_PASS)
                    {
                        result = ConstCommon.TEST_PASS;
                        sErrorMessage = string.Empty;
                        Epia3Common.WriteTestLogPass(slogFilePath, testname, sOnlyUITest);
                    }
                }
                */
                #endregion
                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                result = ConstCommon.TEST_EXCEPTION;
                Console.WriteLine("Fatal error: " + ex.Message + "----: " + ex.StackTrace);
                //Epia3Common.WriteTestLogMsg(slogFilePath, "===" + ex.Message + "------" + ex.StackTrace, sOnlyUITest);
            }
            finally
            {
                if (System.IO.Directory.Exists(EpiaServerFolder))
                {
                    //DirectoryInfo dirInfo = new DirectoryInfo(EpiaServerFolder);
                    while (System.IO.Directory.Exists(EpiaServerFolder))
                    {
                        //FileManipulation.DeleteRecursiveFolder(dirInfo);
                        Thread.Sleep(2000);
                    }
                }

                if (System.IO.Directory.Exists(EpiaShellFolder))
                {
                    //DirectoryInfo dirInfo = new DirectoryInfo(EpiaShellFolder);
                    while (System.IO.Directory.Exists(EpiaShellFolder))
                    {
                        //FileManipulation.DeleteRecursiveFolder(dirInfo);
                        Thread.Sleep(2000);
                    }
                }
            }
        }
        #endregion Epia4CleanUninstallCheckXP

    }


}
