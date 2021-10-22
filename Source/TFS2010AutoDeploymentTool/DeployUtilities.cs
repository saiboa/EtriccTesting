using System;
using System.IO;
using System.Configuration;

using System.Windows;
using System.Windows.Automation;
using System.Threading;

using TestTools;

namespace TFS2010AutoDeploymentTool
{
    class DeployUtilities
    {
        static public AutomationElement GetMainWindow(string mainFormId)
        {
            AutomationElement aeWindow = null;
            AutomationElementCollection aeAllWindows = null;
            // find main window
            System.Windows.Automation.Condition cWindows = new PropertyCondition(
                AutomationElement.ControlTypeProperty, ControlType.Window);

            int k = 0;
            DateTime mStartTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mStartTime;
            while (aeWindow == null && mTime.TotalSeconds <= 120)
            {
                Console.WriteLine("aeWindow[k]=");
                k++;
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
                mTime = DateTime.Now - mStartTime;
            }

            return aeWindow;
        }

        static public void StartExecution()
        {
            //System.Windows.Forms.MessageBox.Show("sss");
            Thread.Sleep(20000);
            AutomationElement aeWindow = DeployUtilities.GetMainWindow("ToolsForm");
            if (aeWindow != null)
            {
                string id = "btnStartAuto";
                AutomationElement aeAutoStartButton = AUIUtilities.FindElementByID(id, aeWindow);
                if (aeAutoStartButton != null)
                {
                    Input.MoveToAndClick(AUIUtilities.GetElementCenterPoint(aeAutoStartButton));
                }
            }


        }

        static public bool CopySetupFilesWithWildcards(string fromPath, string toPath, string filenameWithWildcards,
            TestTools.Logger logger  /*, ref Tester.STATE m_State*/)
        {
            /*if (fromPath.StartsWith(@"\\"))
            {
                //if the first action fails try to logon to the server
                if (CreateDriveMap(fromPath) != 0)
                {
                    System.Windows.MessageBox.Show("CreateDriveMap2   failed:" + fromPath);
                    return false;
                }
            }*/

            if (!Directory.Exists(fromPath))
            {
                Directory.CreateDirectory(fromPath);
            }


            if (Directory.Exists(toPath))
            {
                DirectoryInfo DirInfo = new DirectoryInfo(toPath);
                FileInfo[] FilesToDelete = DirInfo.GetFiles();

                foreach (FileInfo file in FilesToDelete)
                {
                    try
                    {
                        FileAttributes attributes = FileAttributes.Normal;
                        File.SetAttributes(file.FullName, attributes);
                        file.Delete();
                    }
                    catch (Exception ex)
                    {
                        //if (m_Settings.EnableLog)
                        //{
                        //string logPath = Configuration.BuildInformationfilePath;
                        //string path = Path.Combine( logPath, Configuration.LogFilename );
                        //Logger logger = new Logger(path );
                        logger.LogMessageToFile("----------Setup Error  --------", 0, 0);
                        logger.LogMessageToFile("CopySetup Exception:" + ex.ToString(), 0, 0);
                        //Log("CopySetup Exception:" + ex.ToString());
                        //}
                        System.Windows.MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
                        //m_State = Tester.STATE.EXCEPTION;
                        return false;
                    }
                }
            }
            else
                Directory.CreateDirectory(toPath);

            FileInfo[] FilesToCopy;
            try
            {
                DirectoryInfo DirInfo = new DirectoryInfo(fromPath);
                FilesToCopy = DirInfo.GetFiles(filenameWithWildcards);

                foreach (FileInfo file in FilesToCopy)
                {
                    file.CopyTo(Path.Combine(toPath, file.Name));
                }
                //Log("Copied Setup from " + fromPath + " to " + toPath);
                try
                {
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    logger.LogMessageToFile("Copied Setup from " + fromPath + " to " + toPath, 0, 0);
                    //}
                }
                catch (Exception ex1)
                {
                    //if (m_Settings.EnableLog)
                    logger.LogMessageToFile("------ Test Exception : " + ex1.Message + "\r\n" + ex1.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                }
            }
            catch (Exception ex)
            {
                try
                {
                    //if (m_Settings.EnableLog)
                    //{
                    //string logPath = Configuration.BuildInformationfilePath;
                    //string path = Path.Combine( logPath, Configuration.LogFilename );
                    //Logger logger = new Logger(path );
                    logger.LogMessageToFile("----------Setup Error --------", 0, 0);
                    logger.LogMessageToFile("CopySetup Exception:" + ex.ToString(), 0, 0);
                    //Log("CopySetup Exception:" + ex.ToString());
                    //}
                }
                catch (Exception ex2)
                {
                    //if (m_Settings.EnableLog)
                    logger.LogMessageToFile("------ Test Exception : " + ex2.Message + "\r\n" + ex2.StackTrace, 0, 0);
                    //MessageBox.Show( ex.ToString() );
                }
                System.Windows.MessageBox.Show("FromPath=" + fromPath + "   " + ex.ToString() + "\r\n" + ex.StackTrace);
                //m_State = Tester.STATE.EXCEPTION;
                return false;
            }
            return true;
        }


    }
}



