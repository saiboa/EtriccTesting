using System;
using System.Collections.Specialized;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace TestTools
{
    public class TestListUtilities
    {
        public static bool IsAllTestDefinitionsTested(string[] mTestDefinitionTypes, string sTestResultFolder,
                                                      ref string sErrorMessage)
        {
            bool found = false;
            try
            {
                string[] testedDefs = File.ReadAllLines(Path.Combine(sTestResultFolder, ConstCommon.TESTINFO_FILENAME));

                for (int i = 0; i < mTestDefinitionTypes.Length; i++)
                {
                    Console.WriteLine(i + " testdefinition : " + mTestDefinitionTypes[i]);
                }

                for (int j = 0; j < mTestDefinitionTypes.Length; j++)
                {
                    found = false;
                    for (int k = 0; k < testedDefs.Length; k++)
                    {
                        if (testedDefs[k].IndexOf(mTestDefinitionTypes[j]) >= 0)
                        {
                            found = true;
                            Console.WriteLine("++++++ test def type tested : " + mTestDefinitionTypes[j]);
                            break;
                        }
                    }

                    if (found == false)
                    {
                        Console.WriteLine("------ test def type not tested : " + mTestDefinitionTypes[j]);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception  " + ex.Message + "---" + ex.StackTrace, "IsAllTestDefinitionsTested");
            }

            return found;
        }

        public static bool IsAllTestStatusPassed(string[] mTestDefinitionTypes, string sTestResultFolder,
                                                 ref string sErrorMessage)
        {
            string path = Path.Combine(sTestResultFolder, ConstCommon.TESTINFO_FILENAME);
            string[] testinfos = File.ReadAllLines(path);
            //var teststatus = new string[testinfos.Length];
            int passCnt = 0;
            Console.WriteLine("<<<>>>> testinfos.Length : " + testinfos.Length);
            try
            {
                for (int i = 1; i < testinfos.Length; i++)
                {
                    Console.WriteLine("<<< testinfos[" + i + "] : " + testinfos[i]);
                    //Console.WriteLine(i + " length : " + (testinfos[i].IndexOf(":") - testinfos[i].IndexOf("-")));
                    if (testinfos[i].Trim().Length > 10) // some time by manual edit info file, info line can be empty
                    {
                        /*if (testinfos[i].IndexOf("-") >= 0)
                        {
                            // Windows7.32[x86.Debug]EPIAAUTOTEST1-GUI Tests Passed:Tests OK
                            teststatus[i] = testinfos[i].Substring(testinfos[i].IndexOf("-") + 1);
                                // become : GUI Tests Passed:Tests OK     
                            Console.WriteLine("<<< teststatus[" + i + "] : " + teststatus[i]);
                            if (teststatus[i].Contains("GUI Tests Passed"))
                                passCnt++;
                        }
                        else //Windows7.64[AnyCPU.Protected] --> no '-' not tested --> not passed
                        {
                            teststatus[i] = testinfos[i];
                        }*/
                        Console.WriteLine("<<< testinfos[" + i + "] : " + testinfos[i]);
                        if (testinfos[i].Contains("GUI Tests Passed"))
                            passCnt++;

                    }
                    else
                        passCnt++; // for empty or corrupt line, also consider as pass line
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(" exception  " + ex.Message + "---" + ex.StackTrace,
                                "IsAllTestStatusPassed:" + "testinfos.Length: " + testinfos.Length);
            }

            if (passCnt == (testinfos.Length - 1))
            {
                StreamWriter sw = File.AppendText(path);
                try
                {
                    sw.WriteLine("All Tests Complete: " + DateTime.Now.ToString());
                }
                finally
                {
                    sw.Close();
                }

                return true;
            }
            else
                return false;
        }

        public static bool UpdateStatusInTestInfoFile(string path, string status, string message, string infoFileKey)
        {
            bool updateOK = false;
            var AllLines = new StringCollection();
            var NewAllLines = new StringCollection();

            // Read all lines from test info file
            StreamReader reader = null;
            StreamWriter write = null;
            while (updateOK == false)
            {
                try
                {
                    reader = File.OpenText(Path.Combine(path, ConstCommon.TESTINFO_FILENAME));
                    string infoline = reader.ReadLine();
                    while (infoline != null)
                    {
                        AllLines.Add(infoline);
                        infoline = reader.ReadLine();
                    }
                    reader.Close();

                    // Update info file
                    foreach (string line in AllLines)
                    {
                        if (line.StartsWith(infoFileKey))
                            NewAllLines.Add(infoFileKey + "-" + status + ":" + message);
                        else
                        {
                            NewAllLines.Add(line);
                        }
                    }

                    write = File.CreateText(Path.Combine(path, ConstCommon.TESTINFO_FILENAME));
                    foreach (string line in NewAllLines)
                    {
                        write.WriteLine(line);
                    }
                    write.Close();
                    updateOK = true;
                }
                catch (IOException ex)
                {
                    string parentPath = Directory.GetParent(path).FullName;
                    if (Directory.Exists(parentPath))
                    {
                        updateOK = true;
                        message = path + " is deleted by Wim";
                    }
                    else
                    {
                        updateOK = false;
                        MessageBoxEx.Show(
                            ex.Message + " --- " + ex.StackTrace + "\nWill try to reconnect the Server ...",
                            "UpdateStatusInTestInfoFile IOException", 60000);
                        Thread.Sleep(60000);
                    }
                }
                catch (Exception ex2)
                {
                    updateOK = false;
                    MessageBoxEx.Show(ex2.Message + " --- " + ex2.StackTrace + "\nWill try to reconnect the Server ...",
                                      "UpdateStatusInTestInfoFile Exception", 60000);
                    Thread.Sleep(60000);
                }
                finally
                {
                    if ( reader != null)
                        reader.Close();

                    if ( write != null)
                        write.Close();
                }
            }

            return updateOK;
        }

        public static bool IsAllTestStatusPassed(string logPath, string[] mTestDefinitionTypes, string sTestResultFolder,
                                                 ref string sErrorMessage)
        {
            string path = Path.Combine(sTestResultFolder, ConstCommon.TESTINFO_FILENAME);
            string[] testinfos = File.ReadAllLines(path);
            var teststatus = new string[testinfos.Length];
            int passCnt = 0;
            Console.WriteLine("<<<>>>> testinfos.Length : " + testinfos.Length);
            Epia3Common.WriteTestLogMsg(logPath, path + "  with <<<>>>> testinfos.Length : " + testinfos.Length, false);
            try
            {
                for (int i = 1; i < testinfos.Length; i++)
                {
                    Console.WriteLine("<<< testinfos[" + i + "] : " + testinfos[i]);
                    Epia3Common.WriteTestLogMsg(logPath, "<<< testinfos[" + i + "] : " + testinfos[i], false);
                    //Console.WriteLine(i + " length : " + (testinfos[i].IndexOf(":") - testinfos[i].IndexOf("-")));
                    if (testinfos[i].Trim().Length > 10) // some time by manual edit info file, info line can be empty
                    {
                        if (testinfos[i].IndexOf("-") >= 0)
                        {
                            // Windows7.32[x86.Debug]EPIAAUTOTEST1-GUI Tests Passed:Tests OK
                            teststatus[i] = testinfos[i].Substring(testinfos[i].IndexOf("-") + 1);
                                // become : GUI Tests Passed:Tests OK     
                            Console.WriteLine("<<< teststatus[" + i + "] : " + teststatus[i]);
                            Epia3Common.WriteTestLogMsg(logPath, "<<< teststatus[" + i + "] : " + teststatus[i], false);
                            if (teststatus[i].StartsWith("GUI Tests Passed"))
                                passCnt++;
                        }
                        else //Windows7.64[AnyCPU.Protected] --> no '-' not tested --> not passed
                        {
                            teststatus[i] = testinfos[i];
                        }
                    }
                    else
                    {
                        passCnt++; // for empty or corrupt line, also consider as pass line
                        Epia3Common.WriteTestLogMsg(logPath,
                                                    "// for empty or corrupt line, also consider as pass line" + passCnt,
                                                    false);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(" exception  " + ex.Message + "---" + ex.StackTrace,
                                "IsAllTestStatusPassed:" + "testinfos.Length: " + testinfos.Length);
            }

            Epia3Common.WriteTestLogMsg(logPath, "passCnt : " + passCnt, false);
            Epia3Common.WriteTestLogMsg(logPath, "testinfos.Length: " + testinfos.Length, false);
            if (passCnt == (testinfos.Length - 1))
            {
                StreamWriter sw = File.AppendText(path);
                try
                {
                    sw.WriteLine("All Tests Complete: " + DateTime.Now.ToString());
                }
                finally
                {
                    sw.Close();
                }

                return true;
            }
            else
                return false;
        }
    }
}