using System;
using System.Windows.Forms;
using Microsoft.Win32;

namespace TestTools
{
    public class RegisterEditor
    {
        #region Methods of RegisterEditor (3)

        public static void InitEpiaService(string version)
        {
            string strKey = "SOFTWARE\\Egemin\\EPIA\\" + version + "\\E'pia Service\\";
            try
            {
                RegistryKey TestProject = Registry.LocalMachine.OpenSubKey(strKey, true);
                TestProject.SetValue("activate", "True", RegistryValueKind.String);
                TestProject.SetValue("afteractivate", "False", RegistryValueKind.String);
                TestProject.SetValue("afterdeactivate", "False", RegistryValueKind.String);
                TestProject.SetValue("assemblyname", "Egemin.EPIA.WCS", RegistryValueKind.String);
                TestProject.SetValue("beforeactivate", "False", RegistryValueKind.String);
                TestProject.SetValue("beforedeactivate", "False", RegistryValueKind.String);
                TestProject.SetValue("coldstart", "False", RegistryValueKind.String);
                TestProject.SetValue("connectionstring", "", RegistryValueKind.String);
                TestProject.SetValue("loadsession", "False", RegistryValueKind.String);
                TestProject.SetValue("objectid", "", RegistryValueKind.String);
                TestProject.SetValue("objectname", "Project", RegistryValueKind.String);
                TestProject.SetValue("objecttype", "Egemin.EPIA.WCS.Core.Project", RegistryValueKind.String);
                TestProject.SetValue("persistencytype", "XML", RegistryValueKind.String);
                TestProject.SetValue("projectfile", "", RegistryValueKind.String);
                TestProject.SetValue("save", "False", RegistryValueKind.String);
                TestProject.SetValue("savesession", "False", RegistryValueKind.String);
                TestProject.SetValue("sessionname", "", RegistryValueKind.String);
                TestProject.SetValue("xmlfile",
                                     "C:\\EpiaTestCenter2\\AutoTestCenter\\Main\\Data\\Xml\\TestEurobalticService.xml",
                                     RegistryValueKind.String);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex + Environment.NewLine + ex.StackTrace,
                                "TestTool RegisterEditor   InitEpiaService Error:" + strKey);
            }
        }

        public static void SetEpiaServiceRegisterNameAndValue(string version, string name, string value)
        {
            string strKey = "SOFTWARE\\Egemin\\EPIA\\" + version + "\\E'pia Service\\";
            try
            {
                RegistryKey TestProject = Registry.LocalMachine.OpenSubKey(strKey, true);
                TestProject.SetValue(name, value, RegistryValueKind.String);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex + Environment.NewLine + ex.StackTrace,
                                "TestTool RegisterEditor  SetEpiaServiceRegisterNameAndValue Error:" + strKey +
                                "  name :" + name);
            }
        }

        public static void SetRegister(string type, string strKey, string kind, string name, string value)
        {
            //string ProductSpecific = "ProductSpecific";
            //string version = epiaVersion.Substring(5);
            //string strKey = "SOFTWARE\\Egemin\\EPIA\\" + version + "\\E'pia Launcher\\";
            try
            {
                if (type.ToUpper().Equals("LOCALMACHINE"))
                {
                    RegistryKey TestProject = Registry.LocalMachine.OpenSubKey(strKey, true);
                    if (kind.ToUpper().Equals("DWORD"))
                        TestProject.SetValue(name, value, RegistryValueKind.DWord);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex + Environment.NewLine + ex.StackTrace,
                                "DeployTestLogic.Tester  UpdateLauncher Register Error:" + strKey);
            }
        }

        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
    }
}