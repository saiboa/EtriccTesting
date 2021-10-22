using System;
using System.IO;
using Microsoft.VisualBasic.Devices;

namespace TestTools
{
    public class HelpUtilities
    {
        #region Constants of HelpUtilities (1)

        private const Int64 MEGABYTES = 1024000;

        #endregion // —— Constants ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        #region Fields of HelpUtilities (1)

        private static Computer mc;

        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        #region Methods of HelpUtilities (8)

        public static string GetBuildAndTestInfo()
        {
            //string root = m_Project.Facilities["Tests"].Parameters["TestCenterRoot"].ValueAsString;
            string root = @"C:\EpiaTestCenter2\AutoTestCenter\Main\Source";
            string SetupInfoOutputFilename = "TestAutoDeploymentOutputLog.txt";
            StreamReader reader = File.OpenText(Path.Combine(root, SetupInfoOutputFilename));
            string TestedInfo = reader.ReadLine();
            reader.Close();

            return TestedInfo;
        }

        private static string GetDriveInfo()
        {
            string strHDInfo = "";
            string strHeading = "";

            // Make column headings 
            strHeading += "Drive Letter, ";
            strHeading += "File Format, ";
            strHeading += "Free Space, ";
            strHeading += "Drive Type";
            strHDInfo += strHeading + "\r\n";

            // Get Drive Info            
            string strDL;
            try
            {
                foreach (DriveInfo objDI in DriveInfo.GetDrives())
                {
                    strDL = objDI.Name.Substring(0, 1);
                    strHDInfo += strDL + ":, ";
                    if (objDI.DriveType == DriveType.Fixed)
                    {
                        strHDInfo += objDI.DriveFormat + ", ";
                        strHDInfo += (objDI.AvailableFreeSpace/MEGABYTES).ToString() + " MB, ";
                        strHDInfo += objDI.DriveType.ToString();
                    }
                    else if (objDI.DriveType == DriveType.CDRom)
                    {
                        strHDInfo += "NA,";
                        strHDInfo += "NA,";
                        strHDInfo += objDI.DriveType.ToString();
                    }
                    else if (objDI.DriveType == DriveType.Removable)
                    {
                        strHDInfo += "NA,";
                        strHDInfo += "NA,";
                        strHDInfo += objDI.DriveType.ToString();
                    }
                    else if (objDI.DriveType == DriveType.Network)
                    {
                        strHDInfo += objDI.DriveFormat + ", ";
                        strHDInfo += (objDI.AvailableFreeSpace/MEGABYTES).ToString() + " MB, ";
                        strHDInfo += objDI.DriveType.ToString();
                    }
                    strHDInfo += "\r\n";
                }
            }
            catch (Exception ex)
            {
                strHDInfo += ex.Message + "- AND -" + ex.StackTrace;
            }
            return strHDInfo;
        }

        private static string GetMemoryInfo()
        {
            string strMemoryInfo = "";
            string strHeading = "";
            strHeading += "Total Phyical Memory, ";
            strHeading += "Total Virtual Memory, ";
            strHeading += "Available Phyical Memory, ";
            strHeading += "Available Virtual Memory ";
            strMemoryInfo += strHeading + "\r\n";
            UInt64 intTPhysicalMem = mc.Info.TotalPhysicalMemory;
            UInt64 intTVirtualMem = mc.Info.TotalVirtualMemory;
            UInt64 intAPhysicalMem = mc.Info.AvailablePhysicalMemory;
            UInt64 intAVirtualMem = mc.Info.AvailableVirtualMemory;
            strMemoryInfo += (intTPhysicalMem/MEGABYTES).ToString() + " MB , ";
            strMemoryInfo += (intTVirtualMem/MEGABYTES).ToString() + "MB , ";
            strMemoryInfo += (intAPhysicalMem/MEGABYTES).ToString() + " MB , ";
            strMemoryInfo += (intAVirtualMem/MEGABYTES).ToString() + "MB ";
            strMemoryInfo += "\r\n";
            return strMemoryInfo;
        }

        public static void GetMemoryInfo(out ulong TPhysicalMem, out ulong APhysicalMem,
                                         out ulong TVirtualMem, out ulong AVirtualMem)
        {
            mc = new Computer();
            UInt64 intTPhysicalMem = mc.Info.TotalPhysicalMemory;
            UInt64 intTVirtualMem = mc.Info.TotalVirtualMemory;
            UInt64 intAPhysicalMem = mc.Info.AvailablePhysicalMemory;
            UInt64 intAVirtualMem = mc.Info.AvailableVirtualMemory;
            TPhysicalMem = (intTPhysicalMem/MEGABYTES);
            TVirtualMem = (intTVirtualMem/MEGABYTES);
            APhysicalMem = (intAPhysicalMem/MEGABYTES);
            AVirtualMem = (intAVirtualMem/MEGABYTES);
        }

        public static void GetPCInfo(out string name, out string os, out string osversion, out string uiculture,
                                     out string timeonpc)
        {
            mc = new Computer();
            name = mc.Name;
            os = mc.Info.OSFullName;
            osversion = mc.Info.OSVersion;
            uiculture = mc.Info.InstalledUICulture.EnglishName;
            timeonpc = mc.Clock.LocalTime.ToString();
        }

        private static string GetPCInfo()
        {
            string strPCInfo = "";
            string strHeading = "";
            strHeading += "Computer Name, ";
            strHeading += "Operating System, ";
            strHeading += "Operating System Version, ";
            strHeading += "UI Culture, ";
            strHeading += "Time on PC";

            mc = new Computer();

            strPCInfo += strHeading + "\r\n";
            strPCInfo += mc.Name + ", ";
            strPCInfo += mc.Info.OSFullName + ", ";
            strPCInfo += mc.Info.OSVersion + ", ";
            strPCInfo += mc.Info.InstalledUICulture.EnglishName + ", ";
            strPCInfo += mc.Clock.LocalTime;
            strPCInfo += "\r\n";
            return strPCInfo;
        }

        public static string GetPCOS()
        {
            var mc = new Computer();
            return mc.Info.OSFullName;
        }

        public static void SavePCInfo(string stdout)
        {
            string strData = "";
            strData += "PC Information" + "\r\n";
            strData += GetPCInfo() + "\r\n";
            strData += "Drive Information" + "\r\n";
            strData += GetDriveInfo() + "\r\n";
            strData += "Memory Information" + "\r\n";
            strData += GetMemoryInfo() + "\r\n";

            try
            {
                if (stdout == "y")
                {
                    var DI = new DirectoryInfo(Directory.GetCurrentDirectory() + "\\PCInfo");

                    if (!(DI.Exists))
                    {
                        DI.Create();
                    }
                    //MessageBox.Show(System.IO.Directory.GetCurrentDirectory() + "\\PCInfo", "xxxxxxxxxx");
                    //strData += "xxxxx";
                    //strData += getRootPath() + "PCInfo" + "\r\n"; 

                    WriteReport(strData);
                }
                else if (stdout == "n")
                {
                    Console.WriteLine(strData);
                }
                else
                {
                    Console.WriteLine("You must chose to write to file or not.");
                    Console.WriteLine("y or n after the program name");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private static void WriteReport(string Data)
        {
            StreamWriter objSW;
            try
            {
                string infoFile = Path.Combine(Directory.GetCurrentDirectory() + "\\PCInfo", "PCInfo.csv");
                objSW = new StreamWriter(infoFile, false);
                objSW.Write(Data);
                objSW.Close();
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 

        #region Nested type: PCINFO

        public struct PCINFO
        {
            public string pcName;
            public string pcOS;
            public string pcOSVersion;
            public string pcUICulture;
            public string timeOnPC;

            public override String ToString()
            {
                String str = "pcName: " + pcName + " pcOS: " + pcOS + " pcOSVersion: " + pcOSVersion + " pcUICulture: " +
                             pcUICulture + " timeOnPC: " + timeOnPC;
                return (str);
            }
        }

        #endregion
    }
}