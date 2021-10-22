using System;
using System.IO;
using System.Management;
using System.Reflection;
using System.Security.Principal;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace TestTools
{
    public class FileManipulation
    {
        public static string CreateOutputInfoFileName(string sInfoFileKey, bool autoTest)
        {
            // out filename 
            string outFilename = sInfoFileKey + "." + DateTime.Now.ToString("MMdd-HH-mm-ss");
            outFilename = outFilename.Replace("[", ".");
            outFilename = outFilename.Replace("]", ".");

            if (!autoTest)
                outFilename = "Manual-" + outFilename;
            return outFilename;
        }

        public static void ShareFolderPermission(string FolderPath, string ShareName, string Description,
                                                 ref string errorMsg)
        {
            try
            {
                // Calling Win32_Share class to create a shared folder
                var managementClass = new ManagementClass("Win32_Share");
                // Get the parameter for the Create Method for the folder
                ManagementBaseObject inParams = managementClass.GetMethodParameters("Create");
                ManagementBaseObject outParams;
                // Assigning the values to the parameters
                inParams["Description"] = Description;
                inParams["Name"] = ShareName;
                inParams["Path"] = FolderPath;
                inParams["Type"] = 0x0;
                // Finally Invoke the Create Method to do the process
                outParams = managementClass.InvokeMethod("Create", inParams, null);
                // Validation done here to check sharing is done or not
                if ((uint) (outParams.Properties["ReturnValue"].Value) != 0)
                    MessageBox.Show("Folder might be already in share or unable to share the directory");
            }
            catch (Exception ex)
            {
                errorMsg = "Create ShareFolderPermission:" + ex.Message + " --- " + ex.StackTrace;
                MessageBox.Show(errorMsg, "create share folder:" + FolderPath);
            }
        }

        public static void DeleteRecursiveFolder(DirectoryInfo dirInfo)
        {
            foreach (DirectoryInfo subDir in dirInfo.GetDirectories())
            {
                DeleteRecursiveFolder(subDir);
            }

            foreach (FileInfo file in dirInfo.GetFiles())
            {
                file.Attributes = FileAttributes.Normal;
                file.Delete();
            }

            dirInfo.Delete();
        }

        public static bool CopyFilesWithWildcards(string origFullPath, string destFullPath, ref string errorMsg)
        {
            bool success = true;
            //This if statement is safe because
            //Files can not be named with the following characters:
            //\ / : ? <> |and,most importantly, *
            if (origFullPath.IndexOf("*") != -1)
            {
                //This is a wild card copy;
                //Seperating the Directory Path and the Filename
                int directoryEnd = origFullPath.LastIndexOf("\\");
                string directoryPath = origFullPath.Substring(0, directoryEnd + 1);
                string fileName = origFullPath.Substring(directoryEnd + 1, origFullPath.Length - (directoryEnd + 1));

                //Directory.GetFiles will return anarray of strings that contain the full file path.
                string[] allFiles = Directory.GetFiles(directoryPath, fileName);
                if (allFiles.Length == 0)
                    //There are no files matching-- so hey -- success! The files are 
                    //not present in the directory, which was our intended result!
                    success = true;
                else
                {
                    success = true;
                    for (int i = 0; i < allFiles.Length; i++)
                    {
                        try
                        {
                            string filename = Path.GetFileName(allFiles[i]);
                            File.Copy(allFiles[i], Path.Combine(destFullPath, filename), true);
                        }
                        catch (Exception ex)
                        {
                            errorMsg = errorMsg + "-->" + ex.Message + "<-->" + ex.StackTrace;
                            //Your error handler here
                            //MessageBox.Show(ex.Message + System.Environment.NewLine + ex.StackTrace);
                            success = false;
                        }
                    }
                }
            }
            return success;
        }

        public static bool DeleteFilesWithWildcards(string fullPath, ref string errorMsg)
        {
            bool success = true;
            //This if statement is safe because
            //Files can not be named with the following characters:
            //\ / : ? <> |and,most importantly, *
            if (fullPath.IndexOf("*") != -1)
            {
                //This is a wild card delete;
                //Seperating the Directory Path and the Filename
                int directoryEnd = fullPath.LastIndexOf("\\");
                string directoryPath = fullPath.Substring(0, directoryEnd + 1);
                string fileName = fullPath.Substring(directoryEnd + 1, fullPath.Length - (directoryEnd + 1));

                //Directory.GetFiles will return anarray of strings that contain the full file path.
                string[] allFiles = Directory.GetFiles(directoryPath, fileName);
                if (allFiles.Length == 0)
                    //There are no files matching-- so hey -- success! The files are 
                    //not present in the directory, which was our intended result!
                    success = true;
                else
                {
                    success = true;
                    for (int i = 0; i < allFiles.Length; i++)
                    {
                        try
                        {
                            File.Delete(allFiles[i]);
                        }
                        catch (Exception ex)
                        {
                            errorMsg = errorMsg + "-->" + ex.Message + "<-->" + ex.StackTrace;
                            //Your error handler here
                            //MessageBox.Show(ex.Message + System.Environment.NewLine + ex.StackTrace);
                            success = false;
                        }
                    }
                }
            }
            return success;
        }

        public static bool CheckSearchTextExistInFile(string checkfile, string searchText, ref string msg)
        {
            bool status = false;
            int MyPos = 1;
            Int16 iCount = 0;
            if (status == false)
            {
                MyPos = 1;
                iCount = 0;
                try
                {
                    StreamReader reader2 = File.OpenText(checkfile);
                    string strFile = reader2.ReadToEnd();
                    do
                    {
                        MyPos = strFile.IndexOf(searchText, MyPos + 1);
                        if (MyPos > 0)
                        {
                            iCount += 1;
                        }
                    } while (!(MyPos == -1));

                    if (iCount == 0)
                        status = false;
                    else
                    {
                        status = true;
                        //Log("Number of '0 errors(0)' found: iCount = " + iCount);
                    }
                }
                catch (Exception ex)
                {
                    msg = msg + "- CheckSearchTextExistInFile:" + ex.Message + " ------ " + ex.StackTrace;
                    //Log(ex.Message + "----" + ex.StackTrace);
                    status = false;
                }
            }
            return status;
        }

        public static bool CheckSearchTextExistInFile(string path, string fileName, string searchText, ref string msg)
        {
            bool status = false;
            int MyPos = 1;
            Int16 iCount = 0;
            if (status == false)
            {
                MyPos = 1;
                iCount = 0;
                try
                {
                    string fileFullname = Path.Combine(path, fileName);
                    StreamReader reader2 = File.OpenText(fileFullname);
                    string strFile = reader2.ReadToEnd();
                    do
                    {
                        MyPos = strFile.IndexOf(searchText, MyPos + 1);
                        if (MyPos > 0)
                        {
                            iCount += 1;
                        }
                    } while (!(MyPos == -1));

                    if (iCount == 0)
                        status = false;
                    else
                    {
                        status = true;
                        //Log("Number of '0 errors(0)' found: iCount = " + iCount);
                    }
                }
                catch (Exception ex)
                {
                    msg = msg + "- CheckSearchTextExistInFile:" + ex.Message + " ------ " + ex.StackTrace;
                    //Log(ex.Message + "----" + ex.StackTrace);
                    status = false;
                }
            }
            return status;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path">testZorking.txt file path</param>
        /// <param name="status">true or false</param>
        public static void UpdateTestWorkingFile(string path, string status)
        {
            string testWorkingTxtFile = Path.Combine(path, ConstCommon.TESTWORKING_FILENAME);
            var file = new FileInfo(testWorkingTxtFile);
            File.SetAttributes(file.FullName, FileAttributes.Normal);

            StreamWriter writeWorking = File.CreateText(testWorkingTxtFile);
            writeWorking.WriteLine(status);
            writeWorking.Close();
        }

        public static void CopyDirectory(string Src, string Dst)
        {
            String[] Files;

            if (Dst[Dst.Length - 1] != Path.DirectorySeparatorChar)
                Dst += Path.DirectorySeparatorChar;

            if (!Directory.Exists(Dst))
                Directory.CreateDirectory(Dst);

            Files = Directory.GetFileSystemEntries(Src);
            foreach (string Element in Files)
            {
                // Sub directories
                if (Directory.Exists(Element))
                    CopyDirectory(Element, Dst + Path.GetFileName(Element));
                    // Files in directory      
                else
                    File.Copy(Element, Dst + Path.GetFileName(Element), true);
            }
        }

        //Excel
        public static void WriteExcelHeader(ref Application xApp, string sExcelVisible, string[] HeaderLines)
        {
            xApp = new Application();
            Workbooks xBooks = xApp.Workbooks;
            Workbook xBook = xBooks.Add(Type.Missing);
            //xSheet = (Excel.Worksheet)xBook.Worksheets[1];
            dynamic xSheet = xApp.ActiveSheet;
            if (sExcelVisible == string.Empty)
                xApp.Visible = true;
            else
            {
                if (sExcelVisible.StartsWith("Visible"))
                    xApp.Visible = true;
                else
                    xApp.Visible = false;
            }

            xApp.Interactive = true;
            for (int i = 0; i < HeaderLines.Length; i++)
            {
                xSheet.Cells[i + 1, 1] = HeaderLines[i].Substring(0, HeaderLines[i].IndexOf('*'));
                xSheet.Cells[i + 1, 2] = HeaderLines[i].Substring(HeaderLines[i].IndexOf('*') + 1);

                //dynamic xRange = xSheet.get_Range("A" + (i + 1), "A" + (i + 1));
                //xRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                //xRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            }
        }

        public static void WriteExcelHeader(ref Application xApp, string AppName, string sExcelVisible, string PCName,
                                            string OSName, string OSVersion,
                                            string UICulture, string TimeOnPC, string sTestToolsVersion,
                                            string networkMap, string sInstallMsiDir)
        {
            xApp = new Application();
            Workbooks xBooks = xApp.Workbooks;
            Workbook xBook = xBooks.Add(Type.Missing);
            //xSheet = (Excel.Worksheet)xBook.Worksheets[1];
            dynamic xSheet = xApp.ActiveSheet;
            if (sExcelVisible == string.Empty)
                xApp.Visible = true;
            else
            {
                if (sExcelVisible.StartsWith("Visible"))
                    xApp.Visible = true;
                else
                    xApp.Visible = false;
            }

            xApp.Interactive = true;
            string today = DateTime.Now.ToString("MMMM-dd");
            xSheet.Cells[1, 1] = today;
            xSheet.Cells[1, 2] = AppName + " UI Test Scenarios";
            xSheet.Cells[2, 1] = "Test Machine:";
            xSheet.Cells[2, 2] = PCName;
            xSheet.Cells[3, 1] = "Tester:";
            xSheet.Cells[3, 2] = WindowsIdentity.GetCurrent().Name;
            xSheet.Cells[4, 1] = "OSName:";
            xSheet.Cells[4, 2] = OSName;
            xSheet.Cells[5, 1] = "OS Version:";
            xSheet.Cells[5, 2] = OSVersion;
            xSheet.Cells[6, 1] = "UI Culture";
            xSheet.Cells[6, 2] = UICulture;
            xSheet.Cells[7, 1] = "Time On PC";
            xSheet.Cells[7, 2] = "local time:" + TimeOnPC;
            xSheet.Cells[8, 1] = "Test Tool Version:";
            xSheet.Cells[8, 2] = sTestToolsVersion;
            xSheet.Cells[9, 1] = "NetworkMap:";
            xSheet.Cells[9, 2] = networkMap;
            if (sInstallMsiDir != null)
            {
                xSheet.Cells[10, 1] = "Build Location:";
                xSheet.Cells[10, 2] = sInstallMsiDir;
            }
        }


        public static void WriteExcelTestCaseResult(Application xApp, int result, int Counter, string name,
                                                    string errorMSG)
        {
            Worksheet xSheet = xApp.ActiveSheet;
            string time = DateTime.Now.ToString("HH:mm:ss");
            xSheet.Cells[Counter + 11, 1] = time;
            xSheet.Cells[Counter + 11, 2] = name;
            xSheet.Cells[Counter + 11, 3] = errorMSG;

            //Excel.Range 
            dynamic xRange = xSheet.get_Range("A" + (Counter + 11), "A" + (Counter + 11));
            xRange.VerticalAlignment = XlVAlign.xlVAlignTop;
            xRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;

            xRange = xSheet.get_Range("C" + (Counter + 11), "C" + (Counter + 11));
            xRange.VerticalAlignment = XlVAlign.xlVAlignTop;

            xRange = xSheet.get_Range("B" + (Counter + 11), "B" + (Counter + 11));
            xRange.VerticalAlignment = XlVAlign.xlVAlignTop;

            switch (result)
            {
                case ConstCommon.TEST_PASS:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_GREEN;
                    break;
                case ConstCommon.TEST_FAIL:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_RED;
                    break;
                case ConstCommon.TEST_EXCEPTION:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_PINK;
                    break;
                case ConstCommon.TEST_UNDEFINED:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_YELLOW;
                    //xSheet.Cells.set_Item(row, 3, testData);
                    break;
            }
        }


        public static void WriteExcelTestCaseResult(Application xApp, int result, int headerRowCount, int Counter,
                                                    string name, string errorMSG)
        {
            int rowNumber = headerRowCount + 1 + Counter;
            Worksheet xSheet = xApp.ActiveSheet;
            string time = DateTime.Now.ToString("HH:mm:ss");
            xSheet.Cells[rowNumber, 1] = time;
            xSheet.Cells[rowNumber, 2] = name;
            xSheet.Cells[rowNumber, 3] = errorMSG;

            dynamic xRange = xSheet.get_Range("A" + (rowNumber), "A" + (rowNumber));
            xRange.VerticalAlignment = XlVAlign.xlVAlignTop;

            xRange = xSheet.get_Range("C" + (rowNumber), "C" + (rowNumber));
            xRange.VerticalAlignment = XlVAlign.xlVAlignTop;

            xRange = xSheet.get_Range("B" + (rowNumber), "B" + (rowNumber));
            xRange.VerticalAlignment = XlVAlign.xlVAlignTop;

            switch (result)
            {
                case ConstCommon.TEST_PASS:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_GREEN;
                    break;
                case ConstCommon.TEST_FAIL:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_RED;
                    break;
                case ConstCommon.TEST_EXCEPTION:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_PINK;
                    break;
                case ConstCommon.TEST_UNDEFINED:
                    xRange.Interior.ColorIndex = ConstCommon.EXCEL_YELLOW;
                    //xSheet.Cells.set_Item(row, 3, testData);
                    break;
            }
        }

        public static void WriteExcelFoot(Application xApp, int Counter, int sTotalCounter, int sTotalPassed,
                                          int sTotalFailed)
        {
            Worksheet xSheet = xApp.ActiveSheet;
            xSheet.Cells[Counter + 2 + 10, 1] = "Total tests: ";
            xSheet.Cells[Counter + 3 + 10, 1] = "Total Passes: ";
            xSheet.Cells[Counter + 4 + 10, 1] = "Total Failed: ";

            xSheet.Cells[Counter + 2 + 10, 2] = sTotalCounter;
            xSheet.Cells[Counter + 3 + 10, 2] = sTotalPassed;
            xSheet.Cells[Counter + 4 + 10, 2] = sTotalFailed;

            ulong TPhysicalMem = 0;
            ulong APhysicalMem = 0;
            ulong TVirtualMem = 0;
            ulong AVirtualMem = 0;

            HelpUtilities.GetMemoryInfo(out TPhysicalMem, out APhysicalMem, out TVirtualMem, out AVirtualMem);
            // Add Legende
            xSheet.Cells[Counter + 5 + 10, 2] = "Legende";
            dynamic xRange = xApp.get_Range("B" + (Counter + 5 + 10));
            xRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            xSheet.Cells[Counter + 6 + 10, 2] = "Pass";
            xRange = xApp.get_Range("B" + (Counter + 6 + 10));
            xRange.Interior.ColorIndex = ConstCommon.EXCEL_GREEN;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            xSheet.Cells[Counter + 7 + 10, 2] = "Fail";
            xRange = xApp.get_Range("B" + (Counter + 7 + 10));
            xRange.Interior.ColorIndex = ConstCommon.EXCEL_RED;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            xSheet.Cells[Counter + 8 + 10, 2] = "Exception";
            xRange = xApp.get_Range("B" + (Counter + 8 + 10));
            xRange.Interior.ColorIndex = ConstCommon.EXCEL_PINK;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            xSheet.Cells[Counter + 9 + 10, 2] = "Untested";
            xRange = xApp.get_Range("B" + (Counter + 9 + 10));
            xRange.Interior.ColorIndex = ConstCommon.EXCEL_YELLOW;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            xSheet.Cells[Counter + 10 + 11, 2] = "TotalPhysicalMemory:" + TPhysicalMem + " MB";
            xSheet.Cells[Counter + 11 + 11, 2] = "AvailablePhysicalMemory:" + APhysicalMem + " MB";
            xSheet.Cells[Counter + 12 + 11, 2] = "TotalVirtualMemory:" + TVirtualMem + " MB";
            xSheet.Cells[Counter + 13 + 11, 2] = "AvailableVirtualMemory:" + AVirtualMem + " MB";

            xSheet.Columns.AutoFit();
            xSheet.Rows.AutoFit();
        }

        public static void WriteExcelFoot(Application xApp, int headerRowCount, int Counter, int sTotalCounter,
                                          int sTotalPassed, int sTotalFailed)
        {
            int beginRowNumber = headerRowCount + sTotalCounter + 2;
            Worksheet xSheet = xApp.ActiveSheet;
            xSheet.Cells[beginRowNumber + 1, 1] = "Total tests: ";
            xSheet.Cells[beginRowNumber + 2, 1] = "Total Passes: ";
            xSheet.Cells[beginRowNumber + 3, 1] = "Total Failed: ";

            xSheet.Cells[beginRowNumber + 1, 2] = sTotalCounter;
            xSheet.Cells[beginRowNumber + 2, 2] = sTotalPassed;
            xSheet.Cells[beginRowNumber + 3, 2] = sTotalFailed;

            ulong TPhysicalMem = 0;
            ulong APhysicalMem = 0;
            ulong TVirtualMem = 0;
            ulong AVirtualMem = 0;

            HelpUtilities.GetMemoryInfo(out TPhysicalMem, out APhysicalMem, out TVirtualMem, out AVirtualMem);
            // Add Legende
            xSheet.Cells[beginRowNumber + 5, 2] = "Legende";
            dynamic xRange = xApp.get_Range("B" + (beginRowNumber + 5));
            xRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            xSheet.Cells[beginRowNumber + 6, 2] = "Pass";
            xRange = xApp.get_Range("B" + (beginRowNumber + 6));
            xRange.Interior.ColorIndex = ConstCommon.EXCEL_GREEN;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            xSheet.Cells[beginRowNumber + 7, 2] = "Fail";
            xRange = xApp.get_Range("B" + (beginRowNumber + 7));
            xRange.Interior.ColorIndex = ConstCommon.EXCEL_RED;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            xSheet.Cells[beginRowNumber + 8, 2] = "Exception";
            xRange = xApp.get_Range("B" + (beginRowNumber + 8));
            xRange.Interior.ColorIndex = ConstCommon.EXCEL_PINK;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            xSheet.Cells[beginRowNumber + 9, 2] = "Untested";
            xRange = xApp.get_Range("B" + (beginRowNumber + 9));
            xRange.Interior.ColorIndex = ConstCommon.EXCEL_YELLOW;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            xSheet.Cells[beginRowNumber + 11, 2] = "TotalPhysicalMemory:" + TPhysicalMem + " MB";
            xSheet.Cells[beginRowNumber + 12, 2] = "AvailablePhysicalMemory:" + APhysicalMem + " MB";
            xSheet.Cells[beginRowNumber + 13, 2] = "TotalVirtualMemory:" + TVirtualMem + " MB";
            xSheet.Cells[beginRowNumber + 14, 2] = "AvailableVirtualMemory:" + AVirtualMem + " MB";

            xSheet.Columns.AutoFit();
            xSheet.Rows.AutoFit();
        }

        public static bool SaveExcel(Application xApp, string sXLSPath, ref string errorMsg)
        {
            bool saveOK = true;
            object missing = Missing.Value;
            Console.WriteLine("Save Excel File : " + sXLSPath);
            dynamic xBook = xApp.ActiveWorkbook;
            //Epia3Common.WriteTestLogMsg(slogFilePath, "Save1 to local machine : " + sXLSPath, sOnlyUITest);
            try
            {
                xBook.SaveAs(sXLSPath, XlFileFormat.xlWorkbookNormal,
                             missing, missing, missing, missing,
                             XlSaveAsAccessMode.xlNoChange,
                             missing, missing, missing, missing, missing);
            }
            catch (Exception ex)
            {
                saveOK = false;
                //string sTXTPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), sOutFilename + ".txt");
                errorMsg = "Save Excel file1 exception: " + ex;
                //StreamWriter write = File.CreateText(sTXTPath);
                //write.WriteLine(writeMsg);
                //Epia3Common.WriteTestLogMsg(slogFilePath, writeMsg, sOnlyUITest);
                //write.Close();
            }
            return saveOK;
        }
    }
}