using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;

using System.Threading;
using System.Windows.Automation;
using TestTools;
using System.Windows.Forms;
using System.Windows;
using System.IO;

using System.Drawing;
using System.Data.SqlClient;

namespace TestSimulators
{
    public class ScannerSimulators
    {
        internal StringCollection m_Logging = new StringCollection();

        // PCinfo
        static public string PCName;
        static public string OSName;
        static public string OSVersion;
        static public string UICulture;
        static public string TimeOnPC;
        static private string sConnectionString = string.Empty;  

        internal STATE m_State;
        string sErrorMessage = string.Empty;
        public static string slogFilePath = @"C:\KC\PutAway\";
        static string testInfoTxtFile;

        internal bool CheckDropLocEmpty = false;

        internal bool foundTPSelectScreen = false;
        internal bool foundOption = false;
        System.Windows.Point AT48Pt = new System.Windows.Point(0,0);
        System.Windows.Point OptionPt = new System.Windows.Point(0,0);

        // common 
        static internal AutomationElement aeForm = null; 
        static internal AutomationElement aeOption = null;
        static internal AutomationElement aeSelectTPScreen = null;
        static internal AutomationElement aePickTPScreen = null;
        static internal AutomationElement aeDropTPScreen = null;
        static internal AutomationElement aeCancelButton = null;
        static internal AutomationElement aeCarrier = null;

        static internal AutomationElement aeSrc = null;
        static internal AutomationElement aeDest = null;

        internal const string InstructionId = "m_LblInstruction";
        static internal AutomationElement aeInstructionLable = null;

        internal const string SelectTPId      = "SelectTransportScreen";
        internal const string PickTPSreenId   = "m_TblContent";
        internal const string DropTPSreenId   = "m_TblContent";
        internal const string ConsumeReelId   = "m_TblContent"; 
        internal const string ReelUnitId      = "m_LblUnitId";
        static internal string optionId = "m_TblOptions";

        static internal string TrUnitId = string.Empty;
        static internal string CarrierId = string.Empty;
        // pronglift
        static internal string SrcLoc = string.Empty;
        static internal string DestLoc = string.Empty;

        internal int sDrop2TCount = 0;
        internal int sDrop7TCount = 0;
        internal int sDropFTCount = 0;  
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Events of TestSimulators (1)
        public event EventHandler OnLoggingChanged;
        #endregion // —— Events •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Properties of TestSimulators (2)
        public StringCollection Logging
        {
            get
            {
                return m_Logging;
            }
        }

        public STATE State
        {
            get
            {
                return m_State;
            }
            set
            {
                m_State = value;
            }
        }
        #endregion // —— Properties •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

         // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Constructors/Destructors/Cleanup of Tester (1)
        public ScannerSimulators()
        {
            try
            {
                HelpUtilities.SavePCInfo("y");
                HelpUtilities.GetPCInfo(out PCName, out OSName, out OSVersion, out UICulture, out TimeOnPC);
                Console.WriteLine("PCName : " + PCName);
                Console.WriteLine("OSName : " + OSName);
                Console.WriteLine("OSVersion : " + OSVersion);
                Console.WriteLine("UICulture : " + UICulture);
                Console.WriteLine("TimeOnPC : " + TimeOnPC);

                sConnectionString =
                         "Integrated Security=SSPI;"
                         + "Persist Security Info=False;"
                         + "Initial Catalog=Ewcs;"
                         + "Data Source=" + PCName;
            }
            catch (Exception ex)
            {
                MessageBox.Show("GetPCInfo:" + ex.Message);
            }

            m_State = STATE.PENDING;
        }
        #endregion // —— Constructors/Destructors/Cleanup •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        //---------------------------------------------------------------------
        internal void Log(string Message)
        {
            Message = System.String.Format("{0:G}: {1}.", System.DateTime.Now, " - " + Message);
            m_Logging.Add(Message);
            if (OnLoggingChanged != null)
                OnLoggingChanged(this, new EventArgs());

            // set max log line is 100
            while (m_Logging.Count > 35)
            {
                m_Logging.RemoveAt(0);
            }
        }

        //=====================================================================
        /// <summary>
        /// Method will start new tests
        /// </summary>
        public void ProngliftScannerStart(string source, string destination)
        {
            string slogMsg = "ProngliftScannerStart : " + System.DateTime.Now;
            Log(slogMsg);
            WriteLog(testInfoTxtFile, slogMsg);

            testInfoTxtFile = Path.Combine(@"C:\KC\PutAway", "ProngLift.log");
            StreamWriter writeInfo = File.CreateText(testInfoTxtFile);
            string info = "Start test : " + DateTime.Now;
            writeInfo.WriteLine(info);
            writeInfo.Close();
            sDrop7TCount = 0;

            DateTime mAppTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mAppTime;

            while (true)
            {
                #region // A find mainform
                mAppTime = DateTime.Now;
                mTime = DateTime.Now - mAppTime;
                slogMsg = "(A)<-- Find Application aeForm : " + System.DateTime.Now;
                Log(slogMsg);
                WriteLog(testInfoTxtFile, slogMsg);
                
                while (aeForm == null && mTime.Minutes < 10)
                {
                    aeForm = AUIUtilities.FindElementByID("MainForm", AutomationElement.RootElement);
                    mTime = DateTime.Now - mAppTime;
                    Thread.Sleep(5000);
                }

                // if after 5 minutes still no mainform,throw exception 
                if (aeForm == null)
                {
                    AutomationElement aeError = AUIUtilities.FindElementByID("ErrorScreen", AutomationElement.RootElement);
                    if (aeError != null)
                        AUICommon.ErrorWindowHandling(aeError, ref sErrorMessage);
                    else
                        sErrorMessage = "Application Startup failed,see logging";

                    throw new Exception(sErrorMessage);
                }
                else
                {
                    Console.WriteLine("Application maeForm name : " + aeForm.Current.Name);
                    Log("Application maeForm name : " + aeForm.Current.Name + " - Time: " + System.DateTime.Now);
                }

                slogMsg = "A.0 --> MainForm founded : " + System.DateTime.Now + " -----------------";
                Log(slogMsg);
                WriteLog(testInfoTxtFile, slogMsg);
                #endregion

                #region  // B Handle transport screen display
                string instructionName = string.Empty;
                AutomationElement aeBtnHome = null;

                mAppTime = DateTime.Now;
                mTime = DateTime.Now - mAppTime;
                aeInstructionLable = null;

                while (aeInstructionLable == null && mTime.Minutes < 5)
                {
                    aeInstructionLable = AUIUtilities.FindElementByID(InstructionId, aeForm);
                    mTime = DateTime.Now - mAppTime;
                    Thread.Sleep(1000);
                }
                if (aeInstructionLable != null)
                {
                    instructionName = aeInstructionLable.Current.Name;
                    slogMsg = " *** (B)--> Handle instruction screens : " + instructionName;
                    Log(slogMsg);
                    WriteLog(testInfoTxtFile, slogMsg);

                    if (instructionName.ToLower().StartsWith("waiting for transports"))
                    {
                        #region // handle waiting for transport instruction
                        aeBtnHome = AUIUtilities.FindElementByID("m_BtnHome", aeForm);
                        if (aeBtnHome != null)
                        {
                            Thread.Sleep(2000);
                            // sometime screen change to other select screen after 'waiting for...'
                            if (aeBtnHome.Current.IsEnabled)
                                Input.MoveToAndClick(aeBtnHome);
                            else
                            {
                                slogMsg = " ***(B.X )--> screens changed during, retry: " + instructionName;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        else
                        {
                            sErrorMessage = "Buttom Home not found:";
                            slogMsg = " ***(B.2 )--> Buttom Home not found:, retry maybe screen already changed : ";
                            Log(slogMsg);
                            WriteLog(testInfoTxtFile, slogMsg);
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("what do you want to do"))
                    {
                        #region // handle what do you want to do screen
                        // click first option  
                        slogMsg = " ***** (B.1)--> Click first option now : ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        aeOption = AUIUtilities.FindElementByID(optionId, aeForm);
                        if (aeOption != null)
                        {
                            Thread.Sleep(1000);

                            // Set a property condition that will be used to find the control.
                            System.Windows.Automation.Condition c = new PropertyCondition(
                                AutomationElement.ControlTypeProperty, ControlType.Button);

                            AutomationElementCollection aeOptionButton = aeOption.FindAll(TreeScope.Element | TreeScope.Descendants, c);
                            Thread.Sleep(1000);

                            OptionPt = AUIUtilities.GetElementCenterPoint(aeOptionButton[0]);
                            Thread.Sleep(1000);
                            Input.MoveTo(OptionPt);

                            WriteLog(testInfoTxtFile, "numOption  : " + 0);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(OptionPt);
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("select a transport"))
                    {
                        #region // handle select a transport screen
                        slogMsg = "(S)--> Selection instruction screen found : ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        #region  // Assign Transport
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        aeSelectTPScreen = null;
                        while (aeSelectTPScreen == null && mTime.Minutes < 10)
                        {
                            Console.WriteLine("Find SelectTransportScreen aeSelectTPScreen : " + System.DateTime.Now);
                            aeSelectTPScreen = AUIUtilities.FindElementByID(SelectTPId, aeForm);
                            Console.WriteLine("SelectTransportScreen aeSelectTPScreen: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            WriteLog(testInfoTxtFile, "Select screen find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(500);
                        }
                        if (aeSelectTPScreen == null)
                        {
                            sErrorMessage = "SelectTransportScreen not found";
                            Console.WriteLine("FindElementByID failed:" + SelectTPId);
                            MessageBox.Show("No new transport exist any more", "Find select a transport screen");
                            Thread.Sleep(3600000 * 20);
                        }
                        else
                        {
                            Console.WriteLine("SelectTPScreen found, now find select button");
                            Thread.Sleep(500);

                            // find reel unit id
                            /*AutomationElement aeUnit = AUIUtilities.FindElementByID(ReelUnitId, aeSelectTPScreen);
                            if (aeUnit == null)
                            {
                                sErrorMessage = "UnitID not found";
                                Console.WriteLine("UnitId not found:" + ReelUnitId);
                                MessageBox.Show(sErrorMessage, "Assign Transport");
                                Thread.Sleep(3600000 * 20);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                TrUnitId = aeUnit.Current.Name;
                                Log("Select transport with unit: " + TrUnitId);
                                WriteLog(testInfoTxtFile, "Select transport with unit: " + TrUnitId);
                            }
                            */

                            // find source loc
                            aeSrc = null;
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeSrc == null && mTime.Minutes < 2)
                            {
                                aeSrc = AUIUtilities.FindElementByID("m_LblSourceLocation", aeSelectTPScreen);
                                mTime = DateTime.Now - mAppTime;
                                WriteLog(testInfoTxtFile, "sourceLoc find time is :" + mTime.TotalMilliseconds / 1000 + "  m_LblSourceLocation");
                                Thread.Sleep(500);
                            }

                            if (aeSrc == null)
                            {
                                sErrorMessage = "Source not found";
                                Console.WriteLine("Source not found:" + "m_LblSourceLocation");
                                MessageBox.Show(sErrorMessage, instructionName);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                SrcLoc = aeSrc.Current.Name;
                                slogMsg = "(S.1)--> transport source is: " + SrcLoc;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }

                            // find destLocation id
                            aeDest = null;
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeDest == null && mTime.Minutes < 2)
                            {
                                aeDest = AUIUtilities.FindElementByID("m_LblDestinationLocation", aeSelectTPScreen);
                                mTime = DateTime.Now - mAppTime;
                                WriteLog(testInfoTxtFile, "destination find time is :" + mTime.TotalMilliseconds / 1000 + "  m_LblDestinationLocation");
                                Thread.Sleep(500);
                            }

                            if (aeDest == null)
                            {
                                sErrorMessage = "dest ID not found";
                                Console.WriteLine("dest Id not found:" + "m_LblDestinationLocation");
                                MessageBox.Show(sErrorMessage, instructionName);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                DestLoc = aeDest.Current.Name;
                                slogMsg = "(S.2)--> Selected transport Destination is: " + DestLoc;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }

                            Thread.Sleep(1000);
                            string SelectId = "m_BtnOK";
                            AutomationElement aeSelectButton = AUIUtilities.FindElementByID(SelectId, aeSelectTPScreen);
                            if (aeSelectButton == null)
                            {
                                slogMsg = "(S.X)--> Selection confirm button not found: " + SelectId;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                            else
                            {
                                Thread.Sleep(500);
                                Input.MoveTo(aeSelectButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSelectButton));
                                slogMsg = "(S.OK)--> Transport is assigned: OK";
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("pick"))
                    {
                        #region // handle pick and scan the pallet or reel screen
                        slogMsg = "(P)--> Handle pick screen : " + instructionName;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        // Pick transport
                        string PickTPSreenId = "m_TblContent";
                        string PickScanId = "m_TxtScannedValue";  // edit control
                        string PickOKId = "m_BtnOK";

                        // Find the pick transport screen  SelectTransportScreen
                        Thread.Sleep(1000);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        AutomationElement aePickTPScreen = null;
                        while (aePickTPScreen == null && mTime.Minutes < 10)
                        {
                            Console.WriteLine("Find PickTransportScreen aeSelectTPScreen : " + System.DateTime.Now);
                            aePickTPScreen = AUIUtilities.FindElementByID(PickTPSreenId, aeForm);
                            Console.WriteLine("PickTransportScreen aePickTPScreen: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(500);
                        }
                        if (aePickTPScreen == null)
                        {
                            sErrorMessage = "PickTransportScreen not found";
                            Console.WriteLine("Pick FindElementByID failed:" + PickTPSreenId);
                            MessageBox.Show("No Pick transport screen exist", "Find pick a transport screen");
                            Thread.Sleep(3600000 * 20);
                        }
                        else
                        {
                            Console.WriteLine("PickTPScreen found, now find reelLd, scan and select button");
                            Thread.Sleep(1000);
                            // Scan location  
                            // "m_TxtScannedValue"  This is Edit Control, should use setFocus + sendKeys
                            if (FindDocumentAndSendText(PickScanId, aePickTPScreen, SrcLoc, ref sErrorMessage))
                            {
                                slogMsg = "(P.1)--> " + SrcLoc + "  scan OK : ";
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                Log("found AND scan field failed::" + sErrorMessage);
                                MessageBox.Show(sErrorMessage, " Pick Transport");
                            }

                            Thread.Sleep(2000);
                            // OK 
                            AutomationElement aePickOKButton = AUIUtilities.FindElementByID(PickOKId, aePickTPScreen);
                            if (aePickOKButton == null)
                            {
                                sErrorMessage = "Select button not found";
                                MessageBox.Show(sErrorMessage, " Pick Transport");
                                Thread.Sleep(1000000000);
                            }
                            else
                            {
                                Thread.Sleep(500);
                                Input.MoveTo(aePickOKButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aePickOKButton));

                                slogMsg = "(P.2)--> " + " Pick transport confirmed : ";
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("drop"))
                    {
                        #region  // Handle drop the reel at the destination location
                        slogMsg = "(D)--> Drop location is --> " + DestLoc;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        // Drop transport
                        string ContentId = "m_TblContent";
                        string DropOKId = "m_BtnOK";

                        #region  // check Drop loc empty  region
                        if (CheckDropLocEmpty && DestLoc.ToLower().StartsWith("hu"))
                        {
                            // Wait until FIL input empty
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            bool FilInputEmpty = false;
                            while (FilInputEmpty == false)
                            {

                                if (IsDropLocationEmpty("HU.1"))
                                {
                                    FilInputEmpty = true;
                                    slogMsg = "(D.1)-->  : " + DestLoc + " ::HU.1--> empty, " + System.DateTime.Now;

                                }
                                else
                                {
                                    FilInputEmpty = false;
                                    mTime = DateTime.Now - mAppTime;
                                    slogMsg = "(D.1)-->  : " + DestLoc + " ::HU.1--> not empty," + System.DateTime.Now;
                                }
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                                Thread.Sleep(5000);
                            }
                        }
                        #endregion

                        // Find the  transport screen  DropTransportScreen
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;

                        AutomationElement aeDropTPScreen = null;
                        string DropScanId = "m_TxtScannedValue";  // edit control
                        AutomationElement aeDropField = null;
                        AutomationElement aeTblContent = null;
                        while (aeDropTPScreen == null && mTime.Minutes < 5)
                        {
                            Console.WriteLine("Find DropTransportScreen : " + System.DateTime.Now);
                            aeDropTPScreen = AUIUtilities.FindElementByID(DropTPSreenId, aeForm);
                            Console.WriteLine("DropTransportScreen found: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(3000);
                        if (aeDropTPScreen == null)
                        {
                            sErrorMessage = "DropTransportScreen not found";
                            Console.WriteLine("Drop FindElementByID failed:" + DropTPSreenId);
                        }
                        else
                        {
                            WriteLog(testInfoTxtFile, "DropTPScreen found, now find TblContent, scan and select button");
                            // find tblContent
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeTblContent == null && mTime.Minutes < 10)
                            {
                                aeTblContent = AUIUtilities.FindElementByID(ContentId, aeDropTPScreen);
                                mTime = DateTime.Now - mAppTime;
                                WriteLog(testInfoTxtFile, "Tbl Content  find time is :" + mTime.TotalMilliseconds / 1000);
                                Thread.Sleep(2000);
                            }

                            if (aeTblContent == null)
                            {
                                Log(sErrorMessage);
                                WriteLog(testInfoTxtFile, "Drop TblContent not found:");
                                MessageBox.Show(sErrorMessage, "Drop");
                                Thread.Sleep(3600000 * 20);
                            }

                            // Scan location  
                            // "m_TxtScannedValue"  This is Edit Control, should use setFocus + sendKeys
                            //if (AUIUtilities.FindDocumentAndSendText(DropScanId, aeDropTPScreen, "HU.1", ref sErrorMessage))
                            if (aeTblContent.Current.Name.StartsWith("Enter"))
                            {
                                mAppTime = DateTime.Now;
                                mTime = DateTime.Now - mAppTime;
                                while (aeDropField == null && mTime.Minutes < 3)
                                {
                                    aeDropField = AUIUtilities.FindElementByID(DropScanId, aeDropTPScreen);
                                    mTime = DateTime.Now - mAppTime;
                                    WriteLog(testInfoTxtFile, "drop field  find time is :" + mTime.TotalMilliseconds / 1000);
                                    Thread.Sleep(2000);
                                }
                                //aeDropField = AUIUtilities.FindElementByID(DropScanId, aeDropTPScreen);
                                string destID = DestLoc;
                                if (aeDropField != null)
                                {
                                    if (SrcLoc.Equals("KCP5.2"))
                                        destID = "KCP5.1";
                                    else if (SrcLoc.Equals("KCP5.2R"))
                                        destID = "KCP5.1R";

                                    Thread.Sleep(3000);
                                    aeDropField.SetFocus();
                                    Thread.Sleep(1000);
                                    System.Windows.Forms.SendKeys.SendWait(destID);
                                    Thread.Sleep(2000);
                                    //ValuePattern vp = (ValuePattern)aeDropField.GetCurrentPattern(ValuePattern.Pattern);
                                    //Thread.Sleep(2000);
                                    //vp.SetValue("HU.1");
                                    Console.WriteLine(destID+" scanned");
                                    Log(destID+"  OK");
                                }
                                else
                                {
                                    Log(sErrorMessage);
                                    WriteLog(testInfoTxtFile, "Drop ScanID not found:");
                                    MessageBox.Show(sErrorMessage, "Drop to " + destID);
                                }
                            }
                            Thread.Sleep(500);

                            // OK 
                            AutomationElement aeDropOKButton = AUIUtilities.FindElementByID(DropOKId, aeDropTPScreen);
                            if (aeDropOKButton == null)
                            {
                                sErrorMessage = "Drop OK button not found";
                                Console.WriteLine("Drop OK button not found:" + DropOKId);
                                Log("Drop OK button not found:" + DropOKId);
                            }
                            else
                            {
                                sDrop7TCount++;
                                Thread.Sleep(500);
                                Input.MoveTo(aeDropOKButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeDropOKButton));
                                slogMsg = "(D.OK)--> --> --> --> Total Drop :" + sDrop7TCount;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("enter"))
                    {
                        #region  // Handle enter the location where the reel has been dropped
                        slogMsg = "(E)--> Enter location is --> " + DestLoc;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        // Drop transport
                        string ContentId = "m_TblContent";
                        string DropOKId = "m_BtnOK";

                        #region  // check Drop loc empty  region
                        if (CheckDropLocEmpty && DestLoc.ToLower().StartsWith("hu"))
                        {
                            // Wait until FIL input empty
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            bool FilInputEmpty = false;
                            while (FilInputEmpty == false)
                            {

                                if (IsDropLocationEmpty("HU.1"))
                                {
                                    FilInputEmpty = true;
                                    slogMsg = "(E.1)-->  : " + DestLoc + " ::HU.1--> empty, " + System.DateTime.Now;

                                }
                                else
                                {
                                    FilInputEmpty = false;
                                    mTime = DateTime.Now - mAppTime;
                                    slogMsg = "(E.1)-->  : " + DestLoc + " ::HU.1--> not empty," + System.DateTime.Now;
                                }
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                                Thread.Sleep(5000);
                            }
                        }
                        #endregion

                        // Find the  transport screen  DropTransportScreen
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;

                        AutomationElement aeDropTPScreen = null;
                        string DropScanId = "m_TxtScannedValue";  // edit control
                        AutomationElement aeDropField = null;
                        AutomationElement aeTblContent = null;
                        while (aeDropTPScreen == null && mTime.Minutes < 5)
                        {
                            Console.WriteLine("Find DropTransportScreen : " + System.DateTime.Now);
                            aeDropTPScreen = AUIUtilities.FindElementByID(DropTPSreenId, aeForm);
                            Console.WriteLine("DropTransportScreen found: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(3000);
                        if (aeDropTPScreen == null)
                        {
                            sErrorMessage = "DropTransportScreen not found";
                            Console.WriteLine("Drop FindElementByID failed:" + DropTPSreenId);
                        }
                        else
                        {
                            WriteLog(testInfoTxtFile, "DropTPScreen found, now find TblContent, scan and select button");
                            // find tblContent
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeTblContent == null && mTime.Minutes < 10)
                            {
                                aeTblContent = AUIUtilities.FindElementByID(ContentId, aeDropTPScreen);
                                mTime = DateTime.Now - mAppTime;
                                WriteLog(testInfoTxtFile, "Tbl Content  find time is :" + mTime.TotalMilliseconds / 1000);
                                Thread.Sleep(2000);
                            }

                            if (aeTblContent == null)
                            {
                                Log(sErrorMessage);
                                WriteLog(testInfoTxtFile, "Drop TblContent not found:");
                                MessageBox.Show(sErrorMessage, "Drop");
                                Thread.Sleep(3600000 * 20);
                            }

                            // Scan location  
                            // "m_TxtScannedValue"  This is Edit Control, should use setFocus + sendKeys
                            //if (AUIUtilities.FindDocumentAndSendText(DropScanId, aeDropTPScreen, "HU.1", ref sErrorMessage))
                            if (aeTblContent.Current.Name.StartsWith("Enter"))
                            {
                                mAppTime = DateTime.Now;
                                mTime = DateTime.Now - mAppTime;
                                while (aeDropField == null && mTime.Minutes < 3)
                                {
                                    aeDropField = AUIUtilities.FindElementByID(DropScanId, aeDropTPScreen);
                                    mTime = DateTime.Now - mAppTime;
                                    WriteLog(testInfoTxtFile, "drop field  find time is :" + mTime.TotalMilliseconds / 1000);
                                    Thread.Sleep(2000);
                                }
                                //aeDropField = AUIUtilities.FindElementByID(DropScanId, aeDropTPScreen);
                                 string destID = DestLoc;
                                if (aeDropField != null)
                                {
                                    if (SrcLoc.Equals("KCP5.2"))
                                        destID = "KCP5.1";
                                    else if (SrcLoc.Equals("KCP5.2R"))
                                        destID = "KCP5.1R";
                                    else if (DestLoc.Equals("FIL.DP"))
                                    {
                                        destID = "FIL.1";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("WIP1"))
                                    {
                                        destID = "WIP1.1";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("WIP2"))
                                    {
                                        destID = "WIP2.1";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("WIP3"))
                                    {
                                        destID = "WIP3.1";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("FC3"))
                                    {
                                        destID = "FC3.1";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("FC4"))
                                    {
                                        destID = "FC4.1";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("KCP1"))
                                    {
                                        destID = "KCP1.1";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("KCP2"))
                                    {
                                        destID = "KCP2.1";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("KCP3"))
                                    {
                                        destID = "KCP3.1";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("KCP4"))
                                    {
                                        destID = "KCP4.1";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("KCP5"))
                                    {
                                        destID = "KCP5.2";
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }

                                    Thread.Sleep(3000);
                                    aeDropField.SetFocus();
                                    Thread.Sleep(1000);
                                    System.Windows.Forms.SendKeys.SendWait(destID);
                                    Thread.Sleep(2000);
                                    //ValuePattern vp = (ValuePattern)aeDropField.GetCurrentPattern(ValuePattern.Pattern);
                                    //Thread.Sleep(2000);
                                    //vp.SetValue("HU.1");
                                    Console.WriteLine(destID+" scanned");
                                    Log(destID+"  OK");
                                }
                                else
                                {
                                    Log(sErrorMessage);
                                    WriteLog(testInfoTxtFile, "Drop ScanID not found:");
                                    MessageBox.Show(sErrorMessage, "Drop to " + destID);
                                }
                            }
                            Thread.Sleep(500);

                            // OK 
                            AutomationElement aeDropOKButton = AUIUtilities.FindElementByID(DropOKId, aeDropTPScreen);
                            if (aeDropOKButton == null)
                            {
                                sErrorMessage = "Drop OK button not found";
                                Console.WriteLine("Drop OK button not found:" + DropOKId);
                                Log("Drop OK button not found:" + DropOKId);
                            }
                            else
                            {
                                sDrop7TCount++;
                                Thread.Sleep(500);
                                Input.MoveTo(aeDropOKButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeDropOKButton));
                                slogMsg = "(E.OK)--> --> --> --> Total Drop" + sDrop7TCount ;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("scan reel to consume"))
                    {
                        #region  // Handle scan reel to consume screen
                        slogMsg = "(T)--> scan reel to consume --> ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        // Find the consume reel screen cancel button
                        Thread.Sleep(500); 
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        AutomationElement aeBtnCancel = null;
                        string CancelBtnId = "m_BtnCancel";  // cancel button
                        while (aeBtnCancel == null && mTime.Minutes < 5)
                        {
                            aeBtnCancel = AUIUtilities.FindElementByID(CancelBtnId, aeForm);
                            mTime = DateTime.Now - mAppTime;
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(1000);
                        if (aeBtnCancel == null)
                        {
                            sErrorMessage = "Screen cancel button not found";
                            MessageBox.Show(sErrorMessage, instructionName);
                        }
                        else
                        {
                            WriteLog(testInfoTxtFile, "Click cancel button");
                            Thread.Sleep(500);
                            Input.MoveTo(aeBtnCancel);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeBtnCancel));
                            slogMsg = "(Consume.OK)--> --> --> --> Next";
                            Log(slogMsg);
                            WriteLog(testInfoTxtFile, slogMsg);
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("transport could not be selected"))
                    {
                        #region  // Handle transport could not be selected screen
                        slogMsg = "(T)--> transport could not be selected, refresh screen --> ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        // Find the transport refresh screen button
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        AutomationElement aeBtnRefresh = null;
                        string RefreshBtnId = "m_BtnRefresh";  // refresh button
                        while (aeBtnRefresh == null && mTime.Minutes < 5)
                        {
                            aeBtnRefresh = AUIUtilities.FindElementByID(RefreshBtnId, aeForm);
                            mTime = DateTime.Now - mAppTime;
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(1000);
                        if (aeBtnRefresh == null)
                        {
                            sErrorMessage = "Screen refresh button not found";
                            MessageBox.Show(sErrorMessage, "Transport could not be selected");
                        }
                        else
                        {
                            WriteLog(testInfoTxtFile, "transport could not be selected, refresh screen");
                            Thread.Sleep(500);
                            Input.MoveTo(aeBtnRefresh);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeBtnRefresh));
                            slogMsg = "(T.OK)--> --> --> --> Next";
                            Log(slogMsg);
                            WriteLog(testInfoTxtFile, slogMsg);
                        }
                        #endregion
                    }
                    else if (instructionName.Equals(string.Empty))
                    {
                        slogMsg = " *** (Empty)--> Handle Empty instruction name : " + instructionName;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        //MessageBox.Show(slogMsg, instructionName);
                    }
                    else
                    {
                        slogMsg = " *** (U)--> Handle unexpected instruction screens : " + instructionName;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        MessageBox.Show(slogMsg, instructionName);
                    }
                }
                else
                {
                    sErrorMessage = "Instruction lable not found:";
                    Log(sErrorMessage);
                    WriteLog(testInfoTxtFile, "Instruction lable not found:");
                    MessageBox.Show(sErrorMessage, "Get instruction screen");
                }
                Thread.Sleep(2000);
                #endregion
                Thread.Sleep(1000);
            }
        }

        public void SmallProngliftScannerStart(string source, string destination)
        {
            string slogMsg = "Pronglift2TScannerStart : " + System.DateTime.Now;
            Log(slogMsg);
            WriteLog(testInfoTxtFile, slogMsg);

            testInfoTxtFile = Path.Combine(@"C:\KC\PutAway", "ProngLift2T.log");
            StreamWriter writeInfo = File.CreateText(testInfoTxtFile);
            string info = "Start test : " + DateTime.Now;
            writeInfo.WriteLine(info);
            writeInfo.Close();
            sDrop2TCount = 0;

            DateTime mAppTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mAppTime;

            Random RandomClass = new Random();
            int numOption = 0;
            string[] KCPLine = { "AV1", "AV2", "AR1", "AR2", "AV3", "AR3" };

            while (true)
            {
                #region // A find mainform
                mAppTime = DateTime.Now;
                mTime = DateTime.Now - mAppTime;
                slogMsg = "(A)<-- Find Application aeForm : " + System.DateTime.Now;
                Log(slogMsg);
                WriteLog(testInfoTxtFile, slogMsg);
                
                while (aeForm == null && mTime.Minutes < 10)
                {
                    aeForm = AUIUtilities.FindElementByID("MainForm", AutomationElement.RootElement);
                    WriteLog(testInfoTxtFile, "MainForm not found " + mTime.Seconds);
                    mTime = DateTime.Now - mAppTime;
                    Thread.Sleep(5000);
                }

                // if after 5 minutes still no mainform,throw exception 
                if (aeForm == null)
                {
                    AutomationElement aeError = AUIUtilities.FindElementByID("ErrorScreen", AutomationElement.RootElement);
                    if (aeError != null)
                        AUICommon.ErrorWindowHandling(aeError, ref sErrorMessage);
                    else
                        sErrorMessage = "Application Startup failed,see logging";

                    throw new Exception(sErrorMessage);
                }
                else
                {
                    Console.WriteLine("Application maeForm name : " + aeForm.Current.Name);
                    Log("Application maeForm name : " + aeForm.Current.Name + " - Time: " + System.DateTime.Now);
                }

                slogMsg = "A.0 --> MainForm founded : " + System.DateTime.Now + " -----------------";
                Log(slogMsg);
                WriteLog(testInfoTxtFile, slogMsg);
                #endregion

                #region  // B Handle transport screen display
                string instructionName = string.Empty;
                AutomationElement aeBtnHome = null;

                mAppTime = DateTime.Now;
                mTime = DateTime.Now - mAppTime;
                aeInstructionLable = null;

                while (aeInstructionLable == null && mTime.Minutes < 5)
                {
                    aeInstructionLable = AUIUtilities.FindElementByID(InstructionId, aeForm);
                    mTime = DateTime.Now - mAppTime;
                    Thread.Sleep(1000);
                }
                if (aeInstructionLable != null)
                {
                    instructionName = aeInstructionLable.Current.Name;
                    slogMsg = " *** (B)--> Handle instruction screens : " + instructionName;
                    Log(slogMsg);
                    WriteLog(testInfoTxtFile, slogMsg);

                    if (instructionName.ToLower().StartsWith("waiting for transports"))
                    {
                        #region // handle waiting for transport instruction
                        aeBtnHome = AUIUtilities.FindElementByID("m_BtnHome", aeForm);
                        if (aeBtnHome != null)
                        {
                            Thread.Sleep(2000);
                            // sometime screen change to other select screen after 'waiting for...'
                            if (aeBtnHome.Current.IsEnabled)
                                Input.MoveToAndClick(aeBtnHome);
                            else
                            {
                                slogMsg = " ***(B.X )--> screens changed during, retry: " + instructionName;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        else
                        {
                            sErrorMessage = "Buttom Home not found:";
                            slogMsg = " ***(B.2 )--> Buttom Home not found:, retry maybe screen already changed : ";
                            Log(slogMsg);
                            WriteLog(testInfoTxtFile, slogMsg);
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("what do you want to do"))
                    {
                        #region // handle what do you want to do screen
                        // click first option  
                        slogMsg = " ***** (B.1)--> Click first option now : ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        aeOption = AUIUtilities.FindElementByID(optionId, aeForm);
                        if (aeOption != null)
                        {
                            Thread.Sleep(1000);

                            // Set a property condition that will be used to find the control.
                            System.Windows.Automation.Condition c = new PropertyCondition(
                                AutomationElement.ControlTypeProperty, ControlType.Button);

                            AutomationElementCollection aeOptionButton = aeOption.FindAll(TreeScope.Element | TreeScope.Descendants, c);
                            Thread.Sleep(1000);

                            OptionPt = AUIUtilities.GetElementCenterPoint(aeOptionButton[0]);
                            Thread.Sleep(1000);
                            Input.MoveTo(OptionPt);

                            WriteLog(testInfoTxtFile, "numOption  : " + 0);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(OptionPt);
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("select a zone")
                        || instructionName.ToLower().StartsWith("all transports in zone are finished"))
                    { 
                        #region // handle select a zone screen
                        // click zone option  
                        slogMsg = " ***** (B.1)--> Click zone option now : ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        aeOption = AUIUtilities.FindElementByID(optionId, aeForm);
                        if (aeOption != null)
                        {
                            Thread.Sleep(1000);

                            // Set a property condition that will be used to find the control.
                            System.Windows.Automation.Condition c = new PropertyCondition(
                                AutomationElement.ControlTypeProperty, ControlType.Button);

                            AutomationElementCollection aeOptionButton = aeOption.FindAll(TreeScope.Element | TreeScope.Descendants, c);
                            Thread.Sleep(1000);
                            
                            numOption = RandomClass.Next(0, aeOptionButton.Count);
                            OptionPt = AUIUtilities.GetElementCenterPoint(aeOptionButton[numOption]);
                            Thread.Sleep(1000);
                            Input.MoveTo(OptionPt);

                            WriteLog(testInfoTxtFile, "numOption  : " + numOption);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(OptionPt);
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("select a transport"))
                    {
                        #region // handle select a transport screen
                        slogMsg = "(S)--> Selection instruction screen found : ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        #region  // Assign Transport
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        aeSelectTPScreen = null;
                        while (aeSelectTPScreen == null && mTime.Minutes < 10)
                        {
                            Console.WriteLine("Find SelectTransportScreen aeSelectTPScreen : " + System.DateTime.Now);
                            aeSelectTPScreen = AUIUtilities.FindElementByID(SelectTPId, aeForm);
                            Console.WriteLine("SelectTransportScreen aeSelectTPScreen: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            WriteLog(testInfoTxtFile, "Select screen find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(500);
                        }
                        if (aeSelectTPScreen == null)
                        {
                            sErrorMessage = "SelectTransportScreen not found";
                            Console.WriteLine("FindElementByID failed:" + SelectTPId);
                            MessageBox.Show("No new transport exist any more", "Find select a transport screen");
                            Thread.Sleep(3600000 * 20);
                        }
                        else
                        {
                            Console.WriteLine("SelectTPScreen found, now find select button");
                            Thread.Sleep(500);

                            // find reel unit id
                            /*AutomationElement aeUnit = AUIUtilities.FindElementByID(ReelUnitId, aeSelectTPScreen);
                            if (aeUnit == null)
                            {
                                sErrorMessage = "UnitID not found";
                                Console.WriteLine("UnitId not found:" + ReelUnitId);
                                MessageBox.Show(sErrorMessage, "Assign Transport");
                                Thread.Sleep(3600000 * 20);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                TrUnitId = aeUnit.Current.Name;
                                Log("Select transport with unit: " + TrUnitId);
                                WriteLog(testInfoTxtFile, "Select transport with unit: " + TrUnitId);
                            }
                            */

                            // find source loc
                            aeSrc = null;
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeSrc == null && mTime.Minutes < 2)
                            {
                                aeSrc = AUIUtilities.FindElementByID("m_LblSourceLocation", aeSelectTPScreen);
                                mTime = DateTime.Now - mAppTime;
                                WriteLog(testInfoTxtFile, "sourceLoc find time is :" + mTime.TotalMilliseconds / 1000 + "  m_LblSourceLocation");
                                Thread.Sleep(500);
                            }

                            if (aeSrc == null)
                            {
                                sErrorMessage = "Source not found";
                                Console.WriteLine("Source not found:" + "m_LblSourceLocation");
                                MessageBox.Show(sErrorMessage, instructionName);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                SrcLoc = aeSrc.Current.Name;
                                slogMsg = "(S.1)--> transport source is: " + SrcLoc;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }

                            // find destLocation id
                            aeDest = null;
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeDest == null && mTime.Minutes < 2)
                            {
                                aeDest = AUIUtilities.FindElementByID("m_LblDestinationLocation", aeSelectTPScreen);
                                mTime = DateTime.Now - mAppTime;
                                WriteLog(testInfoTxtFile, "destination find time is :" + mTime.TotalMilliseconds / 1000 + "  m_LblDestinationLocation");
                                Thread.Sleep(500);
                            }

                            if (aeDest == null)
                            {
                                sErrorMessage = "dest ID not found";
                                Console.WriteLine("dest Id not found:" + "m_LblDestinationLocation");
                                MessageBox.Show(sErrorMessage, instructionName);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                DestLoc = aeDest.Current.Name;
                                slogMsg = "(S.2)--> Selected transport Destination is: " + DestLoc;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }

                            Thread.Sleep(1000);
                            string SelectId = "m_BtnOK";
                            AutomationElement aeSelectButton = AUIUtilities.FindElementByID(SelectId, aeSelectTPScreen);
                            if (aeSelectButton == null)
                            {
                                slogMsg = "(S.X)--> Selection confirm button not found: " + SelectId;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                            else
                            {
                                Thread.Sleep(500);
                                Input.MoveTo(aeSelectButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSelectButton));
                                slogMsg = "(S.OK)--> Transport is assigned: OK";
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("pick"))
                    {
                        #region // handle pick and scan the pallet or reel screen
                        slogMsg = "(P)--> Handle pick screen : " + instructionName;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        // Pick transport
                        string PickTPSreenId = "m_TblContent";
                        string PickScanId = "m_TxtScannedValue";  // edit control
                        string PickOKId = "m_BtnOK";

                        // Find the pick transport screen  SelectTransportScreen
                        Thread.Sleep(1000);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        AutomationElement aePickTPScreen = null;
                        while (aePickTPScreen == null && mTime.Minutes < 10)
                        {
                            Console.WriteLine("Find PickTransportScreen aeSelectTPScreen : " + System.DateTime.Now);
                            aePickTPScreen = AUIUtilities.FindElementByID(PickTPSreenId, aeForm);
                            Console.WriteLine("PickTransportScreen aePickTPScreen: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(500);
                        }
                        if (aePickTPScreen == null)
                        {
                            sErrorMessage = "PickTransportScreen not found";
                            Console.WriteLine("Pick FindElementByID failed:" + PickTPSreenId);
                            MessageBox.Show("No Pick transport screen exist", "Find pick a transport screen");
                            Thread.Sleep(3600000 * 20);
                        }
                        else
                        {
                            Console.WriteLine("PickTPScreen found, now find reelLd, scan and select button");
                            Thread.Sleep(1000);
                            // Scan location  
                            // "m_TxtScannedValue"  This is Edit Control, should use setFocus + sendKeys
                            if (FindDocumentAndSendText(PickScanId, aePickTPScreen, SrcLoc, ref sErrorMessage))
                            {
                                slogMsg = "(P.1)--> " + SrcLoc + "  scan OK : ";
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                Log("found AND scan field failed::" + sErrorMessage);
                                MessageBox.Show(sErrorMessage, " Pick Transport");
                            }

                            Thread.Sleep(2000);
                            // OK 
                            AutomationElement aePickOKButton = AUIUtilities.FindElementByID(PickOKId, aePickTPScreen);
                            if (aePickOKButton == null)
                            {
                                sErrorMessage = "Select button not found";
                                MessageBox.Show(sErrorMessage, " Pick Transport");
                                Thread.Sleep(1000000000);
                            }
                            else
                            {
                                Thread.Sleep(500);
                                Input.MoveTo(aePickOKButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aePickOKButton));

                                slogMsg = "(P.2)--> " + " Pick transport confirmed : ";
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("drop"))
                    {
                        #region  // Handle drop the reel at the destination location
                        slogMsg = "(D)--> Drop location is --> " + DestLoc;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        // Drop transport
                        string ContentId = "m_TblContent";
                        string DropOKId = "m_BtnOK";

                        #region  // check Drop loc empty  region
                        if (CheckDropLocEmpty && DestLoc.ToLower().StartsWith("hu"))
                        {
                            // Wait until FIL input empty
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            bool FilInputEmpty = false;
                            while (FilInputEmpty == false)
                            {

                                if (IsDropLocationEmpty("HU.1"))
                                {
                                    FilInputEmpty = true;
                                    slogMsg = "(D.1)-->  : " + DestLoc + " ::HU.1--> empty, " + System.DateTime.Now;

                                }
                                else
                                {
                                    FilInputEmpty = false;
                                    mTime = DateTime.Now - mAppTime;
                                    slogMsg = "(D.1)-->  : " + DestLoc + " ::HU.1--> not empty," + System.DateTime.Now;
                                }
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                                Thread.Sleep(5000);
                            }
                        }
                        #endregion

                        // Find the  transport screen  DropTransportScreen
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;

                        AutomationElement aeDropTPScreen = null;
                        string DropScanId = "m_TxtScannedValue";  // edit control
                        AutomationElement aeDropField = null;
                        AutomationElement aeTblContent = null;
                        while (aeDropTPScreen == null && mTime.Minutes < 5)
                        {
                            Console.WriteLine("Find DropTransportScreen : " + System.DateTime.Now);
                            aeDropTPScreen = AUIUtilities.FindElementByID(DropTPSreenId, aeForm);
                            Console.WriteLine("DropTransportScreen found: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(3000);
                        if (aeDropTPScreen == null)
                        {
                            sErrorMessage = "DropTransportScreen not found";
                            Console.WriteLine("Drop FindElementByID failed:" + DropTPSreenId);
                        }
                        else
                        {
                            WriteLog(testInfoTxtFile, "DropTPScreen found, now find TblContent, scan and select button");
                            // find tblContent
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeTblContent == null && mTime.Minutes < 10)
                            {
                                aeTblContent = AUIUtilities.FindElementByID(ContentId, aeDropTPScreen);
                                mTime = DateTime.Now - mAppTime;
                                WriteLog(testInfoTxtFile, "Tbl Content  find time is :" + mTime.TotalMilliseconds / 1000);
                                Thread.Sleep(2000);
                            }

                            if (aeTblContent == null)
                            {
                                Log(sErrorMessage);
                                WriteLog(testInfoTxtFile, "Drop TblContent not found:");
                                MessageBox.Show(sErrorMessage, "Drop");
                                Thread.Sleep(3600000 * 20);
                            }

                            // Scan location  
                            // "m_TxtScannedValue"  This is Edit Control, should use setFocus + sendKeys
                            //if (AUIUtilities.FindDocumentAndSendText(DropScanId, aeDropTPScreen, "HU.1", ref sErrorMessage))
                            if (aeTblContent.Current.Name.StartsWith("Enter"))
                            {
                                mAppTime = DateTime.Now;
                                mTime = DateTime.Now - mAppTime;
                                while (aeDropField == null && mTime.Minutes < 3)
                                {
                                    aeDropField = AUIUtilities.FindElementByID(DropScanId, aeDropTPScreen);
                                    mTime = DateTime.Now - mAppTime;
                                    WriteLog(testInfoTxtFile, "drop field  find time is :" + mTime.TotalMilliseconds / 1000);
                                    Thread.Sleep(2000);
                                }
                                //aeDropField = AUIUtilities.FindElementByID(DropScanId, aeDropTPScreen);
                                string destID = DestLoc;
                                if (aeDropField != null)
                                {
                                    if (SrcLoc.Equals("KCP5.2"))
                                        destID = "KCP5.1";
                                    else if (SrcLoc.Equals("KCP5.2R"))
                                        destID = "KCP5.1R";

                                    Thread.Sleep(3000);
                                    aeDropField.SetFocus();
                                    Thread.Sleep(1000);
                                    System.Windows.Forms.SendKeys.SendWait(destID);
                                    Thread.Sleep(2000);
                                    //ValuePattern vp = (ValuePattern)aeDropField.GetCurrentPattern(ValuePattern.Pattern);
                                    //Thread.Sleep(2000);
                                    //vp.SetValue("HU.1");
                                    Console.WriteLine(destID + " scanned");
                                    Log(destID + "  OK");
                                }
                                else
                                {
                                    Log(sErrorMessage);
                                    WriteLog(testInfoTxtFile, "Drop ScanID not found:");
                                    MessageBox.Show(sErrorMessage, "Drop to " + destID);
                                }
                            }
                            Thread.Sleep(500);

                            // OK 
                            AutomationElement aeDropOKButton = AUIUtilities.FindElementByID(DropOKId, aeDropTPScreen);
                            if (aeDropOKButton == null)
                            {
                                sErrorMessage = "Drop OK button not found";
                                Console.WriteLine("Drop OK button not found:" + DropOKId);
                                Log("Drop OK button not found:" + DropOKId);
                            }
                            else
                            {
                                sDrop2TCount++;
                                Thread.Sleep(500);
                                Input.MoveTo(aeDropOKButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeDropOKButton));
                                slogMsg = "(D.OK)--> --> --> --> total drop :" + sDrop2TCount; ;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("enter"))
                    {
                        #region  // Handle enter the location where the reel has been dropped
                        slogMsg = "(E)--> Enter location is --> " + DestLoc;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        // Drop transport
                        string ContentId = "m_TblContent";
                        string DropOKId = "m_BtnOK";

                        // Find the  transport screen  DropTransportScreen
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;

                        AutomationElement aeDropTPScreen = null;
                        string DropScanId = "m_TxtScannedValue";  // edit control
                        AutomationElement aeDropField = null;
                        AutomationElement aeTblContent = null;
                        while (aeDropTPScreen == null && mTime.Minutes < 5)
                        {
                            Console.WriteLine("Find DropTransportScreen : " + System.DateTime.Now);
                            aeDropTPScreen = AUIUtilities.FindElementByID(DropTPSreenId, aeForm);
                            Console.WriteLine("DropTransportScreen found: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(3000);
                        if (aeDropTPScreen == null)
                        {
                            sErrorMessage = "DropTransportScreen not found";
                            Console.WriteLine("Drop FindElementByID failed:" + DropTPSreenId);
                        }
                        else
                        {
                            WriteLog(testInfoTxtFile, "DropTPScreen found, now find TblContent, scan and select button");
                            // find tblContent
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeTblContent == null && mTime.Minutes < 10)
                            {
                                aeTblContent = AUIUtilities.FindElementByID(ContentId, aeDropTPScreen);
                                mTime = DateTime.Now - mAppTime;
                                WriteLog(testInfoTxtFile, "Tbl Content  find time is :" + mTime.TotalMilliseconds / 1000);
                                Thread.Sleep(2000);
                            }

                            if (aeTblContent == null)
                            {
                                Log(sErrorMessage);
                                WriteLog(testInfoTxtFile, "Drop TblContent not found:");
                                MessageBox.Show(sErrorMessage, "Drop");
                                Thread.Sleep(3600000 * 20);
                            }

                            // Scan location  
                            // "m_TxtScannedValue"  This is Edit Control, should use setFocus + sendKeys
                            //if (AUIUtilities.FindDocumentAndSendText(DropScanId, aeDropTPScreen, "HU.1", ref sErrorMessage))
                            if (aeTblContent.Current.Name.StartsWith("Enter"))
                            {
                                mAppTime = DateTime.Now;
                                mTime = DateTime.Now - mAppTime;
                                while (aeDropField == null && mTime.Minutes < 3)
                                {
                                    aeDropField = AUIUtilities.FindElementByID(DropScanId, aeDropTPScreen);
                                    mTime = DateTime.Now - mAppTime;
                                    WriteLog(testInfoTxtFile, "drop field  find time is :" + mTime.TotalMilliseconds / 1000);
                                    Thread.Sleep(2000);
                                }
                                //aeDropField = AUIUtilities.FindElementByID(DropScanId, aeDropTPScreen);
                                string destID = DestLoc;
                                if (aeDropField != null)
                                {
                                   if (DestLoc.Equals("KCP1"))
                                    {
                                        destID = "KCP1."+KCPLine[numOption = RandomClass.Next(0, 6)];
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("KCP2"))
                                    {
                                        destID = "KCP2." + KCPLine[numOption = RandomClass.Next(0, 4)];
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("KCP3"))
                                    {
                                        destID = "KCP3." + KCPLine[numOption = RandomClass.Next(0, 6)];
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                    else if (DestLoc.Equals("KCP4"))
                                    {
                                        destID = "KCP4." + KCPLine[numOption = RandomClass.Next(0, 4)];
                                        while (CheckDropLocEmpty && !IsDropLocationEmpty(destID))
                                        {
                                            mTime = DateTime.Now - mAppTime;
                                            slogMsg = "(E.1)-->  : " + destID + " ::--> not empty," + System.DateTime.Now;
                                            Log(slogMsg);
                                            WriteLog(testInfoTxtFile, slogMsg);
                                            Thread.Sleep(5000);
                                        }
                                    }
                                   
                                    Thread.Sleep(3000);
                                    aeDropField.SetFocus();
                                    Thread.Sleep(1000);
                                    System.Windows.Forms.SendKeys.SendWait(destID);
                                    Thread.Sleep(2000);
                                    Console.WriteLine(destID + " scanned");
                                    Log(destID + "  OK");
                                }
                                else
                                {
                                    Log(sErrorMessage);
                                    WriteLog(testInfoTxtFile, "Drop ScanID not found:");
                                    MessageBox.Show(sErrorMessage, "Drop to " + destID);
                                }
                            }
                            Thread.Sleep(500);

                            // OK 
                            AutomationElement aeDropOKButton = AUIUtilities.FindElementByID(DropOKId, aeDropTPScreen);
                            if (aeDropOKButton == null)
                            {
                                sErrorMessage = "Drop OK button not found";
                                Console.WriteLine("Drop OK button not found:" + DropOKId);
                                Log("Drop OK button not found:" + DropOKId);
                            }
                            else
                            {
                                sDrop2TCount++;
                                Thread.Sleep(500);
                                Input.MoveTo(aeDropOKButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeDropOKButton));
                                slogMsg = "(E.OK)--> --> --> --> Total drop :" + sDrop2TCount;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("scan reel to consume"))
                    {
                        #region  // Handle scan reel to consume screen
                        slogMsg = "(T)--> scan reel to consume --> ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        // Find the consume reel screen cancel button
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        AutomationElement aeBtnCancel = null;
                        string CancelBtnId = "m_BtnCancel";  // cancel button
                        while (aeBtnCancel == null && mTime.Minutes < 5)
                        {
                            aeBtnCancel = AUIUtilities.FindElementByID(CancelBtnId, aeForm);
                            mTime = DateTime.Now - mAppTime;
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(1000);
                        if (aeBtnCancel == null)
                        {
                            sErrorMessage = "Screen cancel button not found";
                            MessageBox.Show(sErrorMessage, instructionName);
                        }
                        else
                        {
                            WriteLog(testInfoTxtFile, "Click cancel button");
                            Thread.Sleep(500);
                            Input.MoveTo(aeBtnCancel);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeBtnCancel));
                            slogMsg = "(Consume.OK)--> --> --> --> Next";
                            Log(slogMsg);
                            WriteLog(testInfoTxtFile, slogMsg);
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("transport could not be selected"))
                    {
                        #region  // Handle transport could not be selected screen
                        slogMsg = "(T)--> transport could not be selected, refresh screen --> ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        // Find the transport refresh screen button
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        AutomationElement aeBtnRefresh = null;
                        string RefreshBtnId = "m_BtnRefresh";  // refresh button
                        while (aeBtnRefresh == null && mTime.Minutes < 5)
                        {
                            aeBtnRefresh = AUIUtilities.FindElementByID(RefreshBtnId, aeForm);
                            mTime = DateTime.Now - mAppTime;
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(1000);
                        if (aeBtnRefresh == null)
                        {
                            sErrorMessage = "Screen refresh button not found";
                            MessageBox.Show(sErrorMessage, "Transport could not be selected");
                        }
                        else
                        {
                            WriteLog(testInfoTxtFile, "transport could not be selected, refresh screen");
                            Thread.Sleep(500);
                            Input.MoveTo(aeBtnRefresh);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeBtnRefresh));
                            slogMsg = "(T.OK)--> --> --> --> Next";
                            Log(slogMsg);
                            WriteLog(testInfoTxtFile, slogMsg);
                        }
                        #endregion
                    }
                    else if (instructionName.Equals(string.Empty))
                    {
                        slogMsg = " *** (Empty)--> Handle Empty instruction name : " + instructionName;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        //MessageBox.Show(slogMsg, instructionName);
                    }
                    else
                    {
                        slogMsg = " *** (U)--> Handle unexpected instruction screens : " + instructionName;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        MessageBox.Show(slogMsg, instructionName);
                    }
                }
                else
                {
                    sErrorMessage = "Instruction lable not found:";
                    Log(sErrorMessage);
                    WriteLog(testInfoTxtFile, "Instruction lable not found:");
                    MessageBox.Show(sErrorMessage, "Get instruction screen");
                }
                Thread.Sleep(2000);
                #endregion
                Thread.Sleep(1000);
            }
            
        }

        //=====================================================================
        /// <summary>
        /// Method will start new tests
        /// </summary>
        public void ForkliftScannerStart(string source, string destination)
        {
            string slogMsg = "ForkliftScannerStart : " + System.DateTime.Now;
            Log(slogMsg);
            WriteLog(testInfoTxtFile, slogMsg);

            testInfoTxtFile = Path.Combine(@"C:\KC\PutAway", "ForkLift.log");
            StreamWriter writeInfo = File.CreateText(testInfoTxtFile);
            string info = "Start test : " + DateTime.Now;
            writeInfo.WriteLine(info);
            writeInfo.Close();
            sDropFTCount = 0;

            DateTime mAppTime = DateTime.Now;
            TimeSpan mTime = DateTime.Now - mAppTime;
            
            while (true)
            {
                #region // A find mainform
                mAppTime = DateTime.Now;
                mTime = DateTime.Now - mAppTime;
                slogMsg = "(A)<-- Find Application aeForm : " + System.DateTime.Now;
                Log(slogMsg);
                WriteLog(testInfoTxtFile, slogMsg);
                
                
                while (aeForm == null && mTime.Minutes < 10)
                {
                    aeForm = AUIUtilities.FindElementByID("MainForm", AutomationElement.RootElement);
                    mTime = DateTime.Now - mAppTime;
                    Thread.Sleep(5000);
                }

                // if after 5 minutes still no mainform,throw exception 
                if (aeForm == null)
                {
                    AutomationElement aeError = AUIUtilities.FindElementByID("ErrorScreen", AutomationElement.RootElement);
                    if (aeError != null)
                        AUICommon.ErrorWindowHandling(aeError, ref sErrorMessage);
                    else
                        sErrorMessage = "Application Startup failed,see logging";

                    throw new Exception(sErrorMessage);
                }
                else
                {
                    Console.WriteLine("Application maeForm name : " + aeForm.Current.Name);
                    Log("Application maeForm name : " + aeForm.Current.Name + " - Time: " + System.DateTime.Now);
                }

                slogMsg = "A.0 --> MainForm founded : " + System.DateTime.Now+" -----------------";
                Log(slogMsg);
                WriteLog(testInfoTxtFile, slogMsg);
                #endregion

                #region  // B Handle transport screen display
                string instructionName = string.Empty;
                AutomationElement aeBtnHome = null;

                mAppTime = DateTime.Now;
                mTime = DateTime.Now - mAppTime;
                aeInstructionLable = null;
                
                while (aeInstructionLable == null && mTime.Minutes < 5)
                {
                    aeInstructionLable = AUIUtilities.FindElementByID(InstructionId, aeForm);
                    mTime = DateTime.Now - mAppTime;
                    Thread.Sleep(1000);
                }
                if (aeInstructionLable != null)
                {
                    instructionName = aeInstructionLable.Current.Name;
                    slogMsg = " *** (B)--> Handle instruction screens : " + instructionName;
                    Log(slogMsg);
                    WriteLog(testInfoTxtFile, slogMsg);

                    if (instructionName.ToLower().StartsWith("waiting for transports"))
                    {
                        #region // handle waiting for transport instruction
                        aeBtnHome = AUIUtilities.FindElementByID("m_BtnHome", aeForm);
                        if (aeBtnHome != null)
                        {
                            Thread.Sleep(2000);
                            // sometime screen change to other select screen after 'waiting for...'
                            if (aeBtnHome.Current.IsEnabled)
                                Input.MoveToAndClick(aeBtnHome);
                            else 
                            {
                                slogMsg = " ***(B.X )--> screens changed during, retry: " + instructionName;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        else
                        {
                            sErrorMessage = "Buttom Home not found:";
                            slogMsg = " ***(B.2 )--> Buttom Home not found:, retry maybe screen already changed : ";
                            Log(slogMsg);
                            WriteLog(testInfoTxtFile, slogMsg);
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("what do you want to do"))
                    {
                        #region // handle what do you want to do screen
                        // click first option  
                        slogMsg = " ***** (B.1)--> Click first option now : ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        aeOption  = AUIUtilities.FindElementByID(optionId, aeForm);
                        if (aeOption != null)
                        {
                            Thread.Sleep(1000);
                            
                            // Set a property condition that will be used to find the control.
                            System.Windows.Automation.Condition c = new PropertyCondition(
                                AutomationElement.ControlTypeProperty, ControlType.Button);

                            AutomationElementCollection aeOptionButton = aeOption.FindAll(TreeScope.Element | TreeScope.Descendants, c);
                            Thread.Sleep(1000);
                           
                            OptionPt = AUIUtilities.GetElementCenterPoint(aeOptionButton[0]);
                            Thread.Sleep(1000);
                            Input.MoveTo(OptionPt);
                 
                            WriteLog(testInfoTxtFile, "numOption  : " + 0);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(OptionPt);
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("select a transport"))
                    {
                        #region // handle select a transport screen
                        slogMsg = "(S)--> Selection instruction screen found : ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        #region  // Assign Transport
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        aeSelectTPScreen = null;
                        while (aeSelectTPScreen == null && mTime.Minutes < 10)
                        {
                            Console.WriteLine("Find SelectTransportScreen aeSelectTPScreen : " + System.DateTime.Now);
                            aeSelectTPScreen = AUIUtilities.FindElementByID(SelectTPId, aeForm);
                            Console.WriteLine("SelectTransportScreen aeSelectTPScreen: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            WriteLog(testInfoTxtFile, "Select screen find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(500);
                        }
                        if (aeSelectTPScreen == null)
                        {
                            sErrorMessage = "SelectTransportScreen not found";
                            Console.WriteLine("FindElementByID failed:" + SelectTPId);
                            MessageBox.Show("No new transport exist any more", "Find select a transport screen");
                            Thread.Sleep(3600000 * 20);
                        }
                        else
                        {
                            Console.WriteLine("SelectTPScreen found, now find select button");
                            Thread.Sleep(500);

                            // find reel unit id
                            /*AutomationElement aeUnit = AUIUtilities.FindElementByID(ReelUnitId, aeSelectTPScreen);
                            if (aeUnit == null)
                            {
                                sErrorMessage = "UnitID not found";
                                Console.WriteLine("UnitId not found:" + ReelUnitId);
                                MessageBox.Show(sErrorMessage, "Assign Transport");
                                Thread.Sleep(3600000 * 20);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                TrUnitId = aeUnit.Current.Name;
                                Log("Select transport with unit: " + TrUnitId);
                                WriteLog(testInfoTxtFile, "Select transport with unit: " + TrUnitId);
                            }
                            */

                            // find carrier id
                            aeCarrier = null;
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeCarrier == null && mTime.Minutes < 2)
                            {
                               aeCarrier = AUIUtilities.FindElementByID("m_LblCarrierId", aeSelectTPScreen);
                               mTime = DateTime.Now - mAppTime;
                               WriteLog(testInfoTxtFile, "carrierId find time is :" + mTime.TotalMilliseconds / 1000 + "  m_LblCarrierId");
                               Thread.Sleep(500);
                            }

                            if (aeCarrier == null)
                            {
                                sErrorMessage = "CarrierID not found";
                                Console.WriteLine("CarrierId not found:" + "m_LblCarrierId");
                                MessageBox.Show(sErrorMessage, "select a transport");
                                Thread.Sleep(3600000 * 20);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                CarrierId = aeCarrier.Current.Name;
                                slogMsg = "(S.1)--> Selected Carrier is: " + CarrierId;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }

                            // find destLocation id
                            AutomationElement aeDest = AUIUtilities.FindElementByID("m_LblDestinationLocation", aeSelectTPScreen);
                            if (aeDest == null)
                            {
                                sErrorMessage = "dest ID not found";
                                Console.WriteLine("dest Id not found:" + "m_LblDestinationLocation");
                                MessageBox.Show(sErrorMessage, "Assign Transport");
                                Thread.Sleep(3600000 * 20);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                DestLoc = aeDest.Current.Name;
                                slogMsg = "(S.2)--> Selected Destination is: " + DestLoc;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }

                            Thread.Sleep(1000);
                            string SelectId = "m_BtnOK";
                            AutomationElement aeSelectButton = AUIUtilities.FindElementByID(SelectId, aeSelectTPScreen);
                            if (aeSelectButton == null)
                            {
                                slogMsg = "(S.X)--> Selection confirm button not found: " + SelectId;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                            else
                            {
                                Thread.Sleep(500);
                                Input.MoveTo(aeSelectButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeSelectButton));
                                slogMsg = "(S.OK)--> Transport is assigned: OK";
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("pick"))
                    {
                        #region // handle pick and scan the pallet or reel screen
                        slogMsg = "(P)--> Handle pick screen : " + instructionName;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        // Pick transport
                        string PickTPSreenId = "m_TblContent";
                        string PickScanId = "m_TxtScannedValue";  // edit control
                        string PickOKId = "m_BtnOK";

                        // Find the pick transport screen  SelectTransportScreen
                        Thread.Sleep(1000);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        AutomationElement aePickTPScreen = null;
                        while (aePickTPScreen == null && mTime.Minutes < 10)
                        {
                            Console.WriteLine("Find PickTransportScreen aeSelectTPScreen : " + System.DateTime.Now);
                            aePickTPScreen = AUIUtilities.FindElementByID(PickTPSreenId, aeForm);
                            Console.WriteLine("PickTransportScreen aePickTPScreen: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(500);
                        }
                        if (aePickTPScreen == null)
                        {
                            sErrorMessage = "PickTransportScreen not found";
                            Console.WriteLine("Pick FindElementByID failed:" + PickTPSreenId);
                            MessageBox.Show("No Pick transport screen exist", "Find pick a transport screen");
                            Thread.Sleep(3600000 * 20);
                        }
                        else
                        {
                            Console.WriteLine("PickTPScreen found, now find reelLd, scan and select button");
                            Thread.Sleep(1000);
                            // Scan location  
                            // "m_TxtScannedValue"  This is Edit Control, should use setFocus + sendKeys
                            if (FindDocumentAndSendText(PickScanId, aePickTPScreen, CarrierId, ref sErrorMessage))
                            {
                                slogMsg = "(P.1)--> "+ CarrierId+"  scan OK : ";
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                Log("found AND scan field failed::" + sErrorMessage);
                                MessageBox.Show(sErrorMessage, " Pick Transport");
                            }

                            Thread.Sleep(2000);
                            // OK 
                            AutomationElement aePickOKButton = AUIUtilities.FindElementByID(PickOKId, aePickTPScreen);
                            if (aePickOKButton == null)
                            {
                                sErrorMessage = "Select button not found";
                                MessageBox.Show(sErrorMessage, " Pick Transport");
                                Thread.Sleep(1000000000);
                            }
                            else
                            {
                                Thread.Sleep(500);
                                Input.MoveTo(aePickOKButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aePickOKButton));

                                slogMsg = "(P.2)--> " + " Pick transport confirmed : ";
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("drop"))
                    {
                        #region  // Handle drop the reel at the destination location
                        slogMsg = "(D)--> Drop location is --> "+ DestLoc;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        // Drop transport
                        string ContentId = "m_TblContent";
                        string DropOKId = "m_BtnOK";

                        // Find the  transport screen  DropTransportScreen
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;

                        AutomationElement aeDropTPScreen = null;
                        string DropScanId = "m_TxtScannedValue";  // edit control
                        AutomationElement aeDropField = null;
                        AutomationElement aeTblContent = null;
                        while (aeDropTPScreen == null && mTime.Minutes < 5)
                        {
                            Console.WriteLine("Find DropTransportScreen : " + System.DateTime.Now);
                            aeDropTPScreen = AUIUtilities.FindElementByID(DropTPSreenId, aeForm);
                            Console.WriteLine("DropTransportScreen found: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(3000);
                        if (aeDropTPScreen == null)
                        {
                            sErrorMessage = "DropTransportScreen not found";
                            Console.WriteLine("Drop FindElementByID failed:" + DropTPSreenId);
                        }
                        else
                        {
                            /*WriteLog(testInfoTxtFile, "DropTPScreen found, now find TblContent, scan and select button");
                            // find tblContent
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeTblContent == null && mTime.Minutes < 10)
                            {
                                aeTblContent = AUIUtilities.FindElementByID(ContentId, aeDropTPScreen);
                                mTime = DateTime.Now - mAppTime;
                                WriteLog(testInfoTxtFile, "Tbl Content  find time is :" + mTime.TotalMilliseconds / 1000);
                                Thread.Sleep(2000);
                            }

                            if (aeTblContent == null)
                            {
                                Log(sErrorMessage);
                                WriteLog(testInfoTxtFile, "Drop TblContent not found:");
                                MessageBox.Show(sErrorMessage, "Drop");
                                Thread.Sleep(3600000 * 20);
                            }*/

                            // OK 
                            AutomationElement aeDropOKButton = AUIUtilities.FindElementByID(DropOKId, aeDropTPScreen);
                            if (aeDropOKButton == null)
                            {
                                sErrorMessage = "Drop OK button not found";
                                Console.WriteLine("Drop OK button not found:" + DropOKId);
                                Log("Drop OK button not found:" + DropOKId);
                            }
                            else
                            {
                                sDropFTCount++;
                                Thread.Sleep(500);
                                Input.MoveTo(aeDropOKButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeDropOKButton));
                                slogMsg = "(D.OK)--> --> --> --> Total Drop ========== "+sDropFTCount;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("enter"))
                    {
                        #region  // Handle enter the location where the reel has been dropped
                        slogMsg = "(E)--> Enter location is --> " + DestLoc;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        // Drop transport
                        string ContentId = "m_TblContent";
                        string DropOKId = "m_BtnOK";

                        #region  // check Drop loc empty  region
                        if (CheckDropLocEmpty && DestLoc.ToLower().StartsWith("hu"))
                        {
                            // Wait until FIL input empty
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            bool FilInputEmpty = false;
                            while (FilInputEmpty == false)
                            {

                                if (IsDropLocationEmpty("HU.1"))
                                {
                                    FilInputEmpty = true;
                                    slogMsg = "(E.1)-->  : " + DestLoc + " ::HU.1--> empty, " + System.DateTime.Now;

                                }
                                else
                                {
                                    FilInputEmpty = false;
                                    mTime = DateTime.Now - mAppTime;
                                    slogMsg = "(E.1)-->  : " + DestLoc + " ::HU.1--> not empty," + System.DateTime.Now;
                                }
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                                Thread.Sleep(5000);
                            }
                        }
                        #endregion

                        // Find the  transport screen  DropTransportScreen
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;

                        AutomationElement aeDropTPScreen = null;
                        string DropScanId = "m_TxtScannedValue";  // edit control
                        AutomationElement aeDropField = null;
                        AutomationElement aeTblContent = null;
                        while (aeDropTPScreen == null && mTime.Minutes < 5)
                        {
                            Console.WriteLine("Find DropTransportScreen : " + System.DateTime.Now);
                            aeDropTPScreen = AUIUtilities.FindElementByID(DropTPSreenId, aeForm);
                            Console.WriteLine("DropTransportScreen found: " + System.DateTime.Now);
                            mTime = DateTime.Now - mAppTime;
                            Console.WriteLine(" find time is :" + mTime.TotalMilliseconds / 1000);
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(3000);
                        if (aeDropTPScreen == null)
                        {
                            sErrorMessage = "DropTransportScreen not found";
                            Console.WriteLine("Drop FindElementByID failed:" + DropTPSreenId);
                        }
                        else
                        {
                            WriteLog(testInfoTxtFile, "DropTPScreen found, now find TblContent, scan and select button");
                            // find tblContent
                            mAppTime = DateTime.Now;
                            mTime = DateTime.Now - mAppTime;
                            while (aeTblContent == null && mTime.Minutes < 10)
                            {
                                aeTblContent = AUIUtilities.FindElementByID(ContentId, aeDropTPScreen);
                                mTime = DateTime.Now - mAppTime;
                                WriteLog(testInfoTxtFile, "Tbl Content  find time is :" + mTime.TotalMilliseconds / 1000);
                                Thread.Sleep(2000);
                            }

                            if (aeTblContent == null)
                            {
                                Log(sErrorMessage);
                                WriteLog(testInfoTxtFile, "Drop TblContent not found:");
                                MessageBox.Show(sErrorMessage, "Drop");
                                Thread.Sleep(3600000 * 20);
                            }

                            // Scan location  
                            // "m_TxtScannedValue"  This is Edit Control, should use setFocus + sendKeys
                            //if (AUIUtilities.FindDocumentAndSendText(DropScanId, aeDropTPScreen, "HU.1", ref sErrorMessage))
                            if (aeTblContent.Current.Name.StartsWith("Enter"))
                            {
                                mAppTime = DateTime.Now;
                                mTime = DateTime.Now - mAppTime;
                                while (aeDropField == null && mTime.Minutes < 3)
                                {
                                    aeDropField = AUIUtilities.FindElementByID(DropScanId, aeDropTPScreen);
                                    mTime = DateTime.Now - mAppTime;
                                    WriteLog(testInfoTxtFile, "drop field  find time is :" + mTime.TotalMilliseconds / 1000);
                                    Thread.Sleep(2000);
                                }
                                //aeDropField = AUIUtilities.FindElementByID(DropScanId, aeDropTPScreen);
                                if (aeDropField != null)
                                {
                                    Thread.Sleep(3000);
                                    aeDropField.SetFocus();
                                    Thread.Sleep(1000);
                                    System.Windows.Forms.SendKeys.SendWait("HU.1");
                                    Thread.Sleep(2000);
                                    //ValuePattern vp = (ValuePattern)aeDropField.GetCurrentPattern(ValuePattern.Pattern);
                                    //Thread.Sleep(2000);
                                    //vp.SetValue("HU.1");
                                    Console.WriteLine("HU.1 scanned");
                                    Log("HU.1  OK");
                                }
                                else
                                {
                                    Log(sErrorMessage);
                                    WriteLog(testInfoTxtFile, "Drop ScanID not found:");
                                    MessageBox.Show(sErrorMessage, "Drop to HU.1");
                                    Thread.Sleep(3600000 * 20);
                                }
                            }
                            Thread.Sleep(500);

                            // OK 
                            AutomationElement aeDropOKButton = AUIUtilities.FindElementByID(DropOKId, aeDropTPScreen);
                            if (aeDropOKButton == null)
                            {
                                sErrorMessage = "Drop OK button not found";
                                Console.WriteLine("Drop OK button not found:" + DropOKId);
                                Log("Drop OK button not found:" + DropOKId);
                            }
                            else
                            {
                                sDropFTCount++;
                                Thread.Sleep(500);
                                Input.MoveTo(aeDropOKButton);
                                Thread.Sleep(1000);
                                Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeDropOKButton));
                                slogMsg = "(E.OK)--> --> --> --> Total Drop ==========" + sDropFTCount;
                                Log(slogMsg);
                                WriteLog(testInfoTxtFile, slogMsg);
                            }
                        }
                        #endregion
                    }
                    else if (instructionName.ToLower().StartsWith("transport could not be selected"))
                    {
                        #region  // Handle transport could not be selected screen
                        slogMsg = "(T)--> transport could not be selected, refresh screen --> ";
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);

                        // Find the transport refresh screen button
                        Thread.Sleep(500);
                        mAppTime = DateTime.Now;
                        mTime = DateTime.Now - mAppTime;
                        AutomationElement aeBtnRefresh = null;
                        string RefreshBtnId = "m_BtnRefresh";  // refresh button
                        while (aeBtnRefresh == null && mTime.Minutes < 5)
                        {
                            aeBtnRefresh = AUIUtilities.FindElementByID(RefreshBtnId, aeForm);
                            mTime = DateTime.Now - mAppTime;
                            Thread.Sleep(1000);
                        }

                        Thread.Sleep(1000);
                        if (aeBtnRefresh == null)
                        {
                            sErrorMessage = "Screen refresh button not found";
                            MessageBox.Show(sErrorMessage, "Transport could not be selected");
                        }
                        else
                        {
                            WriteLog(testInfoTxtFile, "transport could not be selected, refresh screen");
                            Thread.Sleep(500);
                            Input.MoveTo(aeBtnRefresh);
                            Thread.Sleep(1000);
                            Input.ClickAtPoint(AUIUtilities.GetElementCenterPoint(aeBtnRefresh));
                            slogMsg = "(T.OK)--> --> --> --> Next";
                            Log(slogMsg);
                            WriteLog(testInfoTxtFile, slogMsg);
                        }
                        #endregion
                    }
                    else if (instructionName.Equals(string.Empty))
                    {
                        slogMsg = " *** (Empty)--> Handle Empty instruction name : " + instructionName;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        //MessageBox.Show(slogMsg, instructionName);
                    }
                    else 
                    {
                        slogMsg = " *** (U)--> Handle unexpected instruction screens : " + instructionName;
                        Log(slogMsg);
                        WriteLog(testInfoTxtFile, slogMsg);
                        MessageBox.Show(slogMsg, instructionName);
                    }
                }
                else
                {
                    sErrorMessage = "Instruction lable not found:";
                    Log(sErrorMessage);
                    WriteLog(testInfoTxtFile, "Instruction lable not found:");
                    MessageBox.Show(sErrorMessage, "Get instruction screen");
                }
                Thread.Sleep(2000);
                #endregion
                Thread.Sleep(1000);
            }

        }

        public enum STATE
        {
            UNDEFINED,
            PENDING,
            INPROGRESS,
            EXCEPTION,
        }


        public bool IsDropLocationEmpty(string locID)
        {
            bool isEmpty = false;
            try
            {
                SqlConnection myConnection = new SqlConnection(sConnectionString);

                try
                {
                    myConnection.Open();
                }
                catch (Exception e)
                {
                    string exc = e.Message + System.Environment.NewLine + e.StackTrace;
                    MessageBox.Show(exc, "open sql connection:");
                    Thread.Sleep(3600*1000*20);
                }

                string sqlCommand = "SELECT COUNT(*) FROM [Ewcs].[dbo].[Carriers], [Ewcs].[dbo].[Locations] "+
                            "Where [Ewcs].[dbo].[Carriers].LocationId = [Ewcs].[dbo].[Locations].LocationId "+
                            "and [Ewcs].[dbo].[Locations].LocationId = '" + locID + "'";
                SqlCommand myCommand = new SqlCommand(sqlCommand, myConnection);
                int count = (int)myCommand.ExecuteScalar();
                if (count == 0 )
                {
                    isEmpty = true;
                }
                else
                {
                    isEmpty = false;
                }

                Thread.Sleep(2000);
                myConnection.Close();
                return isEmpty;
            }
            catch (Exception ex)
            {
                sErrorMessage = sErrorMessage + " CONN:" + sConnectionString;
                sErrorMessage = ex.Message + "----: " + ex.StackTrace;
                MessageBox.Show(sErrorMessage, "Check " + locID + "  sql exception");
                Thread.Sleep(3600000 * 20);
                return isEmpty;
            }

        }
        
        public void WriteLog(string slogFilePath, string msg)
        {
            try
            {
                System.IO.StreamWriter sw = null;
                try
                {
                    sw = System.IO.File.AppendText(slogFilePath);
                    //Path.Combine( logFilePath, logFileName ));
                    string logLine = System.String.Format(
                        "{0:G}: {1}.", System.DateTime.Now, "\t" + msg);
                    sw.WriteLine(logLine);
                }
                finally
                {
                    sw.Close();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public static bool FindDocumentAndSendText(String automationID, AutomationElement rootAE,
         string setValue, ref string msg)
        {
            Console.WriteLine("FindTextBoxAndChangeValue: " + automationID);
            try
            {
                AutomationElement aeTextBox = AUIUtilities.FindElementByID(automationID, rootAE);
                if (aeTextBox != null)
                {
                    Thread.Sleep(500);
                    System.Windows.Point pnt = AUIUtilities.GetElementCenterPoint(aeTextBox);
                    Input.MoveTo(pnt);
                    Thread.Sleep(1000);

                    aeTextBox.SetFocus();
                    Thread.Sleep(3000);
                    System.Windows.Forms.SendKeys.SendWait(setValue);
                    Thread.Sleep(2000);
                    // Check Field value, NOT CHECK, Application will va;idate input
                    //TextPattern tp = (TextPattern)aeTextBox.GetCurrentPattern(TextPattern.Pattern);
                    //Thread.Sleep(1000);
                    //string v = tp.DocumentRange.GetText(-1).Trim();
                    //Thread.Sleep(500);
                    //Console.WriteLine("filled text is : " + v);
                    //if (v.Equals(setValue))
                    //{
                        return true;
                    //}
                    //else
                    //{
                    //    msg = "input value  not correct" + v;
                    //    return false;
                    //}
                }
                else
                {
                    msg = automationID + " not found";
                    return false;
                }
            }
            catch (Exception ex)
            {
                msg = ex.Message + ":" + ex.StackTrace;
                return false;
            }
        }

    }
}
