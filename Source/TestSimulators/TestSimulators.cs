using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;

//using System.Windows;
using System.Windows.Automation;
using TestTools;
using System.Windows.Input;
//using System.Windows.Forms;

namespace TestSimulators
{
    public partial class TestSimulators : Form
    {
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Fields of TestSimulators : System.Windows.Forms.Form (1)
        public ScannerSimulators Scanner = new ScannerSimulators();
        #endregion // —— Fields •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Constructors/Destructors/Cleanup of TestSimulators : System.Windows.Forms.Form (1)
        public TestSimulators()
        {
            InitializeComponent();
        }
        #endregion // —— Constructors/Destructors/Cleanup •••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

        [STAThread]
        static void Main(string[] args)
        {
            Application.Run(new TestSimulators());
        }

        private void BtnTestFile_Click(object sender, EventArgs e)
        {
            // ------------  if write infotext file failure, not do anything , continue go to iteration 
            string testInfoTxtFile = Path.Combine(@"C:\KC\PutAway", "Map3000Reels.csv");
            StreamWriter writeInfo = File.CreateText(testInfoTxtFile);

            string info = "UnitId,ProductId,Batch,Diameter,CoreDiameter,Weight,PaperWidth,PaperLength,PaperBulk,Section";
            try
            {
                writeInfo.WriteLine(info);
                for (int i = 0; i < 1000; i++)
                {
                    string id = String.Format("{0:000}", i);
                    info = "BEPUT_20091001-" + id + "A,Z5705035,BATCH-000-01,2600,350,2000,2600,100000,9.01,A";
                    writeInfo.WriteLine(info);
                }

                for (int i = 0; i < 1000; i++)
                {
                    string id = String.Format("{0:000}", i);
                    info = "FRPUT_20091002-" + id + "A,Z5705035,BATCH-000-01,2600,350,2000,2600,100000,9.01,A";
                    writeInfo.WriteLine(info);
                }

                for (int i = 0; i < 1000; i++)
                {
                    string id = String.Format("{0:000}", i);
                    info = "NLPUT_20091003-" + id + "A,Z5705035,BATCH-000-01,2600,350,2000,2600,100000,9.01,A";
                    writeInfo.WriteLine(info);
                }

                for (int i = 0; i < 3000; i++)
                {
                    string id = String.Format("{0:0000}", i);
                    info = "( 'P" + id + "', 'Pallet', 'FIL.3', '0', 'OK' ),";
                    writeInfo.WriteLine(info);
                }
            }
            catch (Exception ex)
            {
                string msg = " Exception Add pc to InfoText:" + "=====" + ex.Message + "" + ex.StackTrace;
                MessageBox.Show(msg);

            }
            writeInfo.Close();
        }

        private void btnForkLift_Click(object sender, EventArgs e)
        {
            Scanner.ForkliftScannerStart("", "HU.1");
        }
        
        private void btnProngLift_Click(object sender, EventArgs e)
        {
            Scanner.CheckDropLocEmpty = chkFILempty.Checked;
            Scanner.ProngliftScannerStart("MAP.1", "FIL.1");
        }

        private void btnClearLogScreen_Click(object sender, EventArgs e)
        {
            TxtOutput.Text = String.Empty;
            Scanner.m_Logging.Clear();
        }

        private void TestSimulators_Load(object sender, System.EventArgs e)
        {
            Scanner.OnLoggingChanged += new EventHandler(scanner_OnLoggingChanged);
        }

        private void scanner_OnLoggingChanged(object sender, System.EventArgs e)
        {
            // make threadsafe
            if (InvokeRequired)
            {
                BeginInvoke(new scanner_OnLoggingChangedDelegate(scanner_OnLoggingChanged), new object[] { sender, e });
                return;
            }
            //update the multiline textBox
            string newLog = string.Empty;

            foreach (string logLine in Scanner.Logging)
                newLog += logLine + Environment.NewLine;

            TxtOutput.Text = newLog;
        }

        private delegate void scanner_OnLoggingChangedDelegate(object sender, System.EventArgs e);

        private void btnStop_Click(object sender, EventArgs e)
        {
            Scanner.State = ScannerSimulators.STATE.PENDING;
        }

        private void chkFILempty_CheckedChanged(object sender, EventArgs e)
        {
            if (chkFILempty.Checked)
                Scanner.CheckDropLocEmpty = true;
            else
                Scanner.CheckDropLocEmpty = false;
                
        }

        private void btnSmallProngLift_Click(object sender, EventArgs e)
        {
            Scanner.CheckDropLocEmpty = chkFILempty.Checked;
            Scanner.SmallProngliftScannerStart("KCP1.1", "KCP1.AV1");
        }

       
       

    }
}
