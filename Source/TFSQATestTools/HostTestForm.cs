using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using TestTools;
//using System.Collections.Specialized;


namespace TFSQATestTools
{
    public partial class HostTestForm : Form
    {
        private static List<DataObject> list = new List<DataObject>();

        public HostTestForm()
        {
            InitializeComponent();
        }

        private void btnFileBrowser_Click(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog();
            //dlg.FileName = "Document"; // Default file name 
            dlg.DefaultExt = ".txt"; // Default file extension 
            dlg.Filter = "Text documents (.txt)|*.txt"; // Filter files by extension 
            // Show open file dialog box
            DialogResult result = dlg.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                txtHostTestInputFile.Text = dlg.FileName;
            }
            Console.WriteLine(result); // <-- 
        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            list = new List<DataObject>();
            StreamReader reader = File.OpenText(txtHostTestInputFile.Text);
            string infoline = reader.ReadLine();
            DataObject data1 = null;
            while (infoline != null)
            {
                string[] strArray = infoline.Split(',');
                data1 = new DataObject
                            {
                                Command = strArray[0],
                                Source = strArray[1],
                                Destination = strArray[2],
                            };
                list.Add(data1);
                infoline = reader.ReadLine();
            }
            reader.Close();
            dataGridView.DataSource = list;
        }

        private void btnHostTestStart_Click(object sender, EventArgs e)
        {
            IEnumerator EmpEnumerator = list.GetEnumerator(); //Getting the Enumerator
            int totalTransportsCount = 100;
            int transportsTimeIntervalSec = 30;

            string totalTransports = textBoxTotalTransports.Text;
            string transportTimeInterval = textBoxTransportTimeInterval.Text;

            // ToInt32 can throw FormatException or OverflowException. 
            try
            {
                totalTransportsCount = Convert.ToInt32(totalTransports);
            }
            catch (FormatException ex)
            {
                Console.WriteLine("Input string is not a sequence of digits.");
                MessageBox.Show("Input string is not a sequence of digits: "+ex.Message, "Validate input");
                return;
            }
            catch (OverflowException ex2)
            {
                Console.WriteLine("The number cannot fit in an Int32: "+ex2.Message);
            }
            finally
            {
                if (totalTransportsCount < Int32.MaxValue)
                {
                    Console.WriteLine("The new value is {0}", totalTransportsCount + 1);
                }
                else
                {
                    Console.WriteLine("numVal cannot be incremented beyond its current value");
                }
            }

            // ToInt32 can throw FormatException or OverflowException. 
            try
            {
                transportsTimeIntervalSec = Convert.ToInt32(transportTimeInterval);
            }
            catch (FormatException ex3)
            {
                Console.WriteLine("Input string is not a sequence of digits.");
                MessageBox.Show("Input string is not a sequence of digits: "+ex3.Message, "Validate tieminterval input");
                return;
            }
            catch (OverflowException ex4)
            {
                Console.WriteLine("The number cannot fit in an Int32: "+ex4.Message);
            }
            finally
            {
                if (transportsTimeIntervalSec < Int32.MaxValue)
                {
                    Console.WriteLine("The new value is {0}", transportsTimeIntervalSec + 1);
                }
                else
                {
                    Console.WriteLine("numVal cannot be incremented beyond its current value");
                }
            }

            int dataCount = list.Count;
            for (int i = 0; i < totalTransportsCount; i++)
            {
                EmpEnumerator.Reset(); //Position at the Beginning
                while (EmpEnumerator.MoveNext()) //Till not finished do print
                {
                    if (i >= totalTransportsCount)
                        break;

                    var transportData = (DataObject) EmpEnumerator.Current;
                    string comm = transportData.Command;
                    string src = transportData.Source;
                    string dest = transportData.Destination;
                    SendTransport(comm, src, dest);
                    Thread.Sleep(transportsTimeIntervalSec*1000);
                    //tester.Log(i + " * " + comm + " --- " + src + " --- " + dest);
                    i++;
                }
                i--;
            }
        }

        private void SendTransport(string command, string src, string destination)
        {
            AutomationElement aeMainForm = ProjBasicUI.GetMainWindowWithinTime("frmMain", 30);
            if (aeMainForm != null)
            {
                AutomationElement aeTab = AUIUtilities.FindElementByID("tabControl", aeMainForm);
                if (aeTab != null)
                {
                    Condition cTabItem = new AndCondition(
                        new PropertyCondition(AutomationElement.NameProperty, "Transports"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem)
                        );
                    AutomationElement aeTransportTabItem = aeTab.FindFirst(TreeScope.Children, cTabItem);
                    if (aeTab != null)
                    {
                        Input.MoveToAndClick(aeTransportTabItem);
                        AutomationElement aeSrc = AUIUtilities.FindElementByID("From", aeTransportTabItem);
                        AutomationElement aeDest = AUIUtilities.FindElementByID("To", aeTransportTabItem);
                        Input.MoveToAndClick(aeSrc);
                        ProjBasicUI.SendTextToElement(aeSrc, src);
                        Input.MoveToAndClick(aeDest);
                        ProjBasicUI.SendTextToElement(aeDest, destination);
                        AutomationElement aeSendTransportButton = AUIUtilities.FindElementByName("Send Transport",
                                                                                                 aeTransportTabItem);
                        if (aeSendTransportButton != null)
                        {
                            Input.MoveToAndClick(aeSendTransportButton);
                        }
                    }
                }
            }
        }

        #region Nested type: DataObject

        private class DataObject
        {
            public string Command { get; set; }
            public string Source { get; set; }
            public string Destination { get; set; }
        }

        #endregion
    }
}