using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DuurTesten
{
    public partial class TestForm : Form
    {
        public static string sApplication = string.Empty;
        public static DateTime sStartUpTime;
        
        public TestForm()
        {
            InitializeComponent();
        }

        [STAThread]
        static void Main(string[] args)
        {
            //try
            //{
            TestTools.Utilities.CloseProcess("EXCEL");
            TestTools.Utilities.CloseProcess("Egemin.Epia.Foundation.ComponentManagement.Host");
            Application.Run(new TestForm());
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(" Application exception" + ex.Message, ex.StackTrace);
            //}
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            DateTime sStartUpTime = DateTime.Now;
            lbStartTime.Text = sStartUpTime.ToString("yyyyMMdd HH:mm:ss");
            DateTime sEndTime = sStartUpTime.AddHours(1);
            lbEndTime.Text = sEndTime.ToString("yyyyMMdd HH:mm:ss");
            sApplication = cmbApplication.SelectedItem.ToString();

            TimeSpan ttl = sEndTime - sStartUpTime;

        }

        private void btnStop_Click(object sender, EventArgs e)
        {
        }

        
    }
}
