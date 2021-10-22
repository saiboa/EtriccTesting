using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;

namespace TFSQATestTools
{
    public partial class GetBuildDefinitionsForm : Form
    {
        TfsTeamProjectCollection tfsProjectCollection;
        string selectedProject = "Epia 4";
        string[] buildDefs;
        string SelectedDateFilter;

        public GetBuildDefinitionsForm()
        {
        }

        public GetBuildDefinitionsForm(string prj)
        {
            InitializeComponent();

            cmbTFSServer.Items.Add(Tfs.Server);
            cmbTFSServer.SelectedIndex = 0;

            Uri serverUri = new Uri(Tfs.ServerUrl);
            System.Net.ICredentials tfsCredentials
                = new System.Net.NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

            tfsProjectCollection
                = new TfsTeamProjectCollection(serverUri, tfsCredentials);

            if (prj.StartsWith("Epia"))
                selectedProject = "Epia 4";
            else if (prj.StartsWith("Etricc"))
                selectedProject = "Etricc 5";
            else if (prj.StartsWith("Ewms"))
                selectedProject = "Epia 3";
                
            lstBoxProject.Items.Add(selectedProject);

            IBuildServer buildServer;
            tfsProjectCollection.EnsureAuthenticated();
            buildServer = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));

            //IBuildDefinition[] buildDefinitions = buildServer.QueryBuildDefinitions(tpp.SelectedProjects[0].Name);
            IBuildDefinition[] buildDefinitions = buildServer.QueryBuildDefinitions(selectedProject);

            int x = 0;
            for (int i = 0; i < buildDefinitions.Length; i++)
            {
                if (buildDefinitions[i].Name.StartsWith("OEM")
                    || buildDefinitions[i].Name.StartsWith("Tools")
                    || buildDefinitions[i].Name.StartsWith("Egv")
                    || buildDefinitions[i].Name.StartsWith("Etricc Stat")
                    || buildDefinitions[i].Name.StartsWith("Etricc Temp")
                    || buildDefinitions[i].Name.StartsWith("Testing")
                    || buildDefinitions[i].Name.IndexOf("use") > 0
                    || buildDefinitions[i].Name.IndexOf("Stat") > 0
                    || buildDefinitions[i].Name.StartsWith("Epia.Development.Dev07.Version") 
                    )
                    x++;
                else
                {
                    lstBoxDuildDefinition.Items.Add(buildDefinitions[i].Name);
                }
            }
        }

        public GetBuildDefinitionsForm(string prj, string testApp, ref List<string> selectedDefs, ref string dateFilter)
        {
            InitializeComponent();

            cmbTFSServer.Items.Add(Tfs.Server);
            cmbTFSServer.SelectedIndex = 0;
            cmbDateFilter.Text = dateFilter;

            Uri serverUri = new Uri(Tfs.ServerUrl);
            System.Net.ICredentials tfsCredentials
                 = new System.Net.NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

            tfsProjectCollection
                = new TfsTeamProjectCollection(serverUri, tfsCredentials);

            if (prj.StartsWith("Epia"))
                selectedProject = "Epia 4";
            else if (prj.StartsWith("Etricc"))
                selectedProject = "Etricc 5";
            else if (prj.StartsWith("Ewms"))
                selectedProject = "Epia 3";

            lstBoxProject.Items.Add(selectedProject);
            listBoxTestApp.Items.Add(testApp);

            IBuildServer buildServer;
            tfsProjectCollection.EnsureAuthenticated();
            buildServer = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));

            //IBuildDefinition[] buildDefinitions = buildServer.QueryBuildDefinitions(tpp.SelectedProjects[0].Name);
            IBuildDefinition[] buildDefinitions = buildServer.QueryBuildDefinitions(selectedProject);

            int x = 0;
            for (int i = 0; i < buildDefinitions.Length; i++)
            {
                if (buildDefinitions[i].Name.StartsWith("OEM")
                    || buildDefinitions[i].Name.StartsWith("Tools")
                    || buildDefinitions[i].Name.StartsWith("Egv")
                    || buildDefinitions[i].Name.StartsWith("Etricc Stat Rep")
                    || buildDefinitions[i].Name.StartsWith("Etricc Temp")
                    || buildDefinitions[i].Name.StartsWith("Testing")
                    || buildDefinitions[i].Name.IndexOf("use") > 0
                    || buildDefinitions[i].Name.StartsWith("Epia.Development.Dev07.Version") 
                   // || buildDefinitions[i].Name.IndexOf("Stat") > 0
                    )
                    x++;
                else
                {

                    if (testApp.StartsWith("EtriccStatistics"))
                    {
                        if (buildDefinitions[i].Name.StartsWith("Etricc Stat Prog"))
                            lstBoxDuildDefinition.Items.Add(buildDefinitions[i].Name);
                    }
                    else
                        if (!buildDefinitions[i].Name.StartsWith("Etricc Stat Prog"))
                            lstBoxDuildDefinition.Items.Add(buildDefinitions[i].Name);
                }
            }

            for (int j = 0; j < lstBoxDuildDefinition.Items.Count; j++)
            {
                if (isBuildDefSelected(lstBoxDuildDefinition.Items[j].ToString(), ref selectedDefs))
                    lstBoxDuildDefinition.SelectedIndex = j;
            }

        }

        public GetBuildDefinitionsForm(string[] testApps, ref List<string> selectedDefs, ref string dateFilter)
        {
            InitializeComponent();

            cmbTFSServer.Items.Add(Tfs.Server);
            cmbTFSServer.SelectedIndex = 0;
            cmbDateFilter.Text = dateFilter;

            Uri serverUri = new Uri(Tfs.ServerUrl);
            System.Net.ICredentials tfsCredentials
                 = new System.Net.NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

            tfsProjectCollection
                = new TfsTeamProjectCollection(serverUri, tfsCredentials);

            IBuildServer buildServer;
            tfsProjectCollection.EnsureAuthenticated();
            buildServer = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));


            for (int ix = 0; ix < testApps.Length; ix++)
            {
                if (testApps[ix].Equals(TestApp.EPIA4) || testApps[ix].Equals(TestApp.EPIANET45))
                    selectedProject = "Epia 4";
                else
                    selectedProject = "Etricc 5";

                //IBuildDefinition[] buildDefinitions = buildServer.QueryBuildDefinitions(tpp.SelectedProjects[0].Name);
                IBuildDefinition[] buildDefinitions = buildServer.QueryBuildDefinitions(selectedProject);

                int x = 0;
                for (int i = 0; i < buildDefinitions.Length; i++)
                {
                    if (buildDefinitions[i].Name.StartsWith("OEM")
                        || buildDefinitions[i].Name.StartsWith("Tools")
                        || buildDefinitions[i].Name.StartsWith("Egv")
                        || buildDefinitions[i].Name.StartsWith("Etricc Stat Rep")
                        || buildDefinitions[i].Name.StartsWith("Etricc Temp")
                        || buildDefinitions[i].Name.StartsWith("Testing")
                        || buildDefinitions[i].Name.IndexOf("use") > 0
                        || buildDefinitions[i].Name.StartsWith("Obsolete")
                        // || buildDefinitions[i].Name.IndexOf("Stat") > 0
                        || buildDefinitions[i].Name.StartsWith("Epia.Development.Dev07.Version") 
                        )
                        x++;
                    else
                    {
                        if (testApps[ix].StartsWith(TestApp.ETRICCSTATISTICS))
                        {
                            if (buildDefinitions[i].Name.StartsWith("Etricc Stat Prog"))
                                lstBoxDuildDefinition.Items.Add(buildDefinitions[i].Name);
                        }
                        else if (testApps[ix].StartsWith(TestApp.EPIANET45))
                        {
                            if (buildDefinitions[i].Name.IndexOf("Dev045") > 0 )
                                lstBoxDuildDefinition.Items.Add(buildDefinitions[i].Name);
                        }
                        else if (testApps[ix].StartsWith(TestApp.EPIA4))
                        {
                            if (buildDefinitions[i].Name.IndexOf("Dev045") < 0)
                                lstBoxDuildDefinition.Items.Add(buildDefinitions[i].Name);
                        }
                        else if (testApps[ix].StartsWith(TestApp.ETRICCUI))
                        {
                            if (!buildDefinitions[i].Name.StartsWith("Etricc Stat Prog") && buildDefinitions[i].Name.IndexOf("Net45") < 0)
                                lstBoxDuildDefinition.Items.Add(buildDefinitions[i].Name);
                        }
                        else if (testApps[ix].StartsWith(TestApp.ETRICCNET45))
                        {
                            if (buildDefinitions[i].Name.IndexOf("Net45") > 0)
                                lstBoxDuildDefinition.Items.Add(buildDefinitions[i].Name);
                        }
                    }
                }
            }

            for (int j = 0; j < lstBoxDuildDefinition.Items.Count; j++)
            {
                if (isBuildDefSelected(lstBoxDuildDefinition.Items[j].ToString(), ref selectedDefs))
                    lstBoxDuildDefinition.SelectedIndex = j;
            }

        }

        private bool isBuildDefSelected(string def, ref List<string> SelecedDefs)
        {
            bool selected = false;
            IEnumerator EmpEnumerator = SelecedDefs.GetEnumerator();
            EmpEnumerator.Reset();
            while (EmpEnumerator.MoveNext())
            {
                if (def.Equals((string)EmpEnumerator.Current))
                {
                    selected = true;
                    break;
                }
            }
            return selected;
        }

        private void btnConn_Click(object sender, EventArgs e)
        {
            IBuildServer buildServer;
            tfsProjectCollection.EnsureAuthenticated();
            buildServer = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));


            if (lstBoxDuildDefinition.SelectedItems.Count == 0)
            {
                System.Windows.MessageBox.Show("Please select a build definition");
                btnConn.DialogResult = System.Windows.Forms.DialogResult.OK;
                // Call back to the parent passing the entire CustomerDialog Instance
                Close();
            }

            buildDefs = new string[lstBoxDuildDefinition.SelectedItems.Count];
            for (int k = 0; k < lstBoxDuildDefinition.SelectedItems.Count; k++)
            {
                //System.Windows.MessageBox.Show("selectedBuldDefinition=" + lstBoxDuildDefinition.SelectedItems[k].ToString());
                buildDefs[k] = lstBoxDuildDefinition.SelectedItems[k].ToString();
                //IBuildDetail[] buildnrs = buildDefinitions[7].QueryBuilds();
                //IBuildDetail[] buildnrs = buildServer.QueryBuilds(selectedProject, lstBoxDuildDefinition.SelectedItems[k].ToString());
                //string bnrs = string.Empty;
                //for (int i = 0; i < buildnrs.Length; i++)
                //{
                //    bnrs = bnrs + "\n " + buildnrs[i].BuildNumber + "\t" + buildnrs[i].Quality + "\t" + buildnrs[i].DropLocation;
                //}
                //System.Windows.MessageBox.Show("Connection OK \n Build definition Name are " + tfsProjectCollection.Name + " \nwith build nrs:" + bnrs);
            }

            SelectedDateFilter = cmbDateFilter.SelectedItem.ToString();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        public string[] getBuildDefinition()
        {
            return buildDefs;
        }

        public string getDateFilter()
        {
            return SelectedDateFilter;
        }

        private void cmbDateFilter_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        ///  if checkbox checked --> all ListBox items are selected
        ///  if checkbox unchecked --> all ListBox items are unselected
        ///  When ListBox changed, if ListBox not all selected or NOT all unselected, 
        ///  then set checkbox state to Indeterminate
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkBuildDefs_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBuildDefs.CheckState == CheckState.Checked)
            {
                for (int i = 0; i < lstBoxDuildDefinition.Items.Count; i++) 
                    lstBoxDuildDefinition.SetSelected(i, true); 
            }
            else if (chkBuildDefs.CheckState == CheckState.Unchecked)
            {
                for (int i = 0; i < lstBoxDuildDefinition.Items.Count; i++)
                    lstBoxDuildDefinition.SetSelected(i, false);
            }
            else if (chkBuildDefs.CheckState == CheckState.Indeterminate)
            {
                // this state is set by ListBox changed event
            }
             
        }

        private void lstBoxDuildDefinition_SelectedIndexChanged(object sender, EventArgs e)
        {
            if  (lstBoxDuildDefinition.SelectedItems.Count == 0)
                chkBuildDefs.CheckState = CheckState.Unchecked;
            else if (lstBoxDuildDefinition.SelectedItems.Count == lstBoxDuildDefinition.Items.Count)
                chkBuildDefs.CheckState = CheckState.Checked;
            else
                chkBuildDefs.CheckState = CheckState.Indeterminate;
        }
    }
}
