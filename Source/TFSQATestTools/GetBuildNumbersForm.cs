using System;
using System.Collections;
using System.Collections.Generic;
using System.Net;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using TestTools;
using MessageBox = System.Windows.MessageBox;

namespace TFSQATestTools
{
    public partial class GetBuildNumbersForm : Form
    {
        private readonly string selectedProject = "Epia 4";
        private readonly TfsTeamProjectCollection tfsProjectCollection;
        private string[] buildDefs;

        public GetBuildNumbersForm()
        {
        }

        public GetBuildNumbersForm(string prj)
        {
            InitializeComponent();

            cmbTFSServer.Items.Add(Tfs.Server);
            cmbTFSServer.SelectedIndex = 0;

            var serverUri = new Uri(Tfs.ServerUrl);
            ICredentials tfsCredentials
                = new NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

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
            buildServer = (IBuildServer) tfsProjectCollection.GetService(typeof (IBuildServer));

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
                    )
                    x++;
                else
                {
                    lstBoxDuildNumber.Items.Add(buildDefinitions[i].Name);
                }
            }
        }

        public GetBuildNumbersForm(string prj, List<string> selectedDefs, string dateFilter, string selectedBuildNr)
        {
            InitializeComponent();

            cmbTFSServer.Items.Add(Tfs.Server);
            cmbTFSServer.SelectedIndex = 0;

            var serverUri = new Uri(Tfs.ServerUrl);
            ICredentials tfsCredentials
                = new NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

            tfsProjectCollection
                = new TfsTeamProjectCollection(serverUri, tfsCredentials);

            if (prj.StartsWith("Epia"))
                selectedProject = "Epia 4";
            else if (prj.StartsWith("Etricc"))
                selectedProject = "Etricc 5";
            else if (prj.StartsWith("Ewms"))
                selectedProject = "Epia 3";

            lstBoxProject.Items.Add(selectedProject);

            DateTime timeNow = DateTime.Now;
            DateTime timeFrom = DateTime.Now;
            if (dateFilter.StartsWith("<Any Time>"))
                timeFrom = DateTime.MinValue;
            else if (dateFilter.StartsWith("Today"))
                timeFrom = DateTime.Today;
            else if (dateFilter.StartsWith("Last 24 hours"))
                timeFrom = DateTime.Now.AddHours(-24);
            else if (dateFilter.StartsWith("Last 48 hours"))
                timeFrom = DateTime.Now.AddHours(-48);
            else if (dateFilter.StartsWith("Last 7 days"))
                timeFrom = DateTime.Today.AddDays(-7);
            else if (dateFilter.StartsWith("Last 14 days"))
                timeFrom = DateTime.Today.AddDays(-14);
            else if (dateFilter.StartsWith("Last 28 days"))
                timeFrom = DateTime.Today.AddDays(-28);

            IBuildServer buildServer;
            tfsProjectCollection.EnsureAuthenticated();
            buildServer = (IBuildServer) tfsProjectCollection.GetService(typeof (IBuildServer));

            var allBuildsInfo = new Dictionary<string, string>();
            var allBuildObjects = new List<BuildObject>();

            foreach (string s in selectedDefs)
            {
                IBuildDetailSpec buildDetailSpec = buildServer.CreateBuildDetailSpec(selectedProject, s);
                //buildDetailSpec.MaxBuildsPerDefinition = 1; 
                buildDetailSpec.QueryOrder = BuildQueryOrder.FinishTimeDescending;
                buildDetailSpec.Status = BuildStatus.Succeeded; //Only get succeeded builds  
                buildDetailSpec.MinFinishTime = timeFrom;

                IBuildQueryResult results = buildServer.QueryBuilds(buildDetailSpec);
                //if (results.Failures.Length == 0 ) 
                //{ 
                //IBuildDetail buildDetail = results.Builds[0]; 
                //Console.WriteLine("Build: " + buildDetail.BuildNumber); 
                //Console.WriteLine("Account requesting build “ + 
                //“(build service user for triggered builds): " + buildDetail.RequestedBy); 
                //   Console.WriteLine("Build triggered by: " + buildDetail.RequestedFor); 
                //}

                IBuildDetail[] buildnrs = results.Builds;
                //IBuildDetail[] buildnrs = buildServer.QueryBuilds(selectedProject, s);
                string bnrs = string.Empty;
                string quality = string.Empty;
                var thisBuild = new BuildObject();

                for (int i = 0; i < buildnrs.Length; i++)
                {
                    //lstBoxDuildNumber.Items.Add(buildnrs[i].BuildNumber);

                    quality = buildnrs[i].Quality;
                    if (buildnrs[i].BuildNumber.ToLower().StartsWith("branch"))
                        continue;

                    string dropLocation = buildnrs[i].DropLocation;
                    string v = ConstCommon.DRIVE_MAP_LETTER + "\\" + dropLocation.Substring(55);

                    thisBuild = new BuildObject();
                    thisBuild.BuildNr = buildnrs[i].BuildNumber;
                    thisBuild.Quality = buildnrs[i].Quality;
                    thisBuild.FinishTime = buildnrs[i].FinishTime;
                    thisBuild.DripLoc = v;
                    allBuildObjects.Add(thisBuild);

                    // convert \\teamsystem.teamsystems.egemin.be\Team Systems Builds\CI... 
                    // to X:\CI...
                    //X:\CI\Epia 4\Epia.Development.Dev03-Net4.CI\Epia.Development.Dev03-Net4.CI_20110808.2
                    allBuildsInfo.Add(buildnrs[i].BuildNumber, v);
                    lstBoxDuildNumber.Items.Add(buildnrs[i].BuildNumber);
                }
            }

            //IBuildDefinition[] buildDefinitions = buildServer.QueryBuildDefinitions(tpp.SelectedProjects[0].Name);
            IBuildDefinition[] buildDefinitions = buildServer.QueryBuildDefinitions(selectedProject);


            for (int j = 0; j < lstBoxDuildNumber.Items.Count; j++)
            {
                if (lstBoxDuildNumber.Items[j].ToString().Equals(selectedBuildNr))
                    lstBoxDuildNumber.SelectedIndex = j;
            }
        }

        public GetBuildNumbersForm(string[] testApps, List<string> selectedDefs, string dateFilter,
                                   string selectedBuildNr)
        {
            InitializeComponent();

            cmbTFSServer.Items.Add(Tfs.Server);
            cmbTFSServer.SelectedIndex = 0;

            var serverUri = new Uri(Tfs.ServerUrl);
            ICredentials tfsCredentials
                = new NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

            tfsProjectCollection
                = new TfsTeamProjectCollection(serverUri, tfsCredentials);


            IBuildServer buildServer;
            tfsProjectCollection.EnsureAuthenticated();
            buildServer = (IBuildServer) tfsProjectCollection.GetService(typeof (IBuildServer));

            DateTime timeNow = DateTime.Now;
            DateTime timeFrom = DateTime.Now;
            if (dateFilter.StartsWith("<Any Time>"))
                timeFrom = DateTime.MinValue;
            else if (dateFilter.StartsWith("Today"))
                timeFrom = DateTime.Today;
            else if (dateFilter.StartsWith("Last 24 hours"))
                timeFrom = DateTime.Now.AddHours(-24);
            else if (dateFilter.StartsWith("Last 48 hours"))
                timeFrom = DateTime.Now.AddHours(-48);
            else if (dateFilter.StartsWith("Last 7 days"))
                timeFrom = DateTime.Today.AddDays(-7);
            else if (dateFilter.StartsWith("Last 14 days"))
                timeFrom = DateTime.Today.AddDays(-14);
            else if (dateFilter.StartsWith("Last 28 days"))
                timeFrom = DateTime.Today.AddDays(-28);


            for (int ix = 0; ix < testApps.Length; ix++)
            {
                if (testApps[ix].Equals(TestApp.EPIA4) || testApps[ix].Equals(TestApp.EPIANET45))
                    selectedProject = "Epia 4";
                else
                    selectedProject = "Etricc 5";

                lstBoxProject.Items.Add(selectedProject);
            }

            var allBuildsInfo = new Dictionary<string, string>();
            var allBuildObjects = new List<BuildObject>();
            foreach (string s in selectedDefs)
            {
                if (s.IndexOf("Epia") >= 0)
                    selectedProject = "Epia 4";
                else
                    selectedProject = "Etricc 5";

                IBuildDetailSpec buildDetailSpec = buildServer.CreateBuildDetailSpec(selectedProject, s);
                //buildDetailSpec.MaxBuildsPerDefinition = 1; 
                buildDetailSpec.QueryOrder = BuildQueryOrder.FinishTimeDescending;
                buildDetailSpec.Status = BuildStatus.Succeeded; //Only get succeeded builds  
                buildDetailSpec.MinFinishTime = timeFrom;
 
                IBuildQueryResult results = buildServer.QueryBuilds(buildDetailSpec);
                //if (results.Failures.Length == 0 ) 
                //{ 
                //IBuildDetail buildDetail = results.Builds[0]; 
                //Console.WriteLine("Build: " + buildDetail.BuildNumber); 
                //Console.WriteLine("Account requesting build “ + 
                //“(build service user for triggered builds): " + buildDetail.RequestedBy); 
                //   Console.WriteLine("Build triggered by: " + buildDetail.RequestedFor); 
                //}

                IBuildDetail[] buildnrs = results.Builds;
                //IBuildDetail[] buildnrs = buildServer.QueryBuilds(selectedProject, s);
                string bnrs = string.Empty;
                string quality = string.Empty;
                var thisBuild = new BuildObject();

                for (int i = 0; i < buildnrs.Length; i++)
                {
                    //lstBoxDuildNumber.Items.Add(buildnrs[i].BuildNumber);

                    quality = buildnrs[i].Quality;
                    if (buildnrs[i].BuildNumber.ToLower().StartsWith("branch"))
                        continue;

                    if (quality.Equals("Unexamined")
                        || quality.Equals("GUI Tests Started"))
                    {
                        string dropLocation = buildnrs[i].DropLocation;
                        string v = ConstCommon.DRIVE_MAP_LETTER + "\\" + dropLocation.Substring(55);

                        thisBuild = new BuildObject();
                        thisBuild.BuildNr = buildnrs[i].BuildNumber;
                        thisBuild.Quality = buildnrs[i].Quality;
                        thisBuild.FinishTime = buildnrs[i].FinishTime;
                        thisBuild.DripLoc = v;
                        allBuildObjects.Add(thisBuild);

                        // convert \\teamsystem.teamsystems.egemin.be\Team Systems Builds\CI... 
                        // to X:\CI...
                        //X:\CI\Epia 4\Epia.Development.Dev03-Net4.CI\Epia.Development.Dev03-Net4.CI_20110808.2
                        allBuildsInfo.Add(buildnrs[i].BuildNumber, v);
                        lstBoxDuildNumber.Items.Add(buildnrs[i].BuildNumber);
                    }
                    else
                    {
                        //logger.LogMessageToFile(buildnr + " has no valid quality:" + quality, sLogCount, sLogInterval);
                        continue;
                    }
                }
            }

            //IBuildDefinition[] buildDefinitions = buildServer.QueryBuildDefinitions(tpp.SelectedProjects[0].Name);
            IBuildDefinition[] buildDefinitions = buildServer.QueryBuildDefinitions(selectedProject);


            for (int j = 0; j < lstBoxDuildNumber.Items.Count; j++)
            {
                if (lstBoxDuildNumber.Items[j].ToString().Equals(selectedBuildNr))
                    lstBoxDuildNumber.SelectedIndex = j;
            }
        }

        private bool isBuildDefSelected(string def, List<string> SelecedDefs)
        {
            bool selected = false;
            IEnumerator EmpEnumerator = SelecedDefs.GetEnumerator();
            EmpEnumerator.Reset();
            while (EmpEnumerator.MoveNext())
            {
                if (def.Equals((string) EmpEnumerator.Current))
                {
                    selected = true;
                    break;
                }
            }
            return selected;
        }

        public string[] getBuildDefinition()
        {
            return buildDefs;
        }


        private void btnConn_Click(object sender, EventArgs e)
        {
            IBuildServer buildServer;
            tfsProjectCollection.EnsureAuthenticated();
            buildServer = (IBuildServer) tfsProjectCollection.GetService(typeof (IBuildServer));


            if (lstBoxDuildNumber.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select a build definition");
                btnOK.DialogResult = DialogResult.OK;
                // Call back to the parent passing the entire CustomerDialog Instance
                Close();
            }

            buildDefs = new string[lstBoxDuildNumber.SelectedItems.Count];
            for (int k = 0; k < lstBoxDuildNumber.SelectedItems.Count; k++)
            {
                //System.Windows.MessageBox.Show("selectedBuldDefinition=" + lstBoxDuildDefinition.SelectedItems[k].ToString());
                buildDefs[k] = lstBoxDuildNumber.SelectedItems[k].ToString();
                //IBuildDetail[] buildnrs = buildDefinitions[7].QueryBuilds();
                //IBuildDetail[] buildnrs = buildServer.QueryBuilds(selectedProject, lstBoxDuildDefinition.SelectedItems[k].ToString());
                //string bnrs = string.Empty;
                //for (int i = 0; i < buildnrs.Length; i++)
                //{
                //    bnrs = bnrs + "\n " + buildnrs[i].BuildNumber + "\t" + buildnrs[i].Quality + "\t" + buildnrs[i].DropLocation;
                //}
                //System.Windows.MessageBox.Show("Connection OK \n Build definition Name are " + tfsProjectCollection.Name + " \nwith build nrs:" + bnrs);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}