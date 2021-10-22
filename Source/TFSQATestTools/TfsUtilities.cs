using System;
using System.Collections.Generic;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using TestTools;

namespace TFSQATestTools
{
    public class TfsUtilities
    {      
        public static bool CheckTFSConnection(ref string msg )
        {
            bool TFSConnected = false;
            TfsTeamProjectCollection tfsProjectCollection;
            IBuildServer m_BuildSvc;
            try
            {
                Uri serverUri = new Uri(Tfs.ServerUrl);
                System.Net.ICredentials tfsCredentials = new System.Net.NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

                tfsProjectCollection = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                tfsProjectCollection.Connect(Microsoft.TeamFoundation.Framework.Common.ConnectOptions.IncludeServices);
                TfsConfigurationServer tfsConfigurationServer = tfsProjectCollection.ConfigurationServer;

                m_BuildSvc = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));
            }
            catch (TeamFoundationServerUnauthorizedException ex)
            {
                msg = ex.Message + System.Environment.NewLine + ex.StackTrace;
                TFSConnected = false;
                return false;
            }
            catch (Exception ex)
            {
                msg = ex.Message + System.Environment.NewLine + ex.StackTrace;
                TFSConnected = false;
                return false;
            }

            if (tfsProjectCollection == null || m_BuildSvc == null)
            {
                TFSConnected = false;
                msg = "tfsProjectCollection == null";
            }
            else
                TFSConnected = true;

            return TFSConnected;

        }

        public static bool CheckTFSConnection(ref string msg, ref TfsTeamProjectCollection tfsProjectCollection)
        {
            bool TFSConnected = false;
            try
            {
                Uri serverUri = new Uri(Tfs.ServerUrl);
                System.Net.ICredentials tfsCredentials
                    = new System.Net.NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

                tfsProjectCollection
                    = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                tfsProjectCollection.Connect(Microsoft.TeamFoundation.Framework.Common.ConnectOptions.IncludeServices);
                TfsConfigurationServer tfsConfigurationServer = tfsProjectCollection.ConfigurationServer;

                //m_BuildSvc = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));

            }
            catch (TeamFoundationServerUnauthorizedException ex)
            {
                msg = ex.Message + System.Environment.NewLine + ex.StackTrace;
                TFSConnected = false;
                return false;
            }
            catch (Exception ex)
            {
                msg = ex.Message + System.Environment.NewLine + ex.StackTrace;
                TFSConnected = false;
                return false;
            }

            if (tfsProjectCollection == null)
            {
                TFSConnected = false;
                msg = "tfsProjectCollection == null";
            }
            else
                TFSConnected = true;

            return TFSConnected;

        }

        static public bool GetTestProjectQA(TfsTeamProjectCollection tfsProjectCollection, string workingDirectory, ref string sErrorMessage)
        {
            bool result = true;
            VersionControlServer versionControlServer = (VersionControlServer)tfsProjectCollection.GetService(typeof(VersionControlServer));
            //=============
            string workspaceName = System.Environment.MachineName;
            //string workspaceName = "PCC7 - 201109029";
            string projectPath = @"$/Epia 3/Testing/Automatic/AutomaticTests/TestData/QA"; // the container Project (like a tabel in sql/ or like a folder) containing the projects sources in a collection (like a database in sql/ or also like a folder) from TFS          
            //string workingDirectory = @"C:\EtriccTests\QA";  // local folder where to save projects sources    
            //string workingDirectory = @"C:\EtriccTests\QA";  // local folder where to save projects sources      
            //TeamFoundationServer tfs = new TeamFoundationServer("http://test-server:8080/tfs/CollectionName", System.Net.CredentialCache.DefaultCredentials); // tfs server url including the  Collection Name --  CollectionName as the existing name of the collection from the tfs server          
            //tfs.EnsureAuthenticated();

            //VersionControlServer sourceControl = (VersionControlServer)tfs.GetService(typeof(VersionControlServer));
            Workspace[] workspaces = versionControlServer.QueryWorkspaces(workspaceName, versionControlServer.AuthorizedUser, Workstation.Current.Name);
            if (workspaces.Length > 0)
            {
                versionControlServer.DeleteWorkspace(workspaceName, versionControlServer.AuthorizedUser);
            }

            Workspace workspace = versionControlServer.CreateWorkspace(workspaceName, versionControlServer.AuthorizedUser, "Temporary Workspace");
            try
            {
                workspace.Map(projectPath, workingDirectory);
                GetRequest request = new GetRequest(new ItemSpec(projectPath, RecursionType.Full), VersionSpec.Latest);
                GetStatus status = workspace.Get(request, GetOptions.GetAll | GetOptions.Overwrite); // this line doesn't do anything - no failures or errors         
            }
            catch (Exception ex)
            {
                sErrorMessage = ex.Message + "-----" + ex.StackTrace;
                result = false; ;
            }
            finally
            {
                if (workspace != null)
                {
                    workspace.Delete();
                    //System.Windows.Forms.MessageBox.Show("The Projects have been brought into the Folder  " + workingDirectory);
                }
            }

            return result;

        }

        public static string GetTeamProjectFromTestApp(string testApp)
        {
            string selectedProject = string.Empty;
            if (testApp.Equals(TestApp.EPIA4) || testApp.Equals(TestApp.EPIANET45))
                selectedProject = "Epia 4";
            else
                selectedProject = "Etricc 5";

            return selectedProject;
        }

        public static string GetTestDefNameFromTestApp(string testApp)
        {
            string testDefName = string.Empty;
            if (testApp.Equals(TestApp.EPIANET45))
                testDefName = TestDefName.EPIANET45;
            else if (testApp.Equals(TestApp.ETRICCNET45))
                testDefName = TestDefName.ETRICCNET45;
            else if (testApp.Equals(TestApp.EPIA4))
                testDefName = TestDefName.EPIA4;
            else if (testApp.Equals(TestApp.ETRICCUI))
                testDefName = TestDefName.ETRICCUI;
            else if (testApp.Equals(TestApp.ETRICCSTATISTICS))
                testDefName = TestDefName.ETRICCSTATISTICS;

            return testDefName;
        }

        public static string GetTeamProjectFromBuildDefinition(string buildDef)
        {
            string sTeamProject = string.Empty;
            if (buildDef.IndexOf("Epia") >=0 )
                sTeamProject = "Epia 4";
            else
                sTeamProject = "Etricc 5";

            return sTeamProject;
        }

        public static string GetTestAppFromBuildDefinition(string buildDef)
        {
            string sTestApp = string.Empty;
            if (buildDef.IndexOf("Epia") >= 0 && buildDef.IndexOf("Dev045") > 0)
                sTestApp = TestApp.EPIANET45;
            else if (buildDef.IndexOf("Epia") >= 0)
                sTestApp = TestApp.EPIA4;
            else if (buildDef.IndexOf("Stat Prog") >= 0)
                sTestApp = TestApp.ETRICCSTATISTICS;
            else if (buildDef.IndexOf("Etricc") >= 0)
                sTestApp = TestApp.ETRICCUI;
            else
                System.Windows.Forms.MessageBox.Show(buildDef + "  has no test App , please check again");
               
            return sTestApp;
        }

        public static List<BuildObject> GetAllBuildObjects(List<string> buildDefinition, /*string selectedProject,*/ string dateFilter, Tester test, Logger logger)
        {
            List<BuildObject> allBuildslist = null;
            string sMsgDebug = Constants.sMsgDebug;
            Uri serverUri = new Uri(Tfs.ServerUrl);
            System.Net.ICredentials tfsCredentials
                = new System.Net.NetworkCredential(Tfs.UserName, Tfs.Password, Tfs.Domain);

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

            TfsTeamProjectCollection tfsProjectCollection  = null;
            IBuildServer buildServer = null; ; 
            bool conn = false;
            int kTime = 0;
            while (conn == false)
            {
                try
                {
                    conn = true;
                    tfsProjectCollection = new TfsTeamProjectCollection(serverUri, tfsCredentials);
                    tfsProjectCollection.EnsureAuthenticated();
                    buildServer = (IBuildServer)tfsProjectCollection.GetService(typeof(IBuildServer));

                    // Get All Build from server
                    #region
                    allBuildslist = new List<BuildObject>();
                    foreach (string s in buildDefinition)
                    {
                        string selectedProject = TfsUtilities.GetTeamProjectFromBuildDefinition(s);
                        IBuildDetailSpec buildDetailSpec = buildServer.CreateBuildDetailSpec(selectedProject, s);
                        //buildDetailSpec.MaxBuildsPerDefinition = 1; 
                        buildDetailSpec.QueryOrder = BuildQueryOrder.FinishTimeDescending;
                        buildDetailSpec.Status = BuildStatus.Succeeded; //Only get succeeded builds  
                        buildDetailSpec.MinFinishTime = timeFrom;
                        buildDetailSpec.QueryOptions = QueryOptions.None;

                        IBuildQueryResult results = buildServer.QueryBuilds(buildDetailSpec);
                        //System.Windows.MessageBox.Show("buildDefinition:" + s + "  has Builds.Length:" + results.Builds.Length);
                       
                        test.Log("buildDefinition:" + s + "  has Builds.Length:" + results.Builds.Length);
                        //if (results.Failures.Length == 0 ) 
                        //{ 
                        //IBuildDetail buildDetail = results.Builds[0]; 
                        //Console.WriteLine("Build: " + buildDetail.BuildNumber); 
                        //Console.WriteLine("Account requesting build “ + 
                        //“(build service user for triggered builds): " + buildDetail.RequestedBy); 
                        //   Console.WriteLine("Build triggered by: " + buildDetail.RequestedFor); 
                        //}

                        IBuildDetail[] buildnrs = null;
                        if (results.Builds != null && results.Builds.Length > 0)
                        {
                            buildnrs = results.Builds;
                            //IBuildDetail[] buildnrs = buildServer.QueryBuilds(selectedProject, s);
                        }
                        else
                        {
                            test.Log("buildDefinition:" + s + "  has Builds.Length:" + results.Builds.Length + " and searching next buildDefinition: ...");
                            logger.LogMessageToFile("", 0, 0);
                            continue;
                        }

                        string bnrs = string.Empty;
                        string quality = string.Empty;
                        BuildObject thisBuild = new BuildObject();

                        for (int i = 0; i < buildnrs.Length; i++)
                        {
                            quality = buildnrs[i].Quality;
                            if (buildnrs[i].BuildNumber.ToLower().StartsWith("branch"))
                                continue;

                            //test.Log("buildnrs[i]:" + buildnrs[i].BuildNumber.ToString() + "   ---  with quality :" + quality);
                            if (quality.Equals("Unexamined")
                                || quality.Equals("GUI Tests Started"))
                            {
                                // convert \\teamsystem.teamsystems.egemin.be\Team Systems Builds\CI... 
                                // to X:\CI...
                                //X:\CI\Epia 4\Epia.Development.Dev03-Net4.CI\Epia.Development.Dev03-Net4.CI_20110808.2
                                string dropLocation = buildnrs[i].DropLocation;
                                //test.Log("dropLocation:" + dropLocation);
                                // 2015Match13 updated : tfs build on azure CI,Nightly,Production are all in lower case
                                dropLocation = dropLocation.ToLower();
                                int ipos = dropLocation.IndexOf("ci");
                                if (ipos == -1)
                                    ipos = dropLocation.IndexOf("nightly");
                                if (ipos == -1)
                                    ipos = dropLocation.IndexOf("production");
                                /*int ipos = dropLocation.IndexOf("CI");
                                if (ipos == -1)
                                    ipos = dropLocation.IndexOf("Nightly");
                                if (ipos == -1)
                                    ipos = dropLocation.IndexOf("Production");
                                */
                                if (ipos == -1)
                                {
                                    // old structure folder before 2011 May 10th.  do not have CI, Nightly or Production folder. but has folder Version ...
                                    // \\TeamSystem.TeamSystems.Egemin.Be\Team Systems Builds\Version\Etricc\Etricc 5 -Statistics Programs\Etr...
                                    // not test anymore 
                                    //System.Windows.Forms.MessageBox.Show("dropLocation:" + dropLocation, "dropLocationRoot:" + buildnrs[i].DropLocationRoot);
                                    continue;
                                }

                                //test.Log("ipos:" + ipos);
                                string xMap = dropLocation.Substring(0, ipos -1);
                                //test.Log("xMap:" + xMap);
                                //string driveMap = System.Configuration.ConfigurationManager.AppSettings.Get("NetworkDriveFolderMap");
                                //string v = ConstCommon.DRIVE_MAP_LETTER + "\\" + dropLocation.Substring(driveMap.Length);
                                //string v = ConstCommon.DRIVE_MAP_LETTER + "\\" + dropLocation.Substring(55);
                                if (sMsgDebug.StartsWith("true"))
                                {
                                    System.Windows.Forms.MessageBox.Show("dropLocation:" + dropLocation, "dropLocationRoot:" + buildnrs[i].DropLocationRoot);
                                    System.Windows.Forms.MessageBox.Show("xMap:" + xMap, "RelativeDropLoc:" + dropLocation.Substring(ipos));
                                    //System.Windows.Forms.MessageBox.Show("new string: " + v, "GetAllBuildObjects");
                                    System.Windows.MessageBox.Show(" buildnrs.Length:" + buildnrs.Length);
                                }
                                // add to build list
                                thisBuild = new BuildObject();
                                thisBuild.BuildNr = buildnrs[i].BuildNumber;
                                thisBuild.BuildDef = s;
                                thisBuild.Quality = buildnrs[i].Quality;
                                thisBuild.FinishTime = buildnrs[i].FinishTime;
                                thisBuild.DripLoc = dropLocation;
                                //thisBuild.DripLoc = v;
                                thisBuild.xMapString = xMap;
                                
                                
                                thisBuild.RelativeDropLoc = dropLocation.Substring(ipos);

                                //System.Windows.MessageBox.Show("dropLocation:" + dropLocation);
                                //System.Windows.MessageBox.Show("xMapString:" + xMap);
                                //System.Windows.MessageBox.Show("RelativeDropLoc:" + thisBuild.RelativeDropLoc);

                                allBuildslist.Add(thisBuild);
                                // build info added
                            }
                            else
                            {
                                //logger.LogMessageToFile(buildnr + " has no valid quality:" + quality, sLogCount, sLogInterval);
                                continue;
                            }
                        }
                    }
                    #endregion

                    allBuildslist.Sort(delegate(BuildObject p1, BuildObject p2)
                    {
                        return p2.FinishTime.CompareTo(p1.FinishTime);
                    });

                    return allBuildslist;
                }
                catch (Microsoft.TeamFoundation.TeamFoundationServiceUnavailableException ex)
                {
                    test.Log("Team Foundation services are not available from server\nWill try to reconnect the Server ...\n" + ex.Message + " --- " +ex.StackTrace);
                    TestTools.MessageBoxEx.Show("Team Foundation services are not available from server\nWill try to reconnect the Server ...\n" + ex.Message,
                        kTime++ + "  This is automatic testing, please not touch the screen, time: " + timeNow.ToLongTimeString(), (uint)Tfs.ReconnectDelay );
                    System.Threading.Thread.Sleep( Tfs.ReconnectDelay );
                    conn = false;
                }
                catch (Exception ex)
                {
                   test.Log("TeamFoundation getService Exception:" + ex.Message + " --- " + ex.StackTrace);
                   TestTools.MessageBoxEx.Show( "TeamFoundation getService Exception:" + ex.Message + " ----- " + ex.StackTrace,
                        kTime++ + " This is automatic testing, please not touch the screen: exception time:" + timeNow.ToLongTimeString(), (uint)Tfs.ReconnectDelay );
                   System.Threading.Thread.Sleep( Tfs.ReconnectDelay );
                    conn = false;
                }
            }

            return allBuildslist;
           
            // sort BuildObjectList will be used later
            //dataGridView1.DataSource = allBuildslist;  // http://www.dotnetperls.com/datagridview
            /*
            allBuildslist.Sort(delegate(BuildObject p1, BuildObject p2) 
            { 
                return p2.BuildNr.CompareTo(p1.BuildNr); 
            });

            string xs = string.Empty;
            IEnumerator EmpEnumerator = allBuildslist.GetEnumerator(); //Getting the Enumerator
            EmpEnumerator.Reset(); //Position at the Beginning
            while (EmpEnumerator.MoveNext()) //Till not finished do print
            {
                BuildObject b = (BuildObject)EmpEnumerator.Current;
                xs = xs + b.BuildNr + " -  " + b.Quality+ " -  " + b.FinishTime.ToShortTimeString() + "\n";
            }

            System.Windows.Forms.MessageBox.Show(xs);


            xs = xs  + "\n";
            allBuildslist.Sort(delegate(BuildObject p1, BuildObject p2)
            {
                return p1.Quality.CompareTo(p2.Quality);
            });
            EmpEnumerator = allBuildslist.GetEnumerator(); //Getting the Enumerator
            EmpEnumerator.Reset(); //Position at the Beginning
            while (EmpEnumerator.MoveNext()) //Till not finished do print
            {
                BuildObject b = (BuildObject)EmpEnumerator.Current;
                xs = xs + b.BuildNr + " -  " + b.Quality + " -  " + b.FinishTime.ToShortTimeString() + "\n";
            }

            System.Windows.Forms.MessageBox.Show(xs);

            xs = xs + "\n";
            */
            // sort buildnr by finished time descendant
        }

        public static string GetProjectName(string BuildApp)
        {
            string projName = string.Empty;
            if (BuildApp.Equals(TestApp.EPIANET45))
            {
                projName = ConstCommon.EPIA_4;
            }
            else if (BuildApp.Equals(TestApp.EPIA4))
            {
                projName = ConstCommon.EPIA_4;
            }
            else if (BuildApp.Equals(Constants.ETRICCUI))
            {
                projName = ConstCommon.ETRICC_5;
            }
            else if (BuildApp.Equals(Constants.ETRICC5))
            {
                projName = ConstCommon.ETRICC_5;
            }
            else if (BuildApp.Equals(Constants.ETRICCSTATISTICS))
            {
                projName = ConstCommon.ETRICC_5;
            }
            else if (BuildApp.Equals(Constants.EWMS))
            {
                projName = ConstCommon.EPIA_3;
            }
            else if (BuildApp.Equals(ConstCommon.KIMBERLY_CLARK))
            {
                projName = ConstCommon.EWCS_PROJECTS;
            }
            else
                projName = "No Project name for this Application";

            return projName;
        }

        public static Uri GetBuildUriFromBuildNumber( IBuildServer buildSvc, string projectName, string buildNumber )
        {
            Uri buildUri = null;
            // Get a list of all builds for the specified project
            ///IBuildDetail[] builds = buildSvc.QueryBuilds(projectName);
            // Locate the desired build URI based on build number
            /*foreach (IBuildDetail build in builds)
            {
                if (build.BuildNumber.Equals(buildNumber, StringComparison.InvariantCultureIgnoreCase))
                {
                    buildUri = build.Uri;
                    break;
                }
            }*/

            IBuildDetailSpec spec = buildSvc.CreateBuildDetailSpec( projectName );
            spec.BuildNumber = buildNumber;
            spec.QueryOptions = QueryOptions.None;
            IBuildQueryResult res = buildSvc.QueryBuilds( spec );
            IBuildDetail buildDetail = res.Builds[0];
            buildUri = buildDetail.Uri;
            return buildUri;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="validBuildDir"> X:\Version\Epia 3\Epia\Epia - Version_20090414.2 </param>
        /// <returns>Version </returns>
        public static string GetTestDefinition(string validBuildDir)
        {
            string def = string.Empty;
            validBuildDir = validBuildDir.ToLower();
            //int ia = validBuildDir.IndexOf(":");
            //def = validBuildDir.Substring(ia+2);
            //def = def.Substring(0, def.IndexOf("\\"));
             if (validBuildDir.IndexOf("version") > 0)
                def = "Version";
            else if (validBuildDir.IndexOf("nightly") > 0)
                def = "Nightly";
            else if (validBuildDir.IndexOf("ci") > 0)
                def = "CI";
            else if (validBuildDir.IndexOf("weekly") > 0)
                def = "Weekly";
            else if (validBuildDir.IndexOf("production") > 0)
                 def = "Production";
            else def = "No def error";
            return def;
        }

        public static string UpdateBuildQualityStatus( Logger logger, Uri uri, string projectName, string newQuality, IBuildServer buildSvc, string demonstration )
        {
            string msgScreenLog = string.Empty;
            IBuildDetail bd = buildSvc.GetMinimalBuildDetails( uri );
            string quality = bd.Quality;
            string buildNr = bd.BuildNumber;
            logger.LogMessageToFile( buildNr + " UpdateBuildQualityStatus projectName:" + projectName + " <old quality> " + quality, 0, 0 );
            logger.LogMessageToFile( buildNr + " quality will be updated from : " + quality + " <to> " + newQuality, 0, 0 );


            if ( demonstration.ToLower().StartsWith( "true" ) )
            {
                msgScreenLog = buildNr + " !!! this is a demonstration ===> no update needed" + quality;
                logger.LogMessageToFile( buildNr + " !!! this is a demonstration ===> no update needed:<old quality>:" + quality, 0, 0 );
                return msgScreenLog;
            }

            try
            {
                if ( quality.Equals( "GUI Tests Failed" ) )
                {
                    msgScreenLog = buildNr + " has failed quality, no update needed :" + quality;
                    logger.LogMessageToFile( buildNr + " has failed quality, no update needed :<old quality>:" + quality, 0, 0 );
                }
                else
                {
                    IBuildDetailSpec spec = buildSvc.CreateBuildDetailSpec( projectName );
                    spec.BuildNumber = buildNr;
                    spec.QueryOptions = QueryOptions.None;
                    IBuildQueryResult res = buildSvc.QueryBuilds( spec );
                    IBuildDetail buildDetail = res.Builds[0];
                    buildDetail.Quality = newQuality;
                    buildDetail.Save();

                    msgScreenLog = buildNr + " has old quality:" + quality;
                    logger.LogMessageToFile( buildNr + " has old quality:" + quality, 0, 0 );
                }
            }
            catch ( Exception ex )
            {
                msgScreenLog = "ERROR - UpdateBuildQualityStatus- " + newQuality + " --- " + ex.Message + " --- " + ex.StackTrace;
            }

            return msgScreenLog;
        }

        public static string UpdateBuildQualityStatusEvenHasFailedStatus( Logger logger, Uri uri, string projectName, string newQuality, IBuildServer buildSvc, string demonstration )
        {
            string msgScreenLog = string.Empty;
            IBuildDetail bd = buildSvc.GetMinimalBuildDetails( uri );
            string quality = bd.Quality;
            string buildNr = bd.BuildNumber;
            logger.LogMessageToFile( buildNr + " UpdateBuildQualityStatus projectName:" + projectName + " <old quality> " + quality, 0, 0 );
            logger.LogMessageToFile( buildNr + " quality will be updated from :" + quality + " <to> " + newQuality, 0, 0 );


            if ( demonstration.ToLower().StartsWith( "true" ) )
            {
                msgScreenLog = buildNr + " !!! this is a demonstration ===> no update needed" + quality;
                logger.LogMessageToFile( buildNr + " !!! this is a demonstration ===> no update needed:<old quality>:" + quality, 0, 0 );
                return msgScreenLog;
            }

            try
            {
                //if (quality.Equals("GUI Tests Failed"))
                //{
                //    msgScreenLog = buildNr + " has failed quality, no update needed :" + quality;
                //    logger.LogMessageToFile(buildNr + " has failed quality, no update needed :<old quality>:" + quality, 0, 0);
                //}
                //else
                //{
                IBuildDetailSpec spec = buildSvc.CreateBuildDetailSpec( projectName );
                spec.BuildNumber = buildNr;
                spec.QueryOptions = QueryOptions.None;
                IBuildQueryResult res = buildSvc.QueryBuilds( spec );
                IBuildDetail buildDetail = res.Builds[0];
                buildDetail.Quality = newQuality;
                buildDetail.Save();

                msgScreenLog = buildNr + " has old quality:" + quality;
                logger.LogMessageToFile( buildNr + " has old quality:" + quality, 0, 0 );
                //}
            }
            catch ( Exception ex )
            {
                msgScreenLog = "ERROR - UpdateBuildQualityStatus- " + newQuality + " --- " + ex.Message + " --- " + ex.StackTrace;
            }

            return msgScreenLog;
        }
    }
}
