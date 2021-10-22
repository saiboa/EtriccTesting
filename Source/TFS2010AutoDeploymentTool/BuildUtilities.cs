using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;

using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
//using Microsoft.TeamFoundation.Build.Proxy;

using TestTools;

namespace TFS2010AutoDeploymentTool
{
    class BuildUtilities
    {
        // ———————————————————————————————————————————————————————————————————————————————————————————————————————————— 
        #region Methods of BuildUtilities (5)
        public static bool CheckTFSConnection(ref string msg )
        {
            bool TFSConnected = false;
            TfsTeamProjectCollection tfsProjectCollection;
            IBuildServer m_BuildSvc;
            try
            {
                //string sTFSServerUrl = Constants.sTFSServer;
                Uri serverUri = new Uri(Constants.sTFSServerUrl);
                System.Net.ICredentials tfsCredentials
                    = new System.Net.NetworkCredential(Constants.sTFSUsername, Constants.sTFSPassword, Constants.sTFSDomain);

                tfsProjectCollection
                    = new TfsTeamProjectCollection(serverUri, tfsCredentials);
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

        public static List<BuildObject> GetAllBuildObjects(List<string> buildDefinition, string selectedProject, string dateFilter)
        {
            List<BuildObject> allBuildslist = null;
            string sMsgDebug = Constants.sMsgDebug;
            Uri serverUri = new Uri(Constants.sTFSServerUrl);
            System.Net.ICredentials tfsCredentials
                = new System.Net.NetworkCredential(Constants.sTFSUsername, Constants.sTFSPassword, Constants.sTFSDomain);

            DateTime timeNow = DateTime.Now;
            DateTime timeFrom = DateTime.Now;
            if (dateFilter.StartsWith("<Any Time>"))
                timeFrom = DateTime.Today.Subtract(TimeSpan.FromDays(365));
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
                        IBuildDetailSpec buildDetailSpec = buildServer.CreateBuildDetailSpec(selectedProject, s);
                        //buildDetailSpec.MaxBuildsPerDefinition = 1; 
                        buildDetailSpec.QueryOrder = BuildQueryOrder.FinishTimeDescending;
                        buildDetailSpec.Status = BuildStatus.Succeeded; //Only get succeeded builds  
                        buildDetailSpec.MinFinishTime = timeFrom;
                        buildDetailSpec.QueryOptions = QueryOptions.None;
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
                        BuildObject thisBuild = new BuildObject();

                        for (int i = 0; i < buildnrs.Length; i++)
                        {
                            quality = buildnrs[i].Quality;
                            if (buildnrs[i].BuildNumber.ToLower().StartsWith("branch"))
                                continue;

                            if (quality == null
                                || quality.Equals("Rejected")
                                || quality.Equals("Released")
                                || quality.Equals("Under Investigation")
                                || quality.Equals("UAT Passed"))
                            {
                                //logger.LogMessageToFile(buildnr + " has no valid quality:" + quality, sLogCount, sLogInterval);
                                continue;
                            }
                            else
                            {
                                // convert \\teamsystem.teamsystems.egemin.be\Team Systems Builds\CI... 
                                // to X:\CI...
                                //X:\CI\Epia 4\Epia.Development.Dev03-Net4.CI\Epia.Development.Dev03-Net4.CI_20110808.2
                                string dropLocation = buildnrs[i].DropLocation;

                                int ipos = dropLocation.IndexOf("CI");
                                if (ipos == -1)
                                    ipos = dropLocation.IndexOf("Nightly");
                                if (ipos == -1)
                                    ipos = dropLocation.IndexOf("Production");
                                string xMap = dropLocation.Substring(0, ipos -1);

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
                                thisBuild.Quality = buildnrs[i].Quality;
                                thisBuild.FinishTime = buildnrs[i].FinishTime;
                                thisBuild.DripLoc = dropLocation;
                                //thisBuild.DripLoc = v;
                                thisBuild.xMapString = xMap;
                                thisBuild.RelativeDropLoc = dropLocation.Substring(ipos);
                                allBuildslist.Add(thisBuild);
                                // build info added
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
                    TestTools.MessageBoxEx.Show("Team Foundation services are not available from server\nWill try to reconnect the Server after 10 minutes",
                        kTime++ + "  This is automatic testing, please not touch the screen, time: " + timeNow.ToLongTimeString(), 10 * 60000);
                    System.Threading.Thread.Sleep(10 * 60000);
                    conn = false;
                }
                catch (Exception ex)
                {
                   TestTools.MessageBoxEx.Show( "TeamFoundation getService Exception:" + ex.Message + " ----- " + ex.StackTrace,
                        kTime++ + " This is automatic testing, please not touch the screen: exception time:" + timeNow.ToLongTimeString(), 10 * 60000);
                    System.Threading.Thread.Sleep(10 * 60000);
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
            if (BuildApp.Equals(Constants.EPIA4))
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="validBuildDir"> X:\Version\Epia 3\Epia\Epia - Version_20090414.2 </param>
        /// <returns>Version </returns>
        public static string getTestDefinition(string validBuildDir)
        {
            string def = string.Empty;
            //int ia = validBuildDir.IndexOf(":");
            //def = validBuildDir.Substring(ia+2);
            //def = def.Substring(0, def.IndexOf("\\"));
             if (validBuildDir.IndexOf("Version") > 0)
                def = "Version";
            else if (validBuildDir.IndexOf("Nightly") > 0)
                def = "Nightly";
            else if (validBuildDir.IndexOf("CI") > 0)
                def = "CI";
            else if (validBuildDir.IndexOf("Weekly") > 0)
                def = "Weekly";
            else if (validBuildDir.IndexOf("Production") > 0)
                 def = "Production";
            else def = "No def error";
            return def;
        }
        
        #endregion // —— Methods ••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

    }
}
