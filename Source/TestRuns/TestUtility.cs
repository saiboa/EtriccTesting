using System;
using System.Collections;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Text;
using Egemin.EPIA;
using Egemin.EPIA.Core.Definitions;
using Egemin.EPIA.WCS.Core;
using Egemin.EPIA.WCS.Resources;
using Egemin.EPIA.WCS.Scheduling;
using Egemin.EPIA.WCS.Transportation;
using Microsoft.Office.Interop.Excel;
using TestTools;
using Constants = Egemin.EPIA.Constants;

namespace TestRuns
{
    public class TestUtility
    {
        /// <summary>
        ///     Check Agv has specific job type
        /// </summary>
        /// <param name="testAgv">testAgv</param>
        /// <param name="jobType">job type</param>
        /// <returns>true found</returns>
        public static bool CheckAgvHasJobWithType(Agv testAgv, string jobType)
        {
            bool found = false;
            for (int i = 0; i < testAgv.Jobs.GetArray().Length; i++)
            {
                var job = (Job) testAgv.Jobs.GetArray()[i];
                if (job.Type.ToString().ToLower().StartsWith(jobType.ToLower()))
                {
                    found = true;
                    break;
                }
            }
            return found;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sTSName"></param>
        /// <param name="arg"></param>
        /// <param name="plusDays"></param>
        /// <param name="plusMin"></param>
        /// <returns></returns>
        public static WeekPlan CreateWeekPlan(string sTSName, string arg, int plusDays, int plusMin)
        {
            var wp = new WeekPlan();
            wp.ID = sTSName;
            //wp.Day = System.DateTime.Now.DayOfWeek;
            wp.Day = DateTime.Today.AddDays(plusDays).DayOfWeek;
            wp.StartHour = DateTime.Now.Hour;
            wp.StartMinute = DateTime.Now.Minute + plusMin;
            wp.Duration = 1; // 1 min
            wp.Arguments.Add(arg);
            wp.Enable();
            return wp;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="m_Project"></param>
        /// <param name="schedulingType"></param>
        /// <param name="ruleType"></param>
        /// <param name="args"></param>
        public static void SetSchedulesRules(ref Project m_Project, string schedulingType, string ruleType, string args)
        {
            string xx = m_Project.Schedules[schedulingType].Type.ToString();
            xx = xx + " voor -->rule type: " + m_Project.Schedules[schedulingType].Rules[0].Type.ToString();

            m_Project.Schedules[schedulingType].Rules.Deactivate();
            m_Project.Schedules[schedulingType].Rules.Clear();

            Rule rule;
            rule = new Rule();
            if (ruleType.Equals("OLDEST"))
            {
                rule.Type = Rule.TYPE.OLDEST;
                rule.Arguments = new Tokeniser(true, Constants.ANY + " 0 0");
            }
            else if (ruleType.Equals("CLOSEST"))
            {
                rule.Type = Rule.TYPE.CLOSEST;
                rule.Arguments = new Tokeniser(true, Constants.ANY + " 0 0");
            }
            else if (ruleType.Equals("CLOSEST_HIGHEST"))
            {
                rule.Type = Rule.TYPE.CLOSEST_HIGHEST;
                rule.Arguments = new Tokeniser(true, Constants.ANY + " 0 0");
            }
            else if (ruleType.Equals("DELAY"))
            {
                rule.Type = Rule.TYPE.DELAY;
                rule.Arguments = new Tokeniser(true, args);
            }
            else if (ruleType.Equals("DIVERT"))
            {
                rule.Type = Rule.TYPE.DIVERT;
                rule.Arguments = new Tokeniser(true, args);
            }
            else if (ruleType.Equals("HIGHEST"))
            {
                rule.Type = Rule.TYPE.HIGHEST;
                rule.Arguments = new Tokeniser(true, Constants.ANY + " 0 0");
            }
            else if (ruleType.Equals("QUEUE"))
            {
                rule.Type = Rule.TYPE.QUEUE;
                rule.Arguments = new Tokeniser(true, args);
            }
            else if (ruleType.Equals("VIA"))
                rule.Type = Rule.TYPE.VIA;

            m_Project.Schedules[schedulingType].Rules.Add(rule);
            m_Project.Schedules[schedulingType].Rules[0].Simulation = true;
            m_Project.Schedules[schedulingType].Rules.Activate();

            string yy = m_Project.Schedules[schedulingType].Type.ToString();
            yy = yy + " na -->rule type: " + m_Project.Schedules[schedulingType].Rules[0].Type.ToString();

            RemoteLogMessage(xx + Environment.NewLine + yy, true, m_Project);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="m_Project"></param>
        /// <param name="schedulingType"></param>
        /// <param name="ruleType"></param>
        /// <param name="args"></param>
        public static void AddSchedulesRules(ref Project m_Project, string schedulingType, string ruleType, string args)
        {
            m_Project.Schedules[schedulingType].Rules.Deactivate();

            Rule rule;
            rule = new Rule();
            if (ruleType.StartsWith("OLDEST"))
            {
                rule.Type = Rule.TYPE.OLDEST;
                rule.Arguments = new Tokeniser(true, Constants.ANY + " 0 0");
            }
            else if (ruleType.StartsWith("CLOSEST_HIGHEST"))
            {
                rule.Type = Rule.TYPE.CLOSEST_HIGHEST;
                rule.Arguments = new Tokeniser(true, Constants.ANY + " 0 0");
            }
            else if (ruleType.StartsWith("DELAY"))
            {
                rule.Type = Rule.TYPE.DELAY;
                rule.Arguments = new Tokeniser(true, args);
            }
            else if (ruleType.StartsWith("DIVERT"))
            {
                rule.Type = Rule.TYPE.DIVERT;
                rule.Arguments = new Tokeniser(true, args);
            }
            else if (ruleType.StartsWith("QUEUE"))
            {
                rule.Type = Rule.TYPE.DIVERT;
                rule.Arguments = new Tokeniser(true, args);
            }
            else if (ruleType.StartsWith("VIA"))
                rule.Type = Rule.TYPE.VIA;


            m_Project.Schedules[schedulingType].Rules.Add(rule);
            m_Project.Schedules[schedulingType].Rules.Activate();

            int iRules = m_Project.Schedules[schedulingType].Rules.GetArray().Length;

            string yy = m_Project.Schedules[schedulingType].Type.ToString();

            for (int i = 0; i < iRules; i++)
            {
                yy = yy + Environment.NewLine
                     + " -->rule type: " + m_Project.Schedules[schedulingType].Rules[i].Type.ToString()
                     + " args: " + m_Project.Schedules[schedulingType].Rules[i].Arguments.ToString();
            }

            RemoteLogMessage(yy, true, m_Project);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="m_Project"></param>
        /// <param name="LocID"></param>
        /// <param name="priority"></param>
        /// <param name="logger"></param>
        public static void SetLocationPriority(ref Project m_Project, string LocID, int priority, ref Logger logger)
        {
            string xx = LocID + " voor -->priority: " + m_Project.Locations[LocID].Priority.ToString();
            RemoteLogMessage(xx, true, m_Project);
            logger.LogMessageToFile(xx);

            m_Project.Locations[LocID].Deactivate();
            m_Project.Locations[LocID].Priority = priority;

            m_Project.Locations[LocID].Activate();

            string yy = LocID + " na -->priority: " + m_Project.Locations[LocID].Priority.ToString();
            RemoteLogMessage(yy, true, m_Project);
            logger.LogMessageToFile(yy);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="m_Project"></param>
        /// <param name="LocID"></param>
        /// <param name="priority"></param>
        /// <param name="logger"></param>
        public static void SetGroupPriority(ref Project m_Project, string LocID, int priority, ref Logger logger)
        {
            string xx = LocID + " voor -->priority: " + m_Project.Groups[LocID].Priority.ToString();
            RemoteLogMessage(xx, true, m_Project);
            logger.LogMessageToFile(xx);

            //m_Project.Groups[LocID].Deactivate();
            m_Project.Groups[LocID].Priority = priority;
            //m_Project.Groups[LocID].Activate();

            string yy = LocID + " na -->priority: " + m_Project.Groups[LocID].Priority.ToString();
            RemoteLogMessage(yy, true, m_Project);
            logger.LogMessageToFile(yy);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="jobType"></param>
        /// <param name="locationID"></param>
        /// <param name="comments"></param>
        /// <returns></returns>
        public static Job CreateTestJob(string jobType, string locationID, string comments)
        {
            var job = new Job();
            job.Clear();
            job.Comments = comments;
            if (jobType == "PARK") job.Type = Job.TYPE.PARK;
            if (jobType == "PICK") job.Type = Job.TYPE.PICK;
            if (jobType == "DROP") job.Type = Job.TYPE.DROP;
            if (jobType == "WAIT") job.Type = Job.TYPE.WAIT;
            if (jobType == "BATT") job.Type = Job.TYPE.BATT;
            job.CarrierID = "1";
            job.LocationID = locationID;
            job.Priority = 5;
            job.OriginatorID = "Epia Auto Test runs";
            job.Reason = Job.REASON.EXTERNAL_REQUEST;
            return job;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="jobID"></param>
        /// <param name="jobType"></param>
        /// <param name="locationID"></param>
        /// <param name="comments"></param>
        /// <returns></returns>
        public static Job CreateTestJob(string jobID, string jobType, string locationID, string comments, int projectID)
        {
            var job = new Job();
            job.Clear();
            job.Comments = comments;
            if (jobType == "PARK") job.Type = Job.TYPE.PARK;
            if (jobType == "PICK") job.Type = Job.TYPE.PICK;
            if (jobType == "DROP") job.Type = Job.TYPE.DROP;
            if (jobType == "WAIT") job.Type = Job.TYPE.WAIT;
            if (jobType == "BATT") job.Type = Job.TYPE.BATT;
            job.ID = jobID;
            //if (projectID == TestConstants.PROJECT_EUROBALTIC)
            job.CarrierID = "1";
            //else
            //    job.CarrierID = "CARRIER1";

            job.LocationID = locationID;
            job.Priority = 5;
            job.OriginatorID = "Epia Auto Test runs";
            job.Reason = Job.REASON.EXTERNAL_REQUEST;
            return job;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="jobType"></param>
        /// <param name="locationID"></param>
        /// <param name="viaStation"></param>
        /// <param name="comments"></param>
        /// <returns></returns>
        public static Job CreateTestJob(string jobID, string jobType, string locationID, string viaStation,
                                        string comments)
        {
            var job = new Job();
            job.Clear();
            job.Comments = comments;
            if (jobType == "PARK") job.Type = Job.TYPE.PARK;
            if (jobType == "PICK") job.Type = Job.TYPE.PICK;
            if (jobType == "DROP") job.Type = Job.TYPE.DROP;
            if (jobType == "WAIT") job.Type = Job.TYPE.WAIT;
            if (jobType == "BATT") job.Type = Job.TYPE.BATT;
            job.ID = jobID;
            job.CarrierID = "1";
            job.LocationID = locationID;
            job.Priority = 5;
            job.ViaLSIDs.Add(viaStation);
            job.OriginatorID = "Epia Auto Test runs";
            job.Reason = Job.REASON.EXTERNAL_REQUEST;
            return job;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="transType"></param>
        /// <param name="agv"></param>
        /// <param name="sourceID"></param>
        /// <param name="destID"></param>
        /// <param name="TestScenarioID"></param>
        /// <param name="m_Project"></param>
        /// <param name="m_Transport"></param>
        /// <returns></returns>
        public static Transport CreateTestTransport(string transType, Agv agv, string sourceID, string destID,
                                                    string TestScenarioID, ref Project m_Project,
                                                    ref Transport m_Transport)
        {
            Transport transport;
            switch (transType)
            {
                case "PICK":
                    if (agv == null)
                        transport = new Transport(Transport.COMMAND.PICK, null, null, sourceID, destID);
                    else
                        transport = new Transport(Transport.COMMAND.PICK, null, agv.ID.ToString(), null, sourceID, null);
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    //transport = new Transport( Transport.COMMAND.PICK, null, null, agv.ID.ToString(), destID, null );
                    break;
                case "DROP":
                    if (agv == null)
                        transport = new Transport(Transport.COMMAND.DROP, null, null, sourceID, destID);
                    else
                        transport = new Transport(Transport.COMMAND.DROP, null, agv.ID.ToString(), null, null, destID);
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    break;
                case "MOVE":
                    if (agv == null)
                        transport = new Transport(Transport.COMMAND.MOVE, null, null, sourceID, destID);
                    else
                        transport = new Transport(Transport.COMMAND.MOVE, null, agv.ID.ToString(), null, sourceID,
                                                  destID);
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    break;
                case "WAIT":
                    if (agv == null)
                        transport = new Transport(Transport.COMMAND.WAIT, null, null, sourceID, destID);
                    else
                        transport = new Transport(Transport.COMMAND.WAIT, null, agv.ID.ToString(), null, null, destID);
                    m_Transport = m_Project.TransportManager.NewTransport(transport);
                    break;
                default:
                    break;
            }
            return m_Transport;
        }

        /// <summary>
        /// WaitUntilTransportState
        /// </summary>
        public static void WaitUntilTransportState(ref int checkRun, Transport transport, DateTime startTime,
                                                   int timeDuration, Transport.STATE state, ref string msg)
        {
            if (transport.State == state)
            {
                msg = "OK " + transport.ID + " is " + state.ToString();
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    msg = "after " + timeDuration + " min " + transport.ID + " state is : " + transport.State.ToString();
                    ;
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        // end WaitUntilTransportState

        public static void WaitUntilJobState(ref int checkRun, Job job, Job.STATE state, DateTime startTime,
                                             int timeDuration, ref string msg)
        {
            if (job.State == state)
            {
                msg = "OK " + job.ID + " is " + state.ToString();
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    msg = "after " + timeDuration + " min " + job.ID + " state is : " + job.State.ToString();
                    ;
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        public static void WaitUntilAgvAllJobsState(ref int checkRun, Agv agv, Job.STATE state, DateTime startTime,
                                                    int timeDuration, ref string msg)
        {
            bool allState = true;
            Jobs jobs = agv.Jobs;
            string allJobsID = "(";
            for (int i = 0; i < jobs.Count; i++)
            {
                allJobsID = allJobsID + jobs[i].ID + "  ;  ";
                if (!(jobs[i].State == Job.STATE.FINISHED))
                {
                    allState = false;
                    break;
                }
            }
            allJobsID = allJobsID + ")";

            if (allState)
            {
                msg = "OK All jobs" + allJobsID + " state are:" + state.ToString();
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    msg = "after " + timeDuration + " min " + allJobsID + " state are not : " + state.ToString();
                    ;
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        public static void WaitUntilAgvLockLSID(ref int checkRun, Agv agv, string lockedLSID, DateTime startTime,
                                                int timeDuration, ref string msg)
        {
            Object[] objLockedIDs = agv.LockedNodeIDs.GetArray();
            bool locked = false;
            string ids = "lockedids are ";
            for (int i = 0; i < objLockedIDs.Length; i++)
            {
                ids = ids + objLockedIDs[i] + Environment.NewLine;
                if (objLockedIDs[i].ToString().Equals(lockedLSID))
                {
                    locked = true;
                    break;
                }
            }

            if (locked)
            {
                msg = "OK " + ids;
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    msg = "after " + timeDuration + " min " + " no lockedid found : " + ids;
                    checkRun = TestConstants.CHECK_END;
                }
                //else
                //    msg = "PPPPPOK :::::" + ids;
            }
        }

        public static void CheckAgvNotMoving(ref int checkRun, Agv agv, string prevLSID, DateTime startTime,
                                             int timeDuration /*sec*/, ref string msg)
        {
            TimeSpan mTime = DateTime.Now - startTime;
            if (mTime.TotalMilliseconds >= timeDuration)
            {
                string currentLSID = agv.CurrentLSID.ToString();
                if (currentLSID.Equals(prevLSID))
                {
                    msg = "OK " + agv.ID + " is not moving and stay at " + prevLSID;
                    checkRun = TestConstants.CHECK_END;
                }
                else
                {
                    msg = "after " + timeDuration + " sec " + agv.ID + " is moving from" + prevLSID + " to " +
                          currentLSID;
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        public static void WaitUntilAgvAllJobsFinished(ref int checkRun, Agv agv, DateTime startTime,
                                                       int timeDuration, ref string msg)
        {
            bool allFinished = true;
            int idx = 0;
            for (int i = 0; i < agv.Jobs.GetArray().Length; i++)
            {
                var job = (Job) agv.Jobs.GetArray()[i];
                if (job.State < Job.STATE.FINISHED)
                {
                    allFinished = false;
                    idx = i;
                    break;
                }
            }

            if (allFinished)
            {
                msg = "OK All jobs are FINISHED";
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    var job = (Job) agv.Jobs.GetArray()[idx];
                    msg = "after " + timeDuration + " min " + job.ID + " state is : " + job.State.ToString();
                    ;
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        public static void WaitUntilAllTransportFinished(ref int checkRun, Project project,
                                                         DateTime startTime,
                                                         int timeDuration, ref string msg)
        {
            bool allFinished = true;
            int idx = 0;
            for (int i = 0; i < project.Transports.GetArray().Length; i++)
            {
                var transport = (Transport) project.Transports.GetArray()[i];
                if (transport.State < Transport.STATE.FINISHED)
                {
                    allFinished = false;
                    idx = i;
                    break;
                }
            }

            if (allFinished)
            {
                msg = "OK All transports are FINISHED";
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    var transport = (Transport) project.Transports.GetArray()[idx];
                    msg = "after " + timeDuration + " min " + transport.ID + " state is : " + transport.State.ToString();
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        /// <summary>
        /// WaitUntilAgvPassNode
        /// </summary>
        public static void WaitUntilAgvPassNode(ref int checkRun, Agv agv, DateTime startTime,
                                                int timeDuration, string lsid, ref string msg)
        {
            if (agv.CurrentLSID.ToString().Equals(lsid))
            {
                msg = "OK " + agv.ID + " at  " + lsid;
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    msg = "after " + timeDuration + " min " + agv.ID + " lsid is : " + agv.CurrentLSID;
                    ;
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        // end WaitUntilAgvPassNode

        public static void WaitUntilAgvsEmpty(ref int checkRun, Project project,
                                              DateTime startTime, int timeDuration, ref string msg)
        {
            Agv agv;
            bool allEmpty = true;
            int idx = 0;
            for (int i = 0; i < project.Agvs.GetArray().Length; i++)
            {
                agv = (Agv) project.Agvs.GetArray()[i];
                if (agv.Loaded)
                {
                    allEmpty = false;
                    idx = i;
                    msg = "after " + timeDuration + " min " + agv.ID + " is still loaded";
                    break;
                }
            }

            if (allEmpty)
            {
                msg = "OK All Agvs are empty ";
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. n*60 sec.
                {
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }


        /// <summary>
        /// WaitUntilAgvsAtInitialPositions
        /// </summary>
        public static void WaitUntilAgvsAtInitialPositions(ref int checkRun, Project project,
                                                           DateTime startTime, int timeDuration, Hashtable agvsInitialID,
                                                           ref string msg)
        {
            Agv agv;
            bool allInitial = true;
            int idx = 0;
            for (int i = 0; i < project.Agvs.GetArray().Length; i++)
            {
                agv = (Agv) project.Agvs.GetArray()[i];
                if (!agv.CurrentLSID.ToString().Equals(agvsInitialID[agv.ID.ToString()].ToString()))
                {
                    allInitial = false;
                    idx = i;
                    msg = "after " + timeDuration + " min " + agv.ID + " lsid is : " + agv.CurrentLSID;
                    break;
                }
            }

            if (allInitial)
            {
                msg = "OK All agvs at Initial positions";
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. n*60 sec.
                {
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        // end WaitUntilAgvsAtInitialPositions

        /// <summary>
        /// WaitUntilAgvThisJobFinished
        /// </summary>
        public static void WaitUntilAgvThisJobFinished(ref int checkRun, Agv agv, string locID, DateTime startTime,
                                                       int timeDuration, ref string msg)
        {
            if (agv.CurrentLSID.ToString().Equals(locID))
            {
                msg = "OK " + agv.ID + " at  " + locID;
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    msg = "after " + timeDuration + " min " + agv.ID + " lsid is : " + agv.CurrentLSID;
                    ;
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        // end WaitUntilAgvThisJobFinished

        public static void WaitUntilAgvsState(ref int checkRun, Agv[] agvs, DateTime startTime,
                                              int timeDuration, Mover.STATE state, ref string msg)
        {
            bool allState = true;
            int idx = 0;
            for (int i = 0; i < agvs.Length; i++)
            {
                if (!(agvs[i].State == state))
                {
                    allState = false;
                    idx = i;
                    break;
                }
            }

            if (allState)
            {
                msg = "OK Agvs state is " + state;
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    msg = "after " + timeDuration + " min " + agvs[idx].ID + " state is : " + agvs[idx].State.ToString();
                    ;
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        public static void WaitUntilAgvsStateReadyOrReadyCharging(ref int checkRun, Agv[] agvs, DateTime startTime,
                                                                  int timeDuration, ref string msg)
        {
            bool allState = true;
            int idx = 0;
            for (int i = 0; i < agvs.Length; i++)
            {
                if (agvs[i].State == Mover.STATE.READY ||
                    agvs[i].State == Mover.STATE.READY_CHARGING)
                {
                    allState = true;
                }
                else
                {
                    allState = false;
                    idx = i;
                    break;
                }
            }

            if (allState)
            {
                msg = "OK Agvs state is READY or READY CHARGING";
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    msg = "after " + timeDuration + " min " + agvs[idx].ID + " state is : " + agvs[idx].State.ToString();
                    ;
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        public static void WaitUntilAgvState(ref int checkRun, Agv agv, DateTime startTime,
                                             int timeDuration, Mover.STATE state, ref string msg)
        {
            if (agv.State == state)
            {
                msg = "OK Agv state is " + state;
                checkRun = TestConstants.CHECK_END;
            }
            else
            {
                TimeSpan mTime = DateTime.Now - startTime;
                if (mTime.TotalMilliseconds >= timeDuration*60000) // wait max. 120 sec.
                {
                    msg = "after " + timeDuration + " min " + agv.ID + " state is : " + agv.State.ToString();
                    ;
                    checkRun = TestConstants.CHECK_END;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static string GetBuildAndTestInfo(string root)
        {
            //string root = m_Project.Facilities["Tests"].Parameters["TestCenterRoot"].ValueAsString;
            //string root = @"C:\EpiaTestCenter2\AutoTestCenter\Main\Source";
            string SetupInfoOutputFilename = "TestAutoDeploymentOutputLog.txt";
            StreamReader reader = File.OpenText(Path.Combine(root, SetupInfoOutputFilename));
            string TestedInfo = reader.ReadLine();
            reader.Close();
            return TestedInfo;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="logger"></param>
        /// <returns></returns>
        public static string UpdateSetupInfoFile(ref Logger logger, string root, string installedApp, int failedCount)
        {
            //string root = m_Project.Facilities["Tests"].Parameters["TestCenterRoot"].ValueAsString;
            //string root = @"C:\EpiaTestCenter2\AutoTestCenter\Main\Source";
            //string root = getRootPath() + @"EpiaAutoTestCenter";
            string SetupInfoOutputFilename = "TestAutoDeploymentOutputLog.txt";
            StreamReader reader = File.OpenText(Path.Combine(root, SetupInfoOutputFilename));
            string TestedInfo = reader.ReadLine();
            reader.Close();

            string infoOutputFile = Path.Combine(root, SetupInfoOutputFilename);
            var file = new FileInfo(infoOutputFile);
            File.SetAttributes(file.FullName, FileAttributes.Normal);

            // empty file
            StreamWriter writer = File.CreateText(Path.Combine(root, SetupInfoOutputFilename));
            writer.Close();

            string testedInfodFile = Path.Combine(root, SetupInfoOutputFilename + "_Tested.txt");
            if (!File.Exists(testedInfodFile))
            {
                logger.LogMessageToFile("1 -- Not Exist ---> " + testedInfodFile);
                //FileStream fs = File.Create(testedInfodFile);
                File.AppendAllText(Path.Combine(root, SetupInfoOutputFilename + "_Tested.txt"),
                                   TestedInfo + "" + Environment.NewLine);
                //fs.Close();
            }
            else
            {
                logger.LogMessageToFile("2 -- Exist ---> " + testedInfodFile);
                //FileInfo tested = new FileInfo(testedInfodFile);
                //File.SetAttributes(tested.FullName, FileAttributes.Normal);
                // append record at end of file
                writer = File.AppendText(Path.Combine(root, SetupInfoOutputFilename + "_Tested.txt"));
                writer.WriteLine(Environment.NewLine);
                writer.WriteLine(TestedInfo);
                writer.Close();
            }
            // end testing  // 
            // Status will be update in Deployment Tester.cs by TestWorking in X: driver
            //string workerStatusFile = Path.Combine(root, TestConstants.TEST_WORKER_STATUS_FILE);
            //FileInfo fileWorker = new FileInfo(workerStatusFile);
            //File.SetAttributes(fileWorker.FullName, FileAttributes.Normal);

            //StreamWriter writerWorker = File.CreateText(workerStatusFile);
            //writerWorker.WriteLine("false");
            //writerWorker.Close();

            // logger.LogMessageToFile("write testWorker: "+Path.Combine(  root,  TestConstants.TEST_WORKER_STATUS_FILE));

            // write test result file
            if (installedApp.Equals("Installed:Etricc 5"))
            {
                string testResultsFile = Path.Combine(root, "TestResults.txt");
                var testResults = new FileInfo(testResultsFile);
                File.SetAttributes(testResults.FullName, FileAttributes.Normal);

                StreamWriter writerResult = File.CreateText(testResultsFile);
                if (failedCount == 0)
                    writerResult.WriteLine("OK");
                else
                    writerResult.WriteLine("Failed" + ":failed count--->" + failedCount);

                writerResult.Close();

                logger.LogMessageToFile("write testResults: " + Path.Combine(root, "TestResults.txt"));
            }
            return TestedInfo;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static bool IsTestMonitorRunning()
        {
            Process[] pTestMonitor = Process.GetProcessesByName("TestScreen");
            try
            {
                if (pTestMonitor[0].Responding)
                    return true;
                else
                {
                    pTestMonitor[0].Kill();
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="agv"></param>
        /// <returns></returns>
        public static bool IsJobsFinished(Agv agv)
        {
            bool finished = true;
            Object[] jobs = agv.Jobs.GetArray();
            if (jobs.Length > 0)
            {
                Job job;
                for (int i = 0; i < jobs.Length; i++)
                {
                    job = (Job) jobs[i];
                    if (job.State < Job.STATE.FINISHED)
                    {
                        finished = false;
                        break;
                    }
                }
                job = null;
            }
            return finished;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="agv"></param>
        /// <param name="logger"></param>
        /// <returns></returns>
        public static bool IsAgvOperational(Agv agv, ref Logger logger)
        {
            if (agv.IsReady() ||
                agv.State.Equals(Mover.STATE.EXECUTING) ||
                agv.State.Equals(Mover.STATE.CHARGING) ||
                agv.State.Equals(Mover.STATE.READY_CHARGING))
            {
                logger.LogMessageToFile(agv.ID + "=1 agvstate: " + agv.State.ToString());
                return true;
            }
            else
            {
                logger.LogMessageToFile(agv.ID + "=2 agvstate: " + agv.State.ToString());
                return false;
            }
        }

        public static bool IsAgvAtParkLocation(Agv agv, Project project, ref string msg)
        {
            bool parked = false;
            string[] parkids = TestData.GetLayoutParkIDS(project);
            msg = agv.ID + " not at park location";

            for (int i = 0; i < parkids.Length; i++)
            {
                //msg += agv.ID.ToString() + " at " + agv.CurrentLSID.ToString() + " and park location is:" + parkids[i] +System.Environment.NewLine;
                if (agv.CurrentLSID.ToString().Equals(parkids[i]))
                {
                    msg = agv.ID + " at park location:" + parkids[i];
                    parked = true;
                    break;
                }
            }
            return parked;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="resultFile"></param>
        /// <param name="layout"></param>
        /// <param name="logger"></param>
        /// <param name="failedCounter"></param>
        /// <param name="testOverview"></param>
        /// <param name="testInputData"></param>
        /// <param name="sendMail"></param>
        public static void SendTestResultToDevelopers(string resultFile, string layout, string buildType,
                                                      ref Logger logger, int failedCounter,
                                                      string testOverview, string testInputData, string sendMail)
        {
            try
            {
                var oMsg = new MailMessage();
                var oAttch = new Attachment(resultFile); //, System.Web.Mail.MailEncoding.Base64);; 
                SendEmailTo(resultFile, ref oMsg, ref oAttch, layout, buildType, failedCounter, testOverview,
                            testInputData, sendMail);

                logger.LogMessageToFile("--------------------------------");
                logger.LogMessageToFile("SmtpServer: " + ConstCommon.SMTP_SERVERID);
                logger.LogMessageToFile("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx send mail ======: " + sendMail);

                var client = new SmtpClient();
                client.Host = ConstCommon.SMTP_SERVERID;

                try
                {
                    client.Send(oMsg);
                }
                catch (Exception ex)
                {
                    logger.LogMessageToFile("The following exception occurred: " + ex);
                    //check the InnerException
                    while (ex.InnerException != null)
                    {
                        logger.LogMessageToFile("--------------------------------");
                        logger.LogMessageToFile("The following InnerException reported: " + ex.InnerException);
                        ex = ex.InnerException;
                    }
                }
                logger.LogMessageToFile("email sent to developers ");
                oMsg = null;
                oAttch = null;
            }
            catch (Exception e)
            {
                logger.LogTestException("send mail : " + e.Message, e.StackTrace);
                Console.WriteLine("{0} Exception caught.", e);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="agvs"></param>
        /// <returns></returns>
        public static bool IsAllAgvsUnloaded(Agv[] agvs)
        {
            for (int i = 0; i < agvs.Length; i++)
            {
                if (agvs[i].Loaded)
                    return false;
            }
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="agvs"></param>
        /// <param name="agvParkIDs"></param>
        /// <returns></returns>
        public static bool IsAllAgvsParked(Agv[] agvs, string[] agvParkIDs)
        {
            for (int i = 0; i < agvs.Length; i++)
            {
                if (!agvs[i].CurrentLSID.Equals(agvParkIDs[i]))
                    return false;
            }
            return true;
        }


        public static void AddTestInfoToExcel(ref Worksheet sheet, TestConstants.TESTINFO testinfo)
        {
            string today = DateTime.Now.ToString("MMMM-dd");
            sheet.Cells[1, 1] = today;
            sheet.Cells[1, 2] = "Test Scenarios";
            //xSheet.Columns.AutoFit();
            sheet.Cells.set_Item(1, 3, "Test Machine: " + Environment.MachineName);
            sheet.Cells.set_Item(2, 3, "OS : " + testinfo.oS_12);
            sheet.Cells.set_Item(3, 3, "OS version: " + Environment.OSVersion);
            sheet.Cells.set_Item(4, 3, "E'pia version: " + testinfo.epiaDeployPath_2);
            sheet.Cells.set_Item(5, 3, "Build type:: " + testinfo.buildType_3);
            sheet.Cells.set_Item(6, 3, "Build Path: " + testinfo.buildInstallScriptDir_7);
            sheet.Cells.set_Item(7, 3, "Layout: " + testinfo.projectFile_6);
            sheet.Cells.set_Item(8, 3, "TestTools version: " + testinfo.testToolsVersion_5);
        }

        public static void AddTestTotalCounterToExcel(ref Worksheet sheet, int beginRow,
                                                      int totTestCnt, int totPassCnt, int totFailCnt)
        {
            sheet.Cells.set_Item(beginRow, 1, "Total tests: ");
            sheet.Cells.set_Item(beginRow + 1, 1, "Total Passes: ");
            sheet.Cells.set_Item(beginRow + 2, 1, "Total Failed: ");

            sheet.Cells.set_Item(beginRow, 2, totTestCnt);
            sheet.Cells.set_Item(beginRow + 1, 2, totPassCnt);
            sheet.Cells.set_Item(beginRow + 2, 2, totFailCnt);
        }

        public static void AddLengendeToExcel(ref Worksheet sheet, ref Range xRange, int beginRow)
        {
            sheet.Cells.set_Item(beginRow, 2, "Legende");
            xRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xRange = sheet.get_Range("B" + (beginRow), "B" + (beginRow));
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            sheet.Cells.set_Item(beginRow + 1, 2, "Pass");
            xRange = sheet.get_Range("B" + (beginRow + 1), "B" + (beginRow + 1));
            xRange.Interior.ColorIndex = TestConstants.EXCEL_GREEN;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            sheet.Cells.set_Item(beginRow + 2, 2, "Fail");
            xRange = sheet.get_Range("B" + (beginRow + 2), "B" + (beginRow + 2));
            xRange.Interior.ColorIndex = TestConstants.EXCEL_RED;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            sheet.Cells.set_Item(beginRow + 3, 2, "Exception");
            xRange = sheet.get_Range("B" + (beginRow + 3), "B" + (beginRow + 3));
            xRange.Interior.ColorIndex = TestConstants.EXCEL_PINK;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            sheet.Cells.set_Item(beginRow + 4, 2, "Untested");
            xRange = sheet.get_Range("B" + (beginRow + 4), "B" + (beginRow + 4));
            xRange.Interior.ColorIndex = TestConstants.EXCEL_YELLOW;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            //sheet.Cells.set_Item(Counter + 10 + 7, 2, "TotalPhysicalMemory:" + TPhysicalMem + " MB");
            //sheet.Cells.set_Item(Counter + 11 + 7, 2, "AvailablePhysicalMemory:" + APhysicalMem + " MB");
            //sheet.Cells.set_Item(Counter + 12 + 7, 2, "TotalVirtualMemory:" + TVirtualMem + " MB");
            //sheet.Cells.set_Item(Counter + 13 + 7, 2, "AvailableVirtualMemory:" + AVirtualMem + " MB");
            sheet.Columns.AutoFit();
        }

        public static void AddLengendeToExcel(ref Worksheet sheet, int beginRow)
        {
            sheet.Cells.set_Item(beginRow, 2, "Legende");
            Range xRange = sheet.get_Range("B" + beginRow, "B" + beginRow);
            xRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xRange = sheet.get_Range("B" + (beginRow), "B" + (beginRow));
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            sheet.Cells.set_Item(beginRow + 1, 2, "Pass");
            xRange = sheet.get_Range("B" + (beginRow + 1), "B" + (beginRow + 1));
            xRange.Interior.ColorIndex = TestConstants.EXCEL_GREEN;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            sheet.Cells.set_Item(beginRow + 2, 2, "Fail");
            xRange = sheet.get_Range("B" + (beginRow + 2), "B" + (beginRow + 2));
            xRange.Interior.ColorIndex = TestConstants.EXCEL_RED;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            sheet.Cells.set_Item(beginRow + 3, 2, "Exception");
            xRange = sheet.get_Range("B" + (beginRow + 3), "B" + (beginRow + 3));
            xRange.Interior.ColorIndex = TestConstants.EXCEL_PINK;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            sheet.Cells.set_Item(beginRow + 4, 2, "Untested");
            xRange = sheet.get_Range("B" + (beginRow + 4), "B" + (beginRow + 4));
            xRange.Interior.ColorIndex = TestConstants.EXCEL_YELLOW;
            xRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            xRange.Borders.Weight = XlBorderWeight.xlMedium;

            //sheet.Cells.set_Item(Counter + 10 + 7, 2, "TotalPhysicalMemory:" + TPhysicalMem + " MB");
            //sheet.Cells.set_Item(Counter + 11 + 7, 2, "AvailablePhysicalMemory:" + APhysicalMem + " MB");
            //sheet.Cells.set_Item(Counter + 12 + 7, 2, "TotalVirtualMemory:" + TVirtualMem + " MB");
            //sheet.Cells.set_Item(Counter + 13 + 7, 2, "AvailableVirtualMemory:" + AVirtualMem + " MB");
            sheet.Columns.AutoFit();
        }

        public static void GetTestInformation(string testFileDirectory, ref TestConstants.TESTINFO myInfo)
        {
            //Add table headers going cell by cell.
            string infoLine = GetBuildAndTestInfo(testFileDirectory); // info is the whole line in outputlog.txt file

            var tokens = new StringCollection();
            if (infoLine != null && infoLine.Length > 0)
            {
                tokens.AddRange(infoLine.Split(new[] {','}));
                string timeInfo_0 = tokens[0].Trim();

                myInfo.installedApp_1 = tokens[1].Trim();
                myInfo.epiaDeployPath_2 = tokens[2].Trim();
                myInfo.buildType_3 = tokens[3].Trim();
                myInfo.buildApplication_4 = tokens[4].Trim();
                myInfo.testToolsVersion_5 = tokens[5].Trim();
                myInfo.projectFile_6 = tokens[6].Trim();
                myInfo.buildInstallScriptDir_7 = tokens[7].Trim();
                myInfo.demo_9 = tokens[9].Trim();
                myInfo.sendMail_10 = tokens[10].Trim();
                myInfo.excelShow_11 = (tokens[11].Trim().ToLower().StartsWith("invisible")) ? false : true;
                myInfo.oS_12 = tokens[12].Trim();
                myInfo.testDirectory_13 = tokens[13].Trim();
                myInfo.testAppWorkingDir_14 = tokens[14].Trim();
                myInfo.autoTestMode_15 = (tokens[15].Trim().ToLower().StartsWith("true")) ? true : false;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="xSheet"></param>
        /// <param name="epiaPath"></param>
        /// <param name="testLayout"></param>
        /// <param name="buildPath"></param>
        /// <param name="demo"></param>
        /// <param name="sendMail"></param>
        /// <returns></returns>
        public static string WriteExcelWorkSheetHeader(
            ref string installedApp_1, ref string epiaAppPath_2,
            ref string buildType_3, ref string buildApp_4,
            ref string testToolsVersion_5, ref string projectFile_6,
            ref string buildInstallScriptPath_7, ref string demo_9,
            ref string sendMail_10, ref bool excelShow_11,
            ref string os_12, ref string testFileDirectory_13,
            ref string testAppWorkingDirectory_14, ref bool AutoTestMode_15)
        {
            //Add table headers going cell by cell.
            string info = GetBuildAndTestInfo(testFileDirectory_13); // info is the whole line in outputlog.txt file

            var tokens = new StringCollection();
            if (info != null && info.Length > 0)
            {
                tokens.AddRange(info.Split(new[] {','}));
                string timeInfo_0 = tokens[0].Trim();
                installedApp_1 = tokens[1].Trim();
                epiaAppPath_2 = tokens[2].Trim();
                buildType_3 = tokens[3].Trim();
                buildApp_4 = tokens[4].Trim();
                testToolsVersion_5 = tokens[5].Trim();
                projectFile_6 = tokens[6].Trim();
                buildInstallScriptPath_7 = tokens[7].Trim();
                demo_9 = tokens[9].Trim();
                sendMail_10 = tokens[10].Trim();
                excelShow_11 = (tokens[11].Trim().ToLower().StartsWith("invisible")) ? false : true;
                os_12 = tokens[12].Trim();
                testFileDirectory_13 = tokens[13].Trim();
                testAppWorkingDirectory_14 = tokens[14].Trim();
                AutoTestMode_15 = (tokens[15].Trim().ToLower().StartsWith("true")) ? true : false;
            }
            return info;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="xSheet"></param>
        /// <param name="result"></param>
        /// <param name="sCounter"></param>
        /// <param name="sTSName"></param>
        /// <param name="testData"></param>
        public static void WriteExcelWorkSheetTestResultOfThisCase(ref Worksheet xSheet, ref Range xRange,
                                                                   int result, int sCounter,
                                                                   string sTSName, string testData)
        {
            string time = DateTime.Now.ToString("HH:mm:ss");
            int row = sCounter;
            xSheet.Cells.set_Item(row, 1, time);
            xSheet.Cells.set_Item(row, 2, sTSName);
            xRange = xSheet.get_Range("B" + row, "B" + row);
            switch (result)
            {
                case TestConstants.TEST_PASS:
                    xRange.Interior.ColorIndex = TestConstants.EXCEL_GREEN;
                    break;
                case TestConstants.TEST_FAIL:
                    xRange.Interior.ColorIndex = TestConstants.EXCEL_RED;
                    xSheet.Cells.set_Item(row, 3, testData);
                    break;
                case TestConstants.TEST_EXCEPTION:
                    xRange.Interior.ColorIndex = TestConstants.EXCEL_PINK;
                    break;
                case TestConstants.TEST_UNDEFINED:
                    xRange.Interior.ColorIndex = TestConstants.EXCEL_YELLOW;
                    xSheet.Cells.set_Item(row, 3, testData);
                    break;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xPath"></param>
        /// <param name="oMsg"></param>
        /// <param name="oAttch"></param>
        /// <param name="layout"></param>
        /// <param name="failedCounter"></param>
        /// <param name="testOverview"></param>
        /// <param name="testInputData"></param>
        /// <param name="sendMail"></param>
        public static void SendEmailTo(string xPath, ref MailMessage oMsg,
                                       ref Attachment oAttch, string layout, string buildType, int failedCounter,
                                       string testOverview, string testInputData, string sendMail)
        {
            int sendHour = DateTime.Now.Hour;
            oMsg.From = new MailAddress("teamsystems@egemin.be");
            //msg.Subject = "Greetings";
            //msg.Body = "This is a  message.";

            // TODO: Replace with recipient e-mail address.
            if (sendMail.ToLower().StartsWith("false"))
            {
                oMsg.To.Add("jiemin.shi@egemin.be");
                oMsg.Subject = "Only ME(" + layout + ")[" + buildType + "]" + DateTime.Now.ToString("ddMMM-HH:mm")
                               + "-[" + Environment.MachineName + "]";
            }
            else
            {
                if (failedCounter > 0)
                {
                    string strAll =
                        "jiemin.shi@egemin.be;Wim.VanBetsbrugge@egemin.be;Jan.Wielemans@egemin.be;Dirk.Declercq@egemin.be;Gunther.Storme@egemin.be;Walter.DeFeyter@egemin.be;Kurt.Schelfthout@egemin.be";
                    oMsg.To.Add(strAll);
                    //oMsg.To = "jiemin.shi@egemin.be;jiemin.shi@egemin.be;";
                    oMsg.Subject = "E'pia Nightly Test Result (" + layout + ")[" + buildType + "]" +
                                   DateTime.Now.ToString("ddMMM-HH:mm")
                                   + "-[" + Environment.MachineName + "]";
                }
                else
                {
                    oMsg.To.Add("jiemin.shi@egemin.be;");
                    oMsg.Subject = "E'pia Nightly Test OK (" + layout + ")[" + buildType + "]" +
                                   DateTime.Now.ToString("ddMMM-HH:mm")
                                   + "-[" + Environment.MachineName + "]";
                }
            }

            oMsg.IsBodyHtml = true;
            // HTML Body (remove HTML tags for plain text).
            //oMsg.Body = "<HTML><BODY><B>Hello World!</B></BODY></HTML>";
            oMsg.Body = testOverview + testInputData;
            oMsg.Attachments.Add(oAttch);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="procName"></param>
        public static void CloseProcesses(string procName)
        {
            Process[] procCmd = Process.GetProcessesByName(procName);
            try
            {
                for (int i = 0; i < procCmd.Length; i++)
                {
                    if (procCmd[i].Responding)
                        procCmd[i].CloseMainWindow();
                    else
                        procCmd[i].Kill();
                }
            }
            catch
            {
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="procName"></param>
        public static void KillProcesses(string procName)
        {
            Process[] pExcel = Process.GetProcessesByName(procName);
            try
            {
                for (int i = 0; i < pExcel.Length; i++)
                    pExcel[i].Kill();
            }
            catch
            {
            }
        }

        #region Remote Logging		

        /// <summary>
        /// 
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="useTestMonitor"></param>
        /// <param name="m_Project"></param>
        public static void RemoteLogMessage(string msg, bool useTestMonitor, Project m_Project)
        {
            if (useTestMonitor)
            {
                msg = "\t" + msg;
                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                m_Project.Facilities["Tests"].Parameters["LogMessage"].ValueAsString = logLine;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="testid"></param>
        /// <param name="useTestMonitor"></param>
        /// <param name="m_Project"></param>
        public static void RemoteLogTestReset(string testid, bool useTestMonitor, Project m_Project)
        {
            if (useTestMonitor)
            {
                string msg = " ------    Test: " + testid + "-------   reset  ------";
                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                m_Project.Facilities["Tests"].Parameters["LogMessage"].ValueAsString = logLine;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="testid"></param>
        /// <param name="useTestMonitor"></param>
        /// <param name="m_Project"></param>
        public static void RemoteLogTestRunStartup(string testid, bool useTestMonitor, Project m_Project)
        {
            if (useTestMonitor)
            {
                string msg = " ------    Test: " + testid + "-------   start run  ------";
                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                m_Project.Facilities["Tests"].Parameters["LogMessage"].ValueAsString = logLine;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ExMsg"></param>
        /// <param name="useTestMonitor"></param>
        /// <param name="m_Project"></param>
        public static void RemoteLogTestException(string ExMsg, bool useTestMonitor, Project m_Project)
        {
            if (useTestMonitor)
            {
                string msg = " ------    Exception: " + ExMsg;
                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                m_Project.Facilities["Tests"].Parameters["LogMessage"].ValueAsString = logLine;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="locID"></param>
        /// <param name="agvID"></param>
        /// <param name="useTestMonitor"></param>
        /// <param name="m_Project"></param>
        public static void RemoteLogCreatedJob(string type, string locID, string agvID, bool useTestMonitor,
                                               Project m_Project)
        {
            if (useTestMonitor)
            {
                string msg = "\t ------    Created Job: " + type + " at " + locID + " by " + agvID;
                ;
                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                m_Project.Facilities["Tests"].Parameters["LogMessage"].ValueAsString = logLine;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="arg"></param>
        /// <param name="useTestMonitor"></param>
        /// <param name="m_Project"></param>
        public static void RemoteLogPassLine(string arg, bool useTestMonitor, Project m_Project)
        {
            if (useTestMonitor)
            {
                string msg = "\t ------   PASS:\t\t " + arg;
                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                m_Project.Facilities["Tests"].Parameters["LogMessage"].ValueAsString = logLine;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="arg"></param>
        /// <param name="useTestMonitor"></param>
        /// <param name="m_Project"></param>
        public static void RemoteLogFailLine(string arg, bool useTestMonitor, Project m_Project)
        {
            if (useTestMonitor)
            {
                string msg = "\t ------   FAIL:\t\t " + arg;
                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                m_Project.Facilities["Tests"].Parameters["LogMessage"].ValueAsString = logLine;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="sourceID"></param>
        /// <param name="destID"></param>
        /// <param name="agvID"></param>
        /// <param name="useTestMonitor"></param>
        /// <param name="m_Project"></param>
        public static void RemoteLogCreatedTransport(string type, string sourceID, string destID, string agvID,
                                                     bool useTestMonitor, Project m_Project)
        {
            if (useTestMonitor)
            {
                string msg = " ------    Created Transport: " + type + " at " + destID + " by " + agvID +
                             " with priority 5";
                if (type == "MOVE")
                    msg = " ------    Created Transport: " + type + " at " + sourceID + " to " + destID + " by " + agvID +
                          " with priority 5";

                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                m_Project.Facilities["Tests"].Parameters["LogMessage"].ValueAsString = logLine;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <param name="sourceID"></param>
        /// <param name="destID"></param>
        /// <param name="agvID"></param>
        /// <param name="priority"></param>
        /// <param name="useTestMonitor"></param>
        /// <param name="m_Project"></param>
        public static void RemoteLogCreatedTransport(string type, string sourceID, string destID, string agvID,
                                                     int priority, bool useTestMonitor, Project m_Project)
        {
            if (useTestMonitor)
            {
                string msg = " ------    Created Transport: " + type + " at " + destID + " by " + agvID +
                             " with priority " + priority;
                if (type == "MOVE")
                    msg = " ------    Created Transport: " + type + " at " + sourceID + " to " + destID + " by " + agvID +
                          " with priority " + priority;

                string logLine = String.Format("{0:G}: {1}.", DateTime.Now, msg);
                m_Project.Facilities["Tests"].Parameters["LogMessage"].ValueAsString = logLine;
            }
        }

        #endregion Remote Logging

        #region Excel

        public static string ConvCHAR(int pPosition)
        {
            string PreChar = "";
            if (pPosition > 26)
            {
                pPosition -= 26;
                PreChar = "A";
            }
            byte aByte = byte.Parse((pPosition + 64).ToString());
            byte[] bytes1 = {aByte, 0x42, 0x43};
            byte[] bytes2 = {0x98, 0xe3};
            var chars = new char[3];

            Decoder d = Encoding.UTF8.GetDecoder();
            int charLen = d.GetChars(bytes1, 0, bytes1.Length, chars, 0);
            // The value of charLen should be 2 now.
            charLen += d.GetChars(bytes2, 0, bytes2.Length, chars, charLen);
            foreach (char c in chars)
            {
                Console.Write("U+" + ((ushort) c).ToString() + "  ");

                return PreChar + c.ToString();
            }
            return "Need a entry";
        }

        public static string SysToCSPro(string s)
        {
            string x = "";
            if (s.StartsWith("System.Boolean"))
                x = "bool";
            else if (s.StartsWith("System.Int16"))
                x = "short";
            else if (s.StartsWith("System.SByte"))
                x = "sbyte";
            else if (s.StartsWith("System.Byte"))
                x = "byte";
            else if (s.StartsWith("System.UI16"))
                x = "ushort";
            else if (s.StartsWith("System.Int32"))
                x = "int";
            else if (s.StartsWith("System.Int64"))
                x = "long";
            else if (s.StartsWith("System.Char"))
                x = "char";
            else if (s.StartsWith("System.Single"))
                x = "float";
            else if (s.StartsWith("System.Double"))
                x = "double";
            else if (s.StartsWith("System.Decimal"))
                x = "decimal";
            else if (s.StartsWith("System.String"))
                x = "string";
            else if (s.StartsWith("System.Object"))
                x = "object";
            else if (s.StartsWith("System.UInt32"))
                x = "uint";
            else if (s.StartsWith("System.UInt64"))
                x = "ulong";
            else
                x = s;

            if (s.EndsWith("[]"))
                x = x + "[]";
            return x;
        }

        public static void ConvertStringToType(string parName, ref Type type)
        {
            /*
			if (parName == "string")
				type =typeof(string);
			else if (parName == "int")
				type =typeof(int);
			else if (parName == "True" || parName =="False" || parName =="bool")
				type =typeof(bool);
			else if (parName == "double")
				type =typeof(double);
			else if (parName == "float")
				type =typeof(float);
			else if (parName == "object")
				type =typeof(object);
			else if (parName == "byte")
				type =typeof(byte);
			else if (parName == "sbyte")
				type =typeof(sbyte);
			else if (parName == "short")
				type =typeof(short);
			else if (parName == "ushort")
				type =typeof(ushort);
			else if (parName == "long")
				type =typeof(long);
			else if (parName == "uint")
				type =typeof(uint);
			else if (parName == "ulong")
				type =typeof(ulong);
			else if (parName == "char")
				type =typeof(char);
			else if (parName == "decimal")
				type =typeof(decimal);
			else if (parName == "bool")
				type =typeof(bool);
			else if (parName == "System.Text.StringBuilder")
				type =typeof(System.Text.StringBuilder);
			else if (parName == "System.IFormatProvider")
				type =typeof(System.IFormatProvider);
			else if (parName == "System.Array")
				type =typeof(System.Array);
			else if (parName == "System.AppDomain")
				type =typeof(System.AppDomain);
			else if (parName == "System.CharEnumerator")
				type =typeof(System.CharEnumerator);
			else if (parName == "System.Type")
				type =typeof(System.Type);
			else if (parName == "System.Runtime.Serialization.SerializationInfo")
				type =typeof(System.Runtime.Serialization.SerializationInfo);
			else if (parName == "VBIDE.CodePane")
				type =typeof(VBIDE.CodePane);
			else if (parName == "VBIDE.VBProject")
				type =typeof(VBIDE.VBProject);
			else if (parName == "VBIDE.vbext_WindowType")
				type =typeof(VBIDE.vbext_WindowType);
			else if (parName == "VBIDE.AddIn")
				type =typeof(VBIDE.AddIn);
			else if (parName == "VBIDE.Window")
				type =typeof(VBIDE.Window);
			else if (parName == "VBIDE.VBComponent")
				type =typeof(VBIDE.VBComponent);
			else if (parName == "VBIDE.Reference")
				type =typeof(VBIDE.Reference);
			else if (parName == "VBIDE._dispReferences_Events_ItemAddedEventHandler")
				type =typeof(VBIDE._dispReferences_Events_ItemAddedEventHandler);
			else if (parName == "VBIDE._dispReferences_Events_ItemRemovedEventHandler")
				type =typeof(VBIDE._dispReferences_Events_ItemRemovedEventHandler);
			else if (parName == "VBIDE._dispCommandBarControlEvents_ClickEventHandler")
				type =typeof(VBIDE._dispCommandBarControlEvents_ClickEventHandler);
			else if (parName == "VBIDE.VBComponent")
				type =typeof(VBIDE.VBComponent);
			else if (parName == "TESTTYPE")
				type =typeof(int);//Type.GetType("TESTTYPE");
		  */
        }

        #endregion Excel

        #region Nested type: TESTSET

        /// <summary>
        /// Test Set related struct
        /// </summary>
        public struct TESTSET
        {
            public string buildPath;
            public string buildTime;
            public string buildType;
            public bool demo;
            public string epiaPath;
            public string rootPath;
            public bool sendMail;
            public string textInstalled;
            public string timeDeployed;

            public override String ToString()
            {
                String str = " time deployed: " + timeDeployed + " deploy path: " + epiaPath
                             + " build type: " + buildType + " build time: " + buildTime
                             + " build path: " + buildPath + " root Path: " + rootPath
                             + " is demo? " + demo.ToString() + " send mail to developers? " + sendMail.ToString();
                return (str);
            }
        }

        #endregion

        #region // sTestInputParams()

        public static Hashtable xGetTestInputParams(Project project, string name)
        {
            var sTestInputParams = new Hashtable();
            sTestInputParams.Clear();
            switch (name)
            {
                case "TS220001JobParkSemiAutomatic":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "PARK_250_1");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "PBAT_09");
                    }
                    break;
                case "TS220002JobBattSemiAutomatic":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "PARK_BAT");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "PBAT_12");
                    }
                    break;
                case "TS220003JobWaitSemiAutomatic":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "W0070-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "AX054");
                    }
                    break;
                case "TS220005JobPickSemiAutomatic":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-03-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_02_01_01_01");
                    }
                    break;
                case "TS220006JobDropSemiAutomatic":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-02");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "M_08_04_01_02");
                        sTestInputParams.Add("sDestination2ID", "M_08_01_01_01");
                    }
                    break;
                case "TS220034JobFlushing":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sLocationID", "W0070-01-01");
                        sTestInputParams.Add("sLocation2ID", "PARK_BAT");
                        sTestInputParams.Add("sLocation3ID", "PARK_040");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sLocationID", "AX054");
                        sTestInputParams.Add("sLocation2ID", "PBAT_12");
                        sTestInputParams.Add("sLocation3ID", "PBAT_09");
                    }
                    break;
                case "TS220016JobCanceling":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource2ID", "0040-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_08_01_01_01");
                    }
                    break;
                case "TS220017JobCancelCurrent":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0040-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS220019JobExhausted":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0060-03-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-05");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS220020JobAborting":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS220023JobSuspending":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS220027JobReleasing":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS220024JobSuspendCurrent":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS220028JobReleaseCurrent":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS220025JobSuspendAll":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sLocation2ID", "W0070-01-01");
                        sTestInputParams.Add("sLocation3ID", "PARK_BAT");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sLocation2ID", "AX054");
                        sTestInputParams.Add("sLocation3ID", "PBAT_09");
                    }
                    break;
                case "TS220029JobReleaseAll":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sLocation2ID", "W0070-01-01");
                        sTestInputParams.Add("sLocation3ID", "PARK_BAT");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sLocation2ID", "AX054");
                        sTestInputParams.Add("sLocation3ID", "PBAT_09");
                    }
                    break;
                case "TS221006JobPickViaStation":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0040-01-01");
                        sTestInputParams.Add("sStationID", "X0040_013");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sStationID", "X090");
                    }
                    break;
                case "TS221007JobParkViaStation":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "PARK_250_1");
                        sTestInputParams.Add("sStationID", "XSR026_1");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "PBAT_09");
                        sTestInputParams.Add("sStationID", "X090");
                    }
                    break;
                case "TS221008JobBattViaStation":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "PARK_BAT");
                        sTestInputParams.Add("sStationID", "X0040_006");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "PBAT_12");
                        sTestInputParams.Add("sStationID", "X094");
                    }
                    break;
                case "TS221009JobWaitViaStation":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "W0070-01-01");
                        sTestInputParams.Add("sStationID", "CX03");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "AX054");
                        sTestInputParams.Add("sStationID", "X090");
                    }
                    break;
                case "TS221010JobDropViaStation":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0040-01-01");
                        sTestInputParams.Add("sDestinationID", "0040-01-05");
                        sTestInputParams.Add("sStationID", "X0040_013");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sStationID", "X095");
                    }
                    break;
                case "TS300005TransOrderPick":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0060-03-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS300006TransOrderDrop":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0060-03-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS300008TransOrderMove":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0060-03-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS300003TransOrderWait":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0060-03-01");
                        sTestInputParams.Add("sLocationID", "W0420-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sLocationID", "AX054");
                    }
                    break;
                case "TS300072-1-TransOrderExcp1":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "AREA1");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "AREA1");
                    }
                    break;

                case "TS300043TransOrderMode":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS300071TransOrderState":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0040-01-01");
                        sTestInputParams.Add("sDestinationID", "0040-01-05");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS300011TransOrderEdit":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sDestination2ID", "0420-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_2_2_T");
                    }
                    break;
                case "TS300034TransOrderFlush":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS300016TransOrderCancel":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS300023TransOrderSuspend":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS300027TransOrderRelease":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS300031TransOrderFinish":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS300025TransOrderSuspendAll":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-02-01-01");
                        sTestInputParams.Add("sSource3ID", "0070-01-03-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_08_04_01_02");
                        sTestInputParams.Add("sSource3ID", "M_08_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_3_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_2_2_T");
                    }
                    break;
                case "TS300029TransOrderReleaseAll":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-02-01-01");
                        sTestInputParams.Add("sSource3ID", "0070-01-03-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_08_04_01_02");
                        sTestInputParams.Add("sSource3ID", "M_08_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_3_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_2_2_T");
                    }
                    break;
                case "TS300026TransOrderSuspendAllPending":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-02-01-01");
                        sTestInputParams.Add("sSource3ID", "0070-01-03-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_08_04_01_02");
                        sTestInputParams.Add("sSource3ID", "M_08_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_3_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_2_2_T");
                    }
                    break;
                case "TS300030TransOrderReleaseAllPending":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-02-01-01");
                        sTestInputParams.Add("sSource3ID", "0070-01-03-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_08_04_01_02");
                        sTestInputParams.Add("sSource3ID", "M_08_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_3_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_2_2_T");
                    }
                    break;
                case "TS350057MutexAuto":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sStationID", "X0500_046");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sStationID", "X034");
                    }
                    break;
                case "TS300016TransportSourceVia":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0040-01-01");
                        sTestInputParams.Add("sStationID", "CX03");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sStationID", "X090");
                    }
                    break;
                case "TS300017TransportDestinationVia":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0040-01-01");
                        sTestInputParams.Add("sDestinationID", "0040-01-05");
                        sTestInputParams.Add("sStationID", "CX03");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sStationID", "X090");
                    }
                    break;
                case "TS241099WeekPlanBatteryCharge":
                case "TS241047WeekPlanBatteryChargeDisable":
                case "TS241048WeekPlanBatteryChargeDisableAll":
                case "TS241049WeekPlanBatteryChargeDelete":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "PARK_BAT");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "PBAT_12");
                    }
                    break;
                case "TS242099WeekPlanCalibration":
                case "TS242047WeekPlanCalibrationDisable":
                case "TS242048WeekPlanCalibrationDisableAll":
                case "TS242049WeekPlanCalibrationDelete":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "CX01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "CAL");
                    }
                    break;
                case "TS200083AgvStop":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS200071AgvState":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "PARK_BAT");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "PBAT_12");
                    }
                    break;
                case "TS200037AgvRetire":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS200053AgvDeploy":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS200023AgvSuspend":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS200027AgvRelease":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS200012AgvModeRemoved":
                    break;
                case "TS200014AgvModeRemovedAll":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-02-01-01");
                        sTestInputParams.Add("sSource3ID", "0070-01-03-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_08_04_01_02");
                        sTestInputParams.Add("sSource3ID", "M_08_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_1_1_T");
                    }
                    break;
                case "TS200047AgvModeDisable":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS200048AgvModeDisableAll":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-02-01-01");
                        sTestInputParams.Add("sSource3ID", "0070-01-03-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_08_04_01_02");
                        sTestInputParams.Add("sSource3ID", "M_08_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_1_1_T");
                    }
                    break;
                case "TS200058AgvModeSemiAutomatic":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sLocationID", "W0420-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sLocationID", "AX054");
                    }
                    break;

                case "TS200061AgvModeSemiAutomaticAll":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sLocationID", "W0420-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-02-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-02");
                        sTestInputParams.Add("sLocation2ID", "W0450-02");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sLocationID", "AX054");
                        sTestInputParams.Add("sSource2ID", "M_08_04_01_02");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_2_T");
                        sTestInputParams.Add("sLocation2ID", "WM_01_01");
                    }
                    break;
                case "TS309306TransPickDeactiveRestart":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS309307TransDropDeactiveRestart":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS420047LocationDisable":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS420045LocationManual":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "PARK_BAT");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "PBAT_12");
                    }
                    break;
                case "TS460047StationDisable":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "PARK_BAT");
                        sTestInputParams.Add("sStationID", "X0060_093");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "PBAT_12");
                        sTestInputParams.Add("sStationID", "X055");
                    }
                    break;
                case "TS400034LoadFlushAndDiscard":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS400076LoadDiscard":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS305513TransOrderDelay":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS305514TransOrderDivert":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sLocationID", "PARK_BAT");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sLocationID", "PBAT_12");
                    }
                    break;
                case "TS304455LocationClosestHighest":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0060-03-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource3ID", "0040-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_08_04_01_02");
                        sTestInputParams.Add("sSource3ID", "M_08_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                    }
                    break;
                case "TS304457GroupHighestPriority":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0030-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-01-01-01");
                        sTestInputParams.Add("sGroupID", "AREA30");
                        sTestInputParams.Add("sGroup2ID", "0070-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_16_01_01");
                        sTestInputParams.Add("sSource2ID", "M_01_01_01_01");
                        sTestInputParams.Add("sGroupID", "AREA116");
                        sTestInputParams.Add("sGroup2ID", "AREA101");
                    }
                    break;
                case "TS304456LoadClosestHighest":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0060-03-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource3ID", "0040-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_03_01_01_01");
                        sTestInputParams.Add("sSource3ID", "M_05_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        //Agv2 executes three transport orders.
                        //TransportA move from PARK_PUSH to M_01_01_01_01 with LoadID LoadA 
                        //TransportB move from PARK_PUSH to M_03_01_01_01 with LoadID LoadB
                        //TransportC move from PARK_PUSH to M_05_01_01_01 with LoadID LoadC
                        //TransportA, TransportB and TransportC have the same priority 5
                        //The equivalent cost of Agv2 for transportA 196 meq
                        //The equivalent cost of Agv2 for transportB 117 meq
                        //The equivalent cost of Agv2 for transportC 63 meq
                        //Set LoadA and LoadC priority to 0 and LoadB priority to 8 
                        //TransportB executed first. After TransportB state become FINISHED
                        //The equivalent cost of Agv2 for transportA 141 meq 
                        //(ABF_1_1_T to M_01_01_01_01)
                        //The equivalent cost of Agv11 for transportC 95 meq (ABF_1_1_T to M_05_01_01_01)
                        //Agv11 will execute TransportC first
                    }
                    break;
                case "TS305501OrderAssignmentClosest":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0240-01-01");
                        sTestInputParams.Add("sSource2ID", "0060-03-01");
                        sTestInputParams.Add("sSource3ID", "0060-13-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_05_14_01_01");
                        sTestInputParams.Add("sSource2ID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource3ID", "RS_D_D");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_1_1_T");
                        //Agv2 (PARK_PUSH), Agv3 (PBAT_07), Agv9 (PBAT_08) with mode automatic.
                        //Create TransportA Move from M_05_14_01_01 to ABF_1_1_T
                        //The equivalent cost for Agvs to location M_05_14_01_01 are:
                        //Agv2 : 41;Agv3 : 140;Agv9 : 113  Agv2 should be assigned to pick the load 
                        //After Agv2 returned to PARK_PUSH, 
                        //Create TranspoerB Move from M_01_01_01_01 to ABF_1_1_T
                        //The equivalent cost for Agvs to location M_01_01_01_01 are:
                        //Agv2 : 196 ;Agv3 : 137 ; Agv9 : 155  Agv3 should be assigned to pick the load 
                        //After Agv3 returned to PBAT_07, 
                        //Create TransportC Move from RS_D_D to ABF_1_1_T
                        //The equivalent cost for Agvs to location RS_D_D are:
                        //Agv2 : 262 ;Agv3 : 96 ; Agv9 : 53  Agv9 should be assigned to pick the load 
                    }
                    break;
                case "TS305502TransOrderClosest":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0060-03-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource3ID", "0040-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_03_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource3ID", "M_05_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_1_1_T");
                        //Agv2 (PARK_PUSH) executes three transport orders.
                        //Create TransportA move from M_03_01_01_01 to ABF_1_1_T.
                        //Create TransportB move from M_01_01_01_01 to ABF_1_1_T, 
                        //Create TransportC move from M_05_01_01_01 to ABF_1_1_T.
                        //The equivalent cost of Agv2 for transportA 117 meq( 149 meq : ABF_1_1_T to M_03_01_01_01 )
                        //The equivalent cost of Agv2 for transportB 196 meq( 141 meq to ABF_1_1_T to M_01_01_01_01)
                        //The equivalent cost of Agv2 for transportC 63 meq
                    }
                    break;
                case "TS305503OrderAssignmentClosestHighest":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0240-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sSource2ID", "0060-03-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sSource3ID", "0060-13-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_05_14_01_01");
                        sTestInputParams.Add("sSource2ID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource3ID", "RS_D_D");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_1_1_T");
                        //Agv2 (PARK_PUSH),  Agv3 (PBAT_07), Agv9 (PBAT_08) with mode automatic.
                        //Create TransportA Move from M_05_14_01_01 to ABF_1_1_T
                        //The equivalent cost for Agvs to location M_05_14_01_01 are:
                        //Agv2 : 41;Agv3 : 140;Agv9 : 113  Agv2 should be assigned to pick the load 
                        //After Agv2 returned to PARK_PUSHK, 
                        //Create TransportB Move from M_01_01_01_01 to ABF_1_1_T
                        //The equivalent cost for Agvs to location M_01_01_01_01 are:
                        //Agv2 : 196 ;Agv3 : 137 ; Agv9 : 155  Agv3 should be assigned to pick the load 
                        //After Agv3 returned to PBAT_07, 
                        //Create TransportC Move from RS_D_D to ABF_1_1_T
                        //The equivalent cost for Agvs to location 0060-13-01 are:
                        //Agv2 : 192 ;Agv3 : 96 ; Agv9 : 53  Agv9 should be assigned to pick the load 
                    }
                    break;
                case "TS305504TransOrderClosestHighest":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sSource2ID", "0060-03-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sSource3ID", "0040-01-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_05_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_03_01_01_01");
                        sTestInputParams.Add("sSource3ID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_1_1_T");
                        //Agv2 (PARK_PUSH) executes three transport orders.
                        //TransportA move from M_05_01_01_01 to ABF_1_1_T with priority 5, 
                        //TransportB move from M_03_01_01_01 to ABF_1_1_T with priority 8 
                        //TransportC move from M_01_01_01_01 to ABF_1_1_T with priority 5.
                        //The equivalent cost of Agv2 for transportA 63 meq
                        //The equivalent cost of Agv2 for transportB 117 meq
                        //The equivalent cost of Agv2 for transportC 196 meq
                        //TransportB executed first. After TransportB state become FINISHED
                        //The equivalent cost of Agv2 for transportA 95 meq 
                        //(ABF_1_1_T to M_05_01_01_01)
                        //The equivalent cost of Agv2 for transportC 141 meq (ABF_1_1_T to M_01_01_01_01)
                    }
                    break;
                case "TS305505TransOrderOldest":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0060-03-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestination2ID", "0360-01-01");
                        sTestInputParams.Add("sSource3ID", "0040-01-01");
                        sTestInputParams.Add("sDestination3ID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_08_04_01_02");
                        sTestInputParams.Add("sSource3ID", "M_08_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination2ID", "ABF_1_1_T");
                        sTestInputParams.Add("sDestination3ID", "ABF_1_1_T");
                    }
                    break;
                case "TS305507SchedulesDeadlockRulesVia":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0030-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sLocationID", "X0030_003");
                        sTestInputParams.Add("sStationID", "X0500_018");
                        sTestInputParams.Add("sScheduleID", "AREA30.DEADLOCK");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "TBC_1_1_D");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        sTestInputParams.Add("sLocationID", "XTBC1B");
                        sTestInputParams.Add("sStationID", "X030");
                        sTestInputParams.Add("sScheduleID", "AREA_LAYOUT_FLV.DEADLOCK");
                    }
                    break;
                case "TS305508ScheduleBattRulesQueueSimLow":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "BAT_080_1");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sLocationID", "BAT_080_1");
                    }
                    break;

                case "TS330056RoutingDynamic":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                        sTestInputParams.Add("sLocationID", "X0500_049");
                        sTestInputParams.Add("sStationID", "X059"); // disabled station
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "AX054");
                    }
                    break;
                case "TS300063TransOrderDoublePlay":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0240-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-02-01-03-01");
                        sTestInputParams.Add("sDestinationID", "0070-02-01-01-01");
                        sTestInputParams.Add("sGroupID", "AREA1");
                        sTestInputParams.Add("sGroup2ID", "0070-02-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_01_16_01_01");
                        sTestInputParams.Add("sDestinationID", "M_01_01_02_01");
                        sTestInputParams.Add("sGroupID", "AREA101");
                        sTestInputParams.Add("sGroup2ID", "AREA116");
                    }
                    break;
                case "TS300064DoublePlayTransReleased":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0240-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-02-01-03-01");
                        sTestInputParams.Add("sDestinationID", "0070-02-01-01-01");
                        sTestInputParams.Add("sGroupID", "AREA1");
                        sTestInputParams.Add("sGroup2ID", "0070-02-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_01_16_01_01");
                        sTestInputParams.Add("sDestinationID", "M_01_01_02_01");
                        sTestInputParams.Add("sGroupID", "AREA101");
                        sTestInputParams.Add("sGroup2ID", "AREA116");
                    }
                    break;
                case "TS830001DBSQLSERVERStopStart":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sDestinationID", "0360-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TESTOPSTELLING"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "AX054");
                    }
                    break;
                default:
                    sTestInputParams.Add("sLocationID", "Default");
                    break;
            }
            return sTestInputParams;
        }

        #endregion // End GetTestInputParams()
    }
}