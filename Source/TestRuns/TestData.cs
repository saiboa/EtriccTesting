using System.Collections;
using Egemin.EPIA.WCS.Core;
using Egemin.EPIA.WCS.Resources;

namespace TestRuns
{
    internal class TestData
    {
        #region // sTestInputParams()     

        public static Hashtable GetTestInputParams(Project project, string name)
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sLocationID", "PBAT_09");
                    }
                    break;
                case "TS220002JobBattSemiAutomatic":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "PARK_BAT");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sLocationID", "PBAT_12");
                    }
                    break;
                case "TS220003JobWaitSemiAutomatic":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "W0070-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS300023TransOrderSuspend":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS300027TransOrderRelease":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS300031TransOrderFinish":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_13_01_01");
                        sTestInputParams.Add("sSource2ID", "M_04_13_01_01");
                        sTestInputParams.Add("sSource3ID", "M_05_13_01_01");
                        sTestInputParams.Add("sDestinationID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestination2ID", "M_04_01_01_01");
                        sTestInputParams.Add("sDestination3ID", "M_05_01_01_01");
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_13_01_01");
                        sTestInputParams.Add("sSource2ID", "M_04_13_01_01");
                        sTestInputParams.Add("sSource3ID", "M_05_13_01_01");
                        sTestInputParams.Add("sDestinationID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestination2ID", "M_04_01_01_01");
                        sTestInputParams.Add("sDestination3ID", "M_05_01_01_01");
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                        sTestInputParams.Add("sLocationID", "W0450-02");
                        //sTestInputParams.Add("sLocationID", "W0420-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "ABF_1_1_T");
                        //sTestInputParams.Add("sLocationID", "AX054");
                        sTestInputParams.Add("sLocationID", "CAL");
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS420045LocationManual":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sLocationID", "PARK_BAT");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
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
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sDestinationID", "AX054");
                    }
                    break;
                case "TS300080TransPickFromGroup":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sGroupID", "0070-01-01");
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sGroupID", "AREA101");
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                    }
                    break;
                case "TS300081TransDropToGroup":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sGroupID", "0070-06-02");
                        sTestInputParams.Add("sDestinationID", "0070-06-02-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sGroupID", "AREA116");
                        sTestInputParams.Add("sDestinationID", "M_01_16_01_01");
                    }
                    break;
                case "TS200080AgvModeSemiToAuto":
                    if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                    {
                        sTestInputParams.Add("sSourceID", "0070-01-01-01-01");
                        sTestInputParams.Add("sSource2ID", "0070-02-01-01-01");
                        sTestInputParams.Add("sSource3ID", "0070-03-01-01-01");
                    }
                    else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
                    {
                        sTestInputParams.Add("sSourceID", "M_01_01_01_01");
                        sTestInputParams.Add("sSource2ID", "M_01_16_01_01");
                        sTestInputParams.Add("sSource3ID", "M_08_01_01_01");
                    }
                    break;
                default:
                    sTestInputParams.Add("sLocationID", "Default");
                    break;
            }
            return sTestInputParams;
        }

        #endregion // End GetTestInputParams()

        #region // getAgvssAgvsInitialID()

        public static Hashtable GetAgvssAgvsInitialID(Project project)
        {
            var sAgvsInitialID = new Hashtable();
            if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
            {
                sAgvsInitialID.Clear();
                sAgvsInitialID.Add("AGV1", "PARK_250_3");
                sAgvsInitialID.Add("AGV2", "0241-01-01-01-01");
                sAgvsInitialID.Add("AGV3", "PARK_060_4");
                sAgvsInitialID.Add("AGV4", "PARK_250_2");
                sAgvsInitialID.Add("AGV5", "PARK_500_1");
                sAgvsInitialID.Add("AGV6", "PARK_250_4");
                sAgvsInitialID.Add("AGV7", "PARK_060_LIFT");
                sAgvsInitialID.Add("AGV8", "PARK_060_1");
                sAgvsInitialID.Add("AGV9", "PARK_240_1");
                sAgvsInitialID.Add("AGV10", "PARK_500_2");
                sAgvsInitialID.Add("AGV11", "PARK_040");
                sAgvsInitialID.Add("AGV12", "PARK_060_2");
            }
            else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
            {
                sAgvsInitialID.Clear();
                sAgvsInitialID.Add("AGV1", "PBAT_01");
                sAgvsInitialID.Add("AGV2", "PARK_PUSH");
                sAgvsInitialID.Add("AGV3", "PBAT_07");
                sAgvsInitialID.Add("AGV4", "PBAT_05");
                sAgvsInitialID.Add("AGV5", "PBAT_03");
                sAgvsInitialID.Add("AGV6", "PBAT_06");
                sAgvsInitialID.Add("AGV7", "PBAT_02");
                sAgvsInitialID.Add("AGV8", "PBAT_04");
                sAgvsInitialID.Add("AGV9", "PBAT_08");
                sAgvsInitialID.Add("AGV10", "X_A01");
                sAgvsInitialID.Add("AGV11", "X_A10");
            }
            else
            {
                sAgvsInitialID.Clear();
                sAgvsInitialID.Add("FLV", "PRK_FLV");
                sAgvsInitialID.Add("TUG", "PRK_TUG");
            }
            return sAgvsInitialID;
        }

        #endregion // End getAgvssAgvsInitialID()

        #region // sAgvsDefaultDropID()

        public static Hashtable GetAgvsDefaultDropID(Project project)
        {
            var sAgvsDefaultDropID = new Hashtable();
            if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
            {
                sAgvsDefaultDropID.Clear();
                sAgvsDefaultDropID.Add("AGV1", "0360-01-01");
                sAgvsDefaultDropID.Add("AGV2", "0360-01-01");
                sAgvsDefaultDropID.Add("AGV3", "0060-04-06");
                sAgvsDefaultDropID.Add("AGV4", "0360-01-01");
                sAgvsDefaultDropID.Add("AGV5", "0360-01-01");
                sAgvsDefaultDropID.Add("AGV6", "0360-01-01");
                sAgvsDefaultDropID.Add("AGV7", "0060-12-01");
                sAgvsDefaultDropID.Add("AGV8", "0360-01-01");
                sAgvsDefaultDropID.Add("AGV9", "0360-01-01");
                sAgvsDefaultDropID.Add("AGV10", "0360-01-01");
                sAgvsDefaultDropID.Add("AGV11", "0360-01-01");
                sAgvsDefaultDropID.Add("AGV12", "0360-01-01");
            }
            else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
            {
                sAgvsDefaultDropID.Clear();
                sAgvsDefaultDropID.Add("AGV1", "ABF_1_1_T");
                sAgvsDefaultDropID.Add("AGV2", "ABF_1_1_T");
                sAgvsDefaultDropID.Add("AGV3", "ABF_1_2_T");
                sAgvsDefaultDropID.Add("AGV4", "ABF_1_1_T");
                sAgvsDefaultDropID.Add("AGV5", "ABF_1_2_T");
                sAgvsDefaultDropID.Add("AGV6", "ABF_1_1_T");
                sAgvsDefaultDropID.Add("AGV7", "ABF_1_1_T");
                sAgvsDefaultDropID.Add("AGV8", "ABF_1_1_T");
                sAgvsDefaultDropID.Add("AGV9", "ABF_1_3_T");
                sAgvsDefaultDropID.Add("AGV10", "ABF_1_1_T");
                sAgvsDefaultDropID.Add("AGV11", "ABF_1_1_T");
            }
            else
            {
                sAgvsDefaultDropID.Clear();
                sAgvsDefaultDropID.Add("FLV", "PRK_FLV");
                sAgvsDefaultDropID.Add("TUG", "PRK_TUG");
            }
            return sAgvsDefaultDropID;
        }

        #endregion // End sAgvsDefaultDropID()

        #region // GetTestAgvs(Egemin.EPIA.WCS.Core.Project project)  OK

        public static Agv[] GetTestAgvs(Project project)
        {
            Agv[] testAgvs;
            if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                testAgvs = new[] {project.Agvs["AGV11"], project.Agvs["AGV3"], project.Agvs["AGV7"]};
            else
                testAgvs = new[] {project.Agvs["AGV2"], project.Agvs["AGV3"], project.Agvs["AGV9"]};

            return testAgvs;
        }

        #endregion // GetTestAgvs()

        #region // GetWaitTime(Egemin.EPIA.WCS.Core.Project project)  OK

        public static int GetWaitTime(Project project)
        {
            if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
                return 2;
            else
                return 6;
        }

        #endregion // GetWaitTime()

        #region // GetLayoutParkIDS(Egemin.EPIA.WCS.Core.Project project ) OK

        public static string[] GetLayoutParkIDS(Project project)
        {
            string[] parkids;
            if (project.ID.ToString().ToUpper().StartsWith("EUROBALTIC"))
            {
                parkids = new[]
                              {
                                  "PARK_040",
                                  "PARK_060_1",
                                  "PARK_060_2",
                                  "PARK_060_3",
                                  "PARK_060_4",
                                  "PARK_060_LIFT",
                                  "PARK_240_1",
                                  "PARK_250_1",
                                  "PARK_250_2",
                                  "PARK_500_3",
                                  "PARK_250_4",
                                  "PARK_500_1",
                                  "PARK_500_2",
                              };
                return parkids;
            }
            else if (project.ID.ToString().ToUpper().StartsWith("TEST"))
            {
                parkids = new[]
                              {
                                  "PARK_PUSH",
                                  "PBAT_01",
                                  "PBAT_01",
                                  "PBAT_02",
                                  "PBAT_03",
                                  "PBAT_04",
                                  "PBAT_05",
                                  "PBAT_06",
                                  "PBAT_07",
                                  "PBAT_08",
                                  "PBAT_09",
                                  "PBAT_10",
                                  "PBAT_11",
                                  "PBAT_12",
                                  "X_A01",
                                  "X_A02",
                                  "X_A03",
                                  "X_A04",
                                  "X_A05",
                                  "X_A06",
                                  "X_A07",
                                  "X_A08",
                                  "X_A09",
                                  "X_A10",
                                  "X_A11",
                                  "X_A12",
                              };
                return parkids;
            }
            else
            {
                parkids = new[]
                              {
                                  "PRK_FLV",
                                  "PRK_TUG",
                              };
                return parkids;
            }
        }

        #endregion // GetLayoutParkIDS(Egemin.EPIA.WCS.Core.Project project)
    }
}