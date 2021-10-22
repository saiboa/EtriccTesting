using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EtriccStatisticsProgTest
{
    static class ReportName
    {
        #region // report name
        public const string PERFORMANCE_VEHICLES_ModeOverview = "Mode: overview";
        public const string PERFORMANCE_VEHICLES_StateOverview = "State: overview";
        public const string PERFORMANCE_VEHICLES_StatusOverview = "Status: overview";
        public const string PERFORMANCE_VEHICLES_StatusCountDayTrend = "Status: count/day-trend";
        public const string PERFORMANCE_VEHICLES_StatusCountTop = "Status: count-top";
        public const string PERFORMANCE_VEHICLES_StatusDurationDayTrend = "Status: duration/day-trend";
        public const string PERFORMANCE_VEHICLES_StatusDurationTop = "Status: duration-top";

        public const string PERFORMANCE_TRANSPORTS_CountBySrcDstGroupHour = "Count by src/dst group (hour)";
        public const string PERFORMANCE_TRANSPORTS_CountBySrcDstGroupDay = "Count by src/dst group (day)";
        public const string PERFORMANCE_TRANSPORTS_CountBySrcDstGroupMonth = "Count by src/dst group (month)";
        public const string PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationHour = "Count by src/dst location or station (hour)";
        public const string PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationDay = "Count by src/dst location or station (day)";
        public const string PERFORMANCE_TRANSPORTS_CountBySrcDstLocationOrStationMonth = "Count by src/dst location or station (month)";

        public const string PERFORMANCE_JOBS_CountByLocationInGroupDay = "Count by location in group (day)";
        public const string PERFORMANCE_JOBS_CountByLocationInGroupMonth = "Count by location in group (month)";
        public const string PERFORMANCE_JOBS_CountByLocationDay = "Count by location (day)";
        public const string PERFORMANCE_JOBS_CountByLocationMonth = "Count by location (month)";

        public const string ANALYSIS_ProjectActivation = "Project activation";
        public const string ANALYSIS_TransportLookupBySrcDstGroup = "Transport lookup by src/dst group";
        public const string ANALYSIS_TransportLookupBySrcDstLocationOrStation = "Transport lookup by src/dst location or station";
        public const string ANALYSIS_TransportWithJobsAndStatusHistory = "Transport with jobs and status history";
        public const string ANALYSIS_LoadHistory = "Load history";

        public const string StatusGraphicalView = "Status: graphical view";



        #endregion
    }
}