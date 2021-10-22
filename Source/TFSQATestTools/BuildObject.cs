using System;

namespace TFSQATestTools
{
    public class BuildObject
    {
        public string BuildNr { get; set; }
        public string BuildDef { get; set; }
        public string Quality { get; set; }
        public DateTime FinishTime { get; set; }
        public string DripLoc { get; set; }
        public string xMapString { get; set; }
        public string RelativeDropLoc { get; set; }
    }
}