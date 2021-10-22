using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TFS2010AutoDeploymentTool
{
    public class BuildObject
    {
        public string BuildNr { get; set; }
        public string Quality { get; set; }
        public DateTime FinishTime { get; set; }
        public string DripLoc { get; set; }
        public string xMapString { get; set; }
		public string RelativeDropLoc { get; set; }
    }
}
