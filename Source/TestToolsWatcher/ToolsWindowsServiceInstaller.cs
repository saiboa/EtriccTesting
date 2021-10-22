using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration.Install;
using System.ComponentModel;
using System.ServiceProcess;

namespace Egemin.Epia.Testing.TestToolsWatcher
{
    [RunInstaller(true)]
    public partial  class ToolsWindowsServiceInstaller : Installer
    {
        ServiceProcessInstaller processInstaller = new ServiceProcessInstaller();
        ServiceInstaller serviceInstaller = new ServiceInstaller();

        #region Constructors

        public ToolsWindowsServiceInstaller()
        {
            processInstaller.Account = ServiceAccount.LocalSystem;
            serviceInstaller.DisplayName = "QATestsService";
            serviceInstaller.StartType = ServiceStartMode.Manual;

            //must be the same as what was set in Program's constructor
            serviceInstaller.ServiceName = "QATestsService";

            this.Installers.Add(processInstaller);
            this.Installers.Add(serviceInstaller);
        }

        #endregion
    }
}
