using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;
using System.Reflection;
using System.Threading.Tasks;
using System.Configuration;

namespace Discern_Notification
{
    [RunInstaller(true)]
    public partial class DiscernNotificationServiceInstaller : System.Configuration.Install.Installer
    {
        public DiscernNotificationServiceInstaller()
        {
            InitializeComponent();
            this.serviceInstaller_DN.ServiceName = GetConfigurationValue("ServiceName");
            this.serviceInstaller_DN.DisplayName = GetConfigurationValue("ServiceName");
            this.serviceInstaller_DN.Description = GetConfigurationValue("ServiceName");
        }

        private static string GetConfigurationValue(string key)
        {


            var service = Assembly.GetAssembly(typeof(DiscernNotificationServiceInstaller));

            Configuration config = ConfigurationManager.OpenExeConfiguration(service.Location);

            if (config.AppSettings.Settings[key] == null)
            {

                throw new IndexOutOfRangeException("Settings collection does not contain the requested key:" + key);

            }


            return config.AppSettings.Settings[key].Value;
        }
    }
}
