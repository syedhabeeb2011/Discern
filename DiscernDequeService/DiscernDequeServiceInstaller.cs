using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;

namespace DiscernDequeService
{
    [RunInstaller(true)]
    public partial class DiscernDequeServiceInstaller : System.Configuration.Install.Installer
    {
        public DiscernDequeServiceInstaller()
        {
            InitializeComponent();
        }
    }
}
