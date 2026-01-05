namespace Discern_Notification
{
    partial class DiscernNotificationServiceInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.serviceProcessInstaller_DN = new System.ServiceProcess.ServiceProcessInstaller();
            this.serviceInstaller_DN = new System.ServiceProcess.ServiceInstaller();
            // 
            // serviceProcessInstaller_DN
            // 
            this.serviceProcessInstaller_DN.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.serviceProcessInstaller_DN.Password = null;
            this.serviceProcessInstaller_DN.Username = null;
            // 
            // serviceInstaller_DN
            // 
            this.serviceInstaller_DN.Description = "Discern Notification";
            this.serviceInstaller_DN.DisplayName = "Discern Notification";
            this.serviceInstaller_DN.ServiceName = "DiscernNotificatoinService";
            // 
            // DiscernNotificationServiceInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.serviceProcessInstaller_DN,
            this.serviceInstaller_DN});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller serviceProcessInstaller_DN;
        private System.ServiceProcess.ServiceInstaller serviceInstaller_DN;
    }
}