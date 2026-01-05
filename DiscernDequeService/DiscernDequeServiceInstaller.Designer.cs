namespace DiscernDequeService
{
    partial class DiscernDequeServiceInstaller
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
            this.serviceProcessInstaller_Deque = new System.ServiceProcess.ServiceProcessInstaller();
            this.serviceInstaller_Deque = new System.ServiceProcess.ServiceInstaller();
            // 
            // serviceProcessInstaller_Deque
            // 
            this.serviceProcessInstaller_Deque.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.serviceProcessInstaller_Deque.Password = null;
            this.serviceProcessInstaller_Deque.Username = null;
            // 
            // serviceInstaller_Deque
            // 
            this.serviceInstaller_Deque.Description = "DiscernDequeService";
            this.serviceInstaller_Deque.DisplayName = "DiscernDequeService";
            this.serviceInstaller_Deque.ServiceName = "DiscernDequeService";
            // 
            // DiscernDequeServiceInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.serviceProcessInstaller_Deque,
            this.serviceInstaller_Deque});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller serviceProcessInstaller_Deque;
        private System.ServiceProcess.ServiceInstaller serviceInstaller_Deque;
    }
}