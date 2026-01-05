namespace SyntBotsDesktopTestingUtility
{
    partial class CallAutomationfrm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CallAutomationfrm));
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtOutPut = new System.Windows.Forms.TextBox();
            this.lblOutput = new System.Windows.Forms.Label();
            this.BaseAssembly = new System.Windows.Forms.Button();
            this.JsonButton = new System.Windows.Forms.Button();
            this.lblResult = new System.Windows.Forms.Label();
            this.btnRunAutomation = new System.Windows.Forms.Button();
            this.txtAssembly = new System.Windows.Forms.TextBox();
            this.lblBaseAssemblyPath = new System.Windows.Forms.Label();
            this.txtJson = new System.Windows.Forms.TextBox();
            this.lblJsonPath = new System.Windows.Forms.Label();
            this.openFileDialogRPA = new System.Windows.Forms.OpenFileDialog();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.LightGray;
            this.panel1.Controls.Add(this.txtOutPut);
            this.panel1.Controls.Add(this.lblOutput);
            this.panel1.Controls.Add(this.BaseAssembly);
            this.panel1.Controls.Add(this.JsonButton);
            this.panel1.Controls.Add(this.lblResult);
            this.panel1.Controls.Add(this.btnRunAutomation);
            this.panel1.Controls.Add(this.txtAssembly);
            this.panel1.Controls.Add(this.lblBaseAssemblyPath);
            this.panel1.Controls.Add(this.txtJson);
            this.panel1.Controls.Add(this.lblJsonPath);
            this.panel1.Location = new System.Drawing.Point(2, 95);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(818, 359);
            this.panel1.TabIndex = 0;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // txtOutPut
            // 
            this.txtOutPut.Location = new System.Drawing.Point(16, 227);
            this.txtOutPut.Multiline = true;
            this.txtOutPut.Name = "txtOutPut";
            this.txtOutPut.Size = new System.Drawing.Size(780, 117);
            this.txtOutPut.TabIndex = 9;
            this.txtOutPut.TextChanged += new System.EventHandler(this.txtOutPut_TextChanged);
            // 
            // lblOutput
            // 
            this.lblOutput.AutoSize = true;
            this.lblOutput.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.lblOutput.Location = new System.Drawing.Point(13, 208);
            this.lblOutput.Name = "lblOutput";
            this.lblOutput.Size = new System.Drawing.Size(46, 16);
            this.lblOutput.TabIndex = 8;
            this.lblOutput.Text = "Output";
            this.lblOutput.Click += new System.EventHandler(this.lblOutput_Click);
            // 
            // BaseAssembly
            // 
            this.BaseAssembly.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(178)))), ((int)(((byte)(34)))));
            this.BaseAssembly.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BaseAssembly.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BaseAssembly.ForeColor = System.Drawing.Color.White;
            this.BaseAssembly.Location = new System.Drawing.Point(721, 85);
            this.BaseAssembly.Name = "BaseAssembly";
            this.BaseAssembly.Size = new System.Drawing.Size(75, 23);
            this.BaseAssembly.TabIndex = 7;
            this.BaseAssembly.Text = "Browse";
            this.BaseAssembly.UseVisualStyleBackColor = false;
            this.BaseAssembly.Click += new System.EventHandler(this.BaseAssembly_Click);
            // 
            // JsonButton
            // 
            this.JsonButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(178)))), ((int)(((byte)(34)))));
            this.JsonButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.JsonButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.JsonButton.ForeColor = System.Drawing.Color.White;
            this.JsonButton.Location = new System.Drawing.Point(721, 35);
            this.JsonButton.Name = "JsonButton";
            this.JsonButton.Size = new System.Drawing.Size(75, 23);
            this.JsonButton.TabIndex = 6;
            this.JsonButton.Text = "Browse";
            this.JsonButton.UseVisualStyleBackColor = false;
            this.JsonButton.Click += new System.EventHandler(this.JsonButton_Click);
            // 
            // lblResult
            // 
            this.lblResult.AutoSize = true;
            this.lblResult.Location = new System.Drawing.Point(46, 252);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(0, 13);
            this.lblResult.TabIndex = 5;
            // 
            // btnRunAutomation
            // 
            this.btnRunAutomation.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(251)))), ((int)(((byte)(178)))), ((int)(((byte)(34)))));
            this.btnRunAutomation.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRunAutomation.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRunAutomation.ForeColor = System.Drawing.Color.White;
            this.btnRunAutomation.Location = new System.Drawing.Point(377, 128);
            this.btnRunAutomation.Name = "btnRunAutomation";
            this.btnRunAutomation.Size = new System.Drawing.Size(161, 35);
            this.btnRunAutomation.TabIndex = 4;
            this.btnRunAutomation.Text = "Run Automation";
            this.btnRunAutomation.UseVisualStyleBackColor = false;
            this.btnRunAutomation.Click += new System.EventHandler(this.btnRunAutomation_Click);
            // 
            // txtAssembly
            // 
            this.txtAssembly.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtAssembly.Location = new System.Drawing.Point(259, 85);
            this.txtAssembly.Name = "txtAssembly";
            this.txtAssembly.Size = new System.Drawing.Size(441, 22);
            this.txtAssembly.TabIndex = 3;
            // 
            // lblBaseAssemblyPath
            // 
            this.lblBaseAssemblyPath.AutoSize = true;
            this.lblBaseAssemblyPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBaseAssemblyPath.Location = new System.Drawing.Point(10, 85);
            this.lblBaseAssemblyPath.Name = "lblBaseAssemblyPath";
            this.lblBaseAssemblyPath.Size = new System.Drawing.Size(243, 16);
            this.lblBaseAssemblyPath.TabIndex = 2;
            this.lblBaseAssemblyPath.Text = "SyntBots Desktop Base Assembly Path";
            this.lblBaseAssemblyPath.Click += new System.EventHandler(this.lblBaseAssemblyPath_Click);
            // 
            // txtJson
            // 
            this.txtJson.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtJson.Location = new System.Drawing.Point(259, 36);
            this.txtJson.Name = "txtJson";
            this.txtJson.Size = new System.Drawing.Size(441, 22);
            this.txtJson.TabIndex = 1;
            // 
            // lblJsonPath
            // 
            this.lblJsonPath.AutoSize = true;
            this.lblJsonPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblJsonPath.Location = new System.Drawing.Point(9, 35);
            this.lblJsonPath.Name = "lblJsonPath";
            this.lblJsonPath.Size = new System.Drawing.Size(92, 16);
            this.lblJsonPath.TabIndex = 0;
            this.lblJsonPath.Text = "Json File Path";
            // 
            // openFileDialogRPA
            // 
            this.openFileDialogRPA.FileName = "openFileDialogRPA";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(2, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(818, 91);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // CallAutomationfrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(818, 451);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "CallAutomationfrm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SyntBots Desktop Testing Utility";
            this.Load += new System.EventHandler(this.CallAutomationfrm_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnRunAutomation;
        private System.Windows.Forms.TextBox txtAssembly;
        private System.Windows.Forms.Label lblBaseAssemblyPath;
        private System.Windows.Forms.TextBox txtJson;
        private System.Windows.Forms.Label lblJsonPath;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Button BaseAssembly;
        private System.Windows.Forms.Button JsonButton;
        private System.Windows.Forms.OpenFileDialog openFileDialogRPA;
        private System.Windows.Forms.TextBox txtOutPut;
        private System.Windows.Forms.Label lblOutput;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}