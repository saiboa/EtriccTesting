namespace Epia3Deployment
{
    partial class Configuration
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Configuration));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtBoxDeployLocation = new System.Windows.Forms.TextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnSaveConfig = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.cmbBuildApp = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cmbBranch = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ckbVersion = new System.Windows.Forms.CheckBox();
            this.ckbNightly = new System.Windows.Forms.CheckBox();
            this.ckbCI = new System.Windows.Forms.CheckBox();
            this.ckbWeekly = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbExcelVisible = new System.Windows.Forms.ComboBox();
            this.button_DeploymentLocation = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.cmbProjectFile = new System.Windows.Forms.ComboBox();
            this.ckbFunctionalTesting = new System.Windows.Forms.CheckBox();
            this.cmbServerRunAs = new System.Windows.Forms.ComboBox();
            this.labServer = new System.Windows.Forms.Label();
            this.ckbMail = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbPlatformTarget = new System.Windows.Forms.ComboBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Project XMLFile:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 45);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Deployment Location:";
            // 
            // txtBoxDeployLocation
            // 
            this.txtBoxDeployLocation.Location = new System.Drawing.Point(156, 45);
            this.txtBoxDeployLocation.Name = "txtBoxDeployLocation";
            this.txtBoxDeployLocation.Size = new System.Drawing.Size(400, 20);
            this.txtBoxDeployLocation.TabIndex = 4;
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            // 
            // btnSaveConfig
            // 
            this.btnSaveConfig.Location = new System.Drawing.Point(504, 226);
            this.btnSaveConfig.Name = "btnSaveConfig";
            this.btnSaveConfig.Size = new System.Drawing.Size(75, 23);
            this.btnSaveConfig.TabIndex = 5;
            this.btnSaveConfig.Text = "Save";
            this.btnSaveConfig.UseVisualStyleBackColor = true;
            this.btnSaveConfig.Click += new System.EventHandler(this.btnSaveConfig_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(423, 226);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // cmbBuildApp
            // 
            this.cmbBuildApp.DisplayMember = "Etricc";
            this.cmbBuildApp.FormattingEnabled = true;
            this.cmbBuildApp.Items.AddRange(new object[] {
            "Epia",
            "Etricc UI",
            "Etricc 5",
            "Etricc+Etricc5",
            "Kimberly Clark",
            "Ewms",
            "All"});
            this.cmbBuildApp.Location = new System.Drawing.Point(109, 22);
            this.cmbBuildApp.Name = "cmbBuildApp";
            this.cmbBuildApp.Size = new System.Drawing.Size(121, 21);
            this.cmbBuildApp.TabIndex = 7;
            this.cmbBuildApp.Text = "Etricc UI";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Build Application:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cmbBranch);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.cmbBuildApp);
            this.groupBox1.Location = new System.Drawing.Point(27, 82);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(250, 224);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Build Info";
            // 
            // cmbBranch
            // 
            this.cmbBranch.DisplayMember = "Etricc";
            this.cmbBranch.FormattingEnabled = true;
            this.cmbBranch.Items.AddRange(new object[] {
            "Main",
            "Dev01",
            "Dev02",
            "Dev03",
            "Dev04",
            "Dev07",
            "Dev08",
            "AllBranchs"});
            this.cmbBranch.Location = new System.Drawing.Point(109, 53);
            this.cmbBranch.Name = "cmbBranch";
            this.cmbBranch.Size = new System.Drawing.Size(77, 21);
            this.cmbBranch.TabIndex = 16;
            this.cmbBranch.Text = "Main";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(50, 56);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(44, 13);
            this.label5.TabIndex = 15;
            this.label5.Text = "Branch:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ckbVersion);
            this.groupBox2.Controls.Add(this.ckbNightly);
            this.groupBox2.Controls.Add(this.ckbCI);
            this.groupBox2.Controls.Add(this.ckbWeekly);
            this.groupBox2.Location = new System.Drawing.Point(19, 87);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(145, 131);
            this.groupBox2.TabIndex = 14;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Build Definition";
            // 
            // ckbVersion
            // 
            this.ckbVersion.AutoSize = true;
            this.ckbVersion.Location = new System.Drawing.Point(47, 101);
            this.ckbVersion.Name = "ckbVersion";
            this.ckbVersion.Size = new System.Drawing.Size(61, 17);
            this.ckbVersion.TabIndex = 27;
            this.ckbVersion.Text = "Version";
            this.ckbVersion.UseVisualStyleBackColor = true;
            // 
            // ckbNightly
            // 
            this.ckbNightly.AutoSize = true;
            this.ckbNightly.Checked = true;
            this.ckbNightly.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ckbNightly.Location = new System.Drawing.Point(47, 55);
            this.ckbNightly.Name = "ckbNightly";
            this.ckbNightly.Size = new System.Drawing.Size(58, 17);
            this.ckbNightly.TabIndex = 26;
            this.ckbNightly.Text = "Nightly";
            this.ckbNightly.UseVisualStyleBackColor = true;
            // 
            // ckbCI
            // 
            this.ckbCI.AutoSize = true;
            this.ckbCI.Location = new System.Drawing.Point(47, 32);
            this.ckbCI.Name = "ckbCI";
            this.ckbCI.Size = new System.Drawing.Size(36, 17);
            this.ckbCI.TabIndex = 25;
            this.ckbCI.Text = "CI";
            this.ckbCI.UseVisualStyleBackColor = true;
            // 
            // ckbWeekly
            // 
            this.ckbWeekly.AutoSize = true;
            this.ckbWeekly.Location = new System.Drawing.Point(47, 78);
            this.ckbWeekly.Name = "ckbWeekly";
            this.ckbWeekly.Size = new System.Drawing.Size(62, 17);
            this.ckbWeekly.TabIndex = 24;
            this.ckbWeekly.Text = "Weekly";
            this.ckbWeekly.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(418, 146);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(36, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "Excel:";
            // 
            // cmbExcelVisible
            // 
            this.cmbExcelVisible.DisplayMember = "Visible";
            this.cmbExcelVisible.FormattingEnabled = true;
            this.cmbExcelVisible.Items.AddRange(new object[] {
            "Visible",
            "Invisible"});
            this.cmbExcelVisible.Location = new System.Drawing.Point(461, 143);
            this.cmbExcelVisible.Name = "cmbExcelVisible";
            this.cmbExcelVisible.Size = new System.Drawing.Size(121, 21);
            this.cmbExcelVisible.TabIndex = 15;
            this.cmbExcelVisible.Text = "Invisible";
            // 
            // button_DeploymentLocation
            // 
            this.button_DeploymentLocation.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_DeploymentLocation.Enabled = false;
            this.button_DeploymentLocation.Location = new System.Drawing.Point(551, 44);
            this.button_DeploymentLocation.Name = "button_DeploymentLocation";
            this.button_DeploymentLocation.Size = new System.Drawing.Size(28, 20);
            this.button_DeploymentLocation.TabIndex = 18;
            this.button_DeploymentLocation.Text = "...";
            this.button_DeploymentLocation.UseVisualStyleBackColor = true;
            this.button_DeploymentLocation.Click += new System.EventHandler(this.button_DeploymentLocation_Click);
            // 
            // folderBrowserDialog
            // 
            this.folderBrowserDialog.Description = "Select Deployment Location";
            // 
            // cmbProjectFile
            // 
            this.cmbProjectFile.DisplayMember = "Demo.xml";
            this.cmbProjectFile.FormattingEnabled = true;
            this.cmbProjectFile.Items.AddRange(new object[] {
            "Demo.xml",
            "Eurobaltic.zip",
            "TestProject.zip"});
            this.cmbProjectFile.Location = new System.Drawing.Point(156, 9);
            this.cmbProjectFile.Name = "cmbProjectFile";
            this.cmbProjectFile.Size = new System.Drawing.Size(423, 21);
            this.cmbProjectFile.TabIndex = 19;
            this.cmbProjectFile.Text = "Demo.xml";
            // 
            // ckbFunctionalTesting
            // 
            this.ckbFunctionalTesting.AutoSize = true;
            this.ckbFunctionalTesting.Location = new System.Drawing.Point(438, 93);
            this.ckbFunctionalTesting.Name = "ckbFunctionalTesting";
            this.ckbFunctionalTesting.Size = new System.Drawing.Size(141, 17);
            this.ckbFunctionalTesting.TabIndex = 20;
            this.ckbFunctionalTesting.Text = "Allow Functional Testing";
            this.ckbFunctionalTesting.UseVisualStyleBackColor = true;
            // 
            // cmbServerRunAs
            // 
            this.cmbServerRunAs.DisplayMember = "Service";
            this.cmbServerRunAs.FormattingEnabled = true;
            this.cmbServerRunAs.Items.AddRange(new object[] {
            "Service",
            "Console"});
            this.cmbServerRunAs.Location = new System.Drawing.Point(461, 116);
            this.cmbServerRunAs.Name = "cmbServerRunAs";
            this.cmbServerRunAs.Size = new System.Drawing.Size(121, 21);
            this.cmbServerRunAs.TabIndex = 21;
            this.cmbServerRunAs.Text = "Service";
            // 
            // labServer
            // 
            this.labServer.AutoSize = true;
            this.labServer.Location = new System.Drawing.Point(376, 119);
            this.labServer.Name = "labServer";
            this.labServer.Size = new System.Drawing.Size(79, 13);
            this.labServer.TabIndex = 22;
            this.labServer.Text = "Server Run As:";
            // 
            // ckbMail
            // 
            this.ckbMail.AutoSize = true;
            this.ckbMail.Location = new System.Drawing.Point(523, 203);
            this.ckbMail.Name = "ckbMail";
            this.ckbMail.Size = new System.Drawing.Size(45, 17);
            this.ckbMail.TabIndex = 24;
            this.ckbMail.Text = "Mail";
            this.ckbMail.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(376, 174);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(78, 13);
            this.label4.TabIndex = 25;
            this.label4.Text = "Platform target:";
            // 
            // cmbPlatformTarget
            // 
            this.cmbPlatformTarget.DisplayMember = "Visible";
            this.cmbPlatformTarget.FormattingEnabled = true;
            this.cmbPlatformTarget.Items.AddRange(new object[] {
            "Any CPU",
            "x86",
            "x64"});
            this.cmbPlatformTarget.Location = new System.Drawing.Point(461, 171);
            this.cmbPlatformTarget.Name = "cmbPlatformTarget";
            this.cmbPlatformTarget.Size = new System.Drawing.Size(121, 21);
            this.cmbPlatformTarget.TabIndex = 26;
            this.cmbPlatformTarget.Text = "Any CPU";
            // 
            // Configuration
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(598, 322);
            this.Controls.Add(this.cmbPlatformTarget);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.ckbMail);
            this.Controls.Add(this.labServer);
            this.Controls.Add(this.cmbServerRunAs);
            this.Controls.Add(this.ckbFunctionalTesting);
            this.Controls.Add(this.cmbProjectFile);
            this.Controls.Add(this.button_DeploymentLocation);
            this.Controls.Add(this.cmbExcelVisible);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSaveConfig);
            this.Controls.Add(this.txtBoxDeployLocation);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Configuration";
            this.Text = "E\'pia 3  Auto-Deployment Tool Configuration";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtBoxDeployLocation;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button btnSaveConfig;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ComboBox cmbBuildApp;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbExcelVisible;
        private System.Windows.Forms.Button button_DeploymentLocation;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.ComboBox cmbProjectFile;
        private System.Windows.Forms.CheckBox ckbFunctionalTesting;
        private System.Windows.Forms.ComboBox cmbServerRunAs;
        private System.Windows.Forms.Label labServer;
        private System.Windows.Forms.CheckBox ckbMail;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox ckbVersion;
        private System.Windows.Forms.CheckBox ckbNightly;
        private System.Windows.Forms.CheckBox ckbCI;
        private System.Windows.Forms.CheckBox ckbWeekly;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbPlatformTarget;
        private System.Windows.Forms.ComboBox cmbBranch;
        private System.Windows.Forms.Label label5;
    }
}