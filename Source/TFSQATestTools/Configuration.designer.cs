namespace TFSQATestTools
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtBoxDeployLocation = new System.Windows.Forms.TextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnSaveConfig = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbExcelVisible = new System.Windows.Forms.ComboBox();
            this.button_DeploymentLocation = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.cmbProjectFile = new System.Windows.Forms.ComboBox();
            this.ckbFunctionalTesting = new System.Windows.Forms.CheckBox();
            this.cmbServerRunAs = new System.Windows.Forms.ComboBox();
            this.labServer = new System.Windows.Forms.Label();
            this.ckbMail = new System.Windows.Forms.CheckBox();
            this.ckbRemoteVMSwitchMode = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 11);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(104, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Project XMLFile:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 55);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(138, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Deployment Location:";
            // 
            // txtBoxDeployLocation
            // 
            this.txtBoxDeployLocation.Location = new System.Drawing.Point(208, 55);
            this.txtBoxDeployLocation.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtBoxDeployLocation.Name = "txtBoxDeployLocation";
            this.txtBoxDeployLocation.Size = new System.Drawing.Size(532, 22);
            this.txtBoxDeployLocation.TabIndex = 4;
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            // 
            // btnSaveConfig
            // 
            this.btnSaveConfig.Location = new System.Drawing.Point(672, 278);
            this.btnSaveConfig.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnSaveConfig.Name = "btnSaveConfig";
            this.btnSaveConfig.Size = new System.Drawing.Size(100, 28);
            this.btnSaveConfig.TabIndex = 5;
            this.btnSaveConfig.Text = "Save";
            this.btnSaveConfig.UseVisualStyleBackColor = true;
            this.btnSaveConfig.Click += new System.EventHandler(this.btnSaveConfig_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(564, 278);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(100, 28);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(23, 206);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(44, 16);
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
            this.cmbExcelVisible.Location = new System.Drawing.Point(136, 206);
            this.cmbExcelVisible.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbExcelVisible.Name = "cmbExcelVisible";
            this.cmbExcelVisible.Size = new System.Drawing.Size(160, 24);
            this.cmbExcelVisible.TabIndex = 15;
            this.cmbExcelVisible.Text = "Invisible";
            // 
            // button_DeploymentLocation
            // 
            this.button_DeploymentLocation.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_DeploymentLocation.Location = new System.Drawing.Point(735, 54);
            this.button_DeploymentLocation.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button_DeploymentLocation.Name = "button_DeploymentLocation";
            this.button_DeploymentLocation.Size = new System.Drawing.Size(37, 25);
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
            this.cmbProjectFile.Location = new System.Drawing.Point(208, 11);
            this.cmbProjectFile.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbProjectFile.Name = "cmbProjectFile";
            this.cmbProjectFile.Size = new System.Drawing.Size(563, 24);
            this.cmbProjectFile.TabIndex = 19;
            this.cmbProjectFile.Text = "Demo.xml";
            // 
            // ckbFunctionalTesting
            // 
            this.ckbFunctionalTesting.AutoSize = true;
            this.ckbFunctionalTesting.Location = new System.Drawing.Point(20, 102);
            this.ckbFunctionalTesting.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ckbFunctionalTesting.Name = "ckbFunctionalTesting";
            this.ckbFunctionalTesting.Size = new System.Drawing.Size(174, 20);
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
            this.cmbServerRunAs.Location = new System.Drawing.Point(136, 172);
            this.cmbServerRunAs.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbServerRunAs.Name = "cmbServerRunAs";
            this.cmbServerRunAs.Size = new System.Drawing.Size(160, 24);
            this.cmbServerRunAs.TabIndex = 21;
            this.cmbServerRunAs.Text = "Service";
            // 
            // labServer
            // 
            this.labServer.AutoSize = true;
            this.labServer.Location = new System.Drawing.Point(23, 176);
            this.labServer.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labServer.Name = "labServer";
            this.labServer.Size = new System.Drawing.Size(97, 16);
            this.labServer.TabIndex = 22;
            this.labServer.Text = "Server Run As:";
            // 
            // ckbMail
            // 
            this.ckbMail.AutoSize = true;
            this.ckbMail.Location = new System.Drawing.Point(20, 130);
            this.ckbMail.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ckbMail.Name = "ckbMail";
            this.ckbMail.Size = new System.Drawing.Size(55, 20);
            this.ckbMail.TabIndex = 24;
            this.ckbMail.Text = "Mail";
            this.ckbMail.UseVisualStyleBackColor = true;
            // 
            // ckbRemoteVMSwitchMode
            // 
            this.ckbRemoteVMSwitchMode.AutoSize = true;
            this.ckbRemoteVMSwitchMode.Location = new System.Drawing.Point(27, 258);
            this.ckbRemoteVMSwitchMode.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ckbRemoteVMSwitchMode.Name = "ckbRemoteVMSwitchMode";
            this.ckbRemoteVMSwitchMode.Size = new System.Drawing.Size(129, 20);
            this.ckbRemoteVMSwitchMode.TabIndex = 25;
            this.ckbRemoteVMSwitchMode.Text = "VM Switch Mode";
            this.ckbRemoteVMSwitchMode.UseVisualStyleBackColor = true;
            // 
            // Configuration
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(797, 396);
            this.Controls.Add(this.ckbRemoteVMSwitchMode);
            this.Controls.Add(this.ckbMail);
            this.Controls.Add(this.labServer);
            this.Controls.Add(this.cmbServerRunAs);
            this.Controls.Add(this.ckbFunctionalTesting);
            this.Controls.Add(this.cmbProjectFile);
            this.Controls.Add(this.button_DeploymentLocation);
            this.Controls.Add(this.cmbExcelVisible);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSaveConfig);
            this.Controls.Add(this.txtBoxDeployLocation);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Configuration";
            this.Text = "E\'pia 3  Auto-Deployment Tool Configuration";
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
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbExcelVisible;
        private System.Windows.Forms.Button button_DeploymentLocation;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.ComboBox cmbProjectFile;
        private System.Windows.Forms.CheckBox ckbFunctionalTesting;
        private System.Windows.Forms.ComboBox cmbServerRunAs;
        private System.Windows.Forms.Label labServer;
        private System.Windows.Forms.CheckBox ckbMail;
        private System.Windows.Forms.CheckBox ckbRemoteVMSwitchMode;
    }
}