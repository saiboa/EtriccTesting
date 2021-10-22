namespace TFS2010AutoDeploymentTool
{
    partial class ToolsForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ToolsForm));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.mnuFile = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuExit = new System.Windows.Forms.ToolStripMenuItem();
            this.configToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.selectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.computerInfoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnStartAuto = new System.Windows.Forms.Button();
            this.btnStopAuto = new System.Windows.Forms.Button();
            this.btnConfig = new System.Windows.Forms.Button();
            this.btnStartManual = new System.Windows.Forms.Button();
            this.btnClearLog = new System.Windows.Forms.Button();
            this.gbAutomaticTesting = new System.Windows.Forms.GroupBox();
            this.lbSelectedConfigFile = new System.Windows.Forms.Label();
            this.lbCurrentBuildInTesting = new System.Windows.Forms.Label();
            this.gbManualTesting = new System.Windows.Forms.GroupBox();
            this.button_buildSelector = new System.Windows.Forms.Button();
            this.labelY = new System.Windows.Forms.Label();
            this.txtBuildPath = new System.Windows.Forms.TextBox();
            this.chkContinueAuto = new System.Windows.Forms.CheckBox();
            this.txtResult = new System.Windows.Forms.TextBox();
            this.labelX = new System.Windows.Forms.Label();
            this.lbStartTime = new System.Windows.Forms.Label();
            this.timStart = new System.Windows.Forms.Timer(this.components);
            this.btnConn = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnConnTFS = new System.Windows.Forms.Button();
            this.cmbProject = new System.Windows.Forms.ComboBox();
            this.lstBoxBuildDefinitions = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkProtected = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbTargetPlatform = new System.Windows.Forms.ComboBox();
            this.cmbTestApp = new System.Windows.Forms.ComboBox();
            this.VMSwitchModeTimer = new System.Windows.Forms.Timer(this.components);
            this.menuStrip1.SuspendLayout();
            this.gbAutomaticTesting.SuspendLayout();
            this.gbManualTesting.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuFile,
            this.configToolStripMenuItem,
            this.mnuHelp});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1092, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // mnuFile
            // 
            this.mnuFile.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuExit});
            this.mnuFile.Name = "mnuFile";
            this.mnuFile.Size = new System.Drawing.Size(37, 20);
            this.mnuFile.Text = "File";
            // 
            // mnuExit
            // 
            this.mnuExit.Name = "mnuExit";
            this.mnuExit.Size = new System.Drawing.Size(92, 22);
            this.mnuExit.Text = "Exit";
            this.mnuExit.Click += new System.EventHandler(this.mnuExit_Click);
            // 
            // configToolStripMenuItem
            // 
            this.configToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.selectToolStripMenuItem,
            this.viewToolStripMenuItem});
            this.configToolStripMenuItem.Name = "configToolStripMenuItem";
            this.configToolStripMenuItem.Size = new System.Drawing.Size(55, 20);
            this.configToolStripMenuItem.Text = "Config";
            // 
            // selectToolStripMenuItem
            // 
            this.selectToolStripMenuItem.Name = "selectToolStripMenuItem";
            this.selectToolStripMenuItem.Size = new System.Drawing.Size(105, 22);
            this.selectToolStripMenuItem.Text = "Select";
            // 
            // viewToolStripMenuItem
            // 
            this.viewToolStripMenuItem.Name = "viewToolStripMenuItem";
            this.viewToolStripMenuItem.Size = new System.Drawing.Size(105, 22);
            this.viewToolStripMenuItem.Text = "View";
            // 
            // mnuHelp
            // 
            this.mnuHelp.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuAbout,
            this.computerInfoToolStripMenuItem});
            this.mnuHelp.Name = "mnuHelp";
            this.mnuHelp.Size = new System.Drawing.Size(44, 20);
            this.mnuHelp.Text = "Help";
            // 
            // mnuAbout
            // 
            this.mnuAbout.Name = "mnuAbout";
            this.mnuAbout.Size = new System.Drawing.Size(152, 22);
            this.mnuAbout.Text = "About";
            this.mnuAbout.Click += new System.EventHandler(this.mnuAbout_Click);
            // 
            // computerInfoToolStripMenuItem
            // 
            this.computerInfoToolStripMenuItem.Name = "computerInfoToolStripMenuItem";
            this.computerInfoToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.computerInfoToolStripMenuItem.Text = "Computer Info";
            this.computerInfoToolStripMenuItem.Click += new System.EventHandler(this.computerInfoToolStripMenuItem_Click);
            // 
            // btnStartAuto
            // 
            this.btnStartAuto.Location = new System.Drawing.Point(13, 26);
            this.btnStartAuto.Name = "btnStartAuto";
            this.btnStartAuto.Size = new System.Drawing.Size(75, 23);
            this.btnStartAuto.TabIndex = 1;
            this.btnStartAuto.Text = "Start";
            this.btnStartAuto.UseVisualStyleBackColor = true;
            this.btnStartAuto.Click += new System.EventHandler(this.btnStartAuto_Click);
            // 
            // btnStopAuto
            // 
            this.btnStopAuto.Location = new System.Drawing.Point(107, 27);
            this.btnStopAuto.Name = "btnStopAuto";
            this.btnStopAuto.Size = new System.Drawing.Size(75, 23);
            this.btnStopAuto.TabIndex = 2;
            this.btnStopAuto.Text = "Stop";
            this.btnStopAuto.UseVisualStyleBackColor = true;
            this.btnStopAuto.Click += new System.EventHandler(this.btnStopAuto_Click);
            // 
            // btnConfig
            // 
            this.btnConfig.Location = new System.Drawing.Point(13, 57);
            this.btnConfig.Name = "btnConfig";
            this.btnConfig.Size = new System.Drawing.Size(75, 23);
            this.btnConfig.TabIndex = 3;
            this.btnConfig.Text = "Config";
            this.btnConfig.UseVisualStyleBackColor = true;
            this.btnConfig.Click += new System.EventHandler(this.btnConfig_Click);
            // 
            // btnStartManual
            // 
            this.btnStartManual.Location = new System.Drawing.Point(6, 56);
            this.btnStartManual.Name = "btnStartManual";
            this.btnStartManual.Size = new System.Drawing.Size(75, 23);
            this.btnStartManual.TabIndex = 4;
            this.btnStartManual.Text = "Start";
            this.btnStartManual.UseVisualStyleBackColor = true;
            this.btnStartManual.Click += new System.EventHandler(this.btnStartManual_Click);
            // 
            // btnClearLog
            // 
            this.btnClearLog.Location = new System.Drawing.Point(340, 507);
            this.btnClearLog.Name = "btnClearLog";
            this.btnClearLog.Size = new System.Drawing.Size(75, 23);
            this.btnClearLog.TabIndex = 5;
            this.btnClearLog.Text = "Clear Log";
            this.btnClearLog.UseVisualStyleBackColor = true;
            // 
            // gbAutomaticTesting
            // 
            this.gbAutomaticTesting.Controls.Add(this.btnConfig);
            this.gbAutomaticTesting.Controls.Add(this.lbSelectedConfigFile);
            this.gbAutomaticTesting.Controls.Add(this.btnStopAuto);
            this.gbAutomaticTesting.Controls.Add(this.btnStartAuto);
            this.gbAutomaticTesting.Controls.Add(this.lbCurrentBuildInTesting);
            this.gbAutomaticTesting.Location = new System.Drawing.Point(12, 36);
            this.gbAutomaticTesting.Name = "gbAutomaticTesting";
            this.gbAutomaticTesting.Size = new System.Drawing.Size(748, 99);
            this.gbAutomaticTesting.TabIndex = 9;
            this.gbAutomaticTesting.TabStop = false;
            this.gbAutomaticTesting.Text = "Automatic deployment";
            // 
            // lbSelectedConfigFile
            // 
            this.lbSelectedConfigFile.AutoSize = true;
            this.lbSelectedConfigFile.Location = new System.Drawing.Point(233, 67);
            this.lbSelectedConfigFile.Name = "lbSelectedConfigFile";
            this.lbSelectedConfigFile.Size = new System.Drawing.Size(127, 13);
            this.lbSelectedConfigFile.TabIndex = 16;
            this.lbSelectedConfigFile.Text = "searching build for testing";
            this.lbSelectedConfigFile.UseMnemonic = false;
            // 
            // lbCurrentBuildInTesting
            // 
            this.lbCurrentBuildInTesting.AutoSize = true;
            this.lbCurrentBuildInTesting.Location = new System.Drawing.Point(113, 67);
            this.lbCurrentBuildInTesting.Name = "lbCurrentBuildInTesting";
            this.lbCurrentBuildInTesting.Size = new System.Drawing.Size(114, 13);
            this.lbCurrentBuildInTesting.TabIndex = 11;
            this.lbCurrentBuildInTesting.Text = "Current build in testing:";
            // 
            // gbManualTesting
            // 
            this.gbManualTesting.Controls.Add(this.button_buildSelector);
            this.gbManualTesting.Controls.Add(this.labelY);
            this.gbManualTesting.Controls.Add(this.txtBuildPath);
            this.gbManualTesting.Controls.Add(this.chkContinueAuto);
            this.gbManualTesting.Controls.Add(this.btnStartManual);
            this.gbManualTesting.Location = new System.Drawing.Point(12, 141);
            this.gbManualTesting.Name = "gbManualTesting";
            this.gbManualTesting.Size = new System.Drawing.Size(748, 91);
            this.gbManualTesting.TabIndex = 10;
            this.gbManualTesting.TabStop = false;
            this.gbManualTesting.Text = "Manual deployment";
            // 
            // button_buildSelector
            // 
            this.button_buildSelector.Location = new System.Drawing.Point(617, 23);
            this.button_buildSelector.Name = "button_buildSelector";
            this.button_buildSelector.Size = new System.Drawing.Size(125, 20);
            this.button_buildSelector.TabIndex = 16;
            this.button_buildSelector.Text = "Get Builds";
            this.button_buildSelector.UseVisualStyleBackColor = true;
            this.button_buildSelector.Click += new System.EventHandler(this.button_buildSelector_Click);
            // 
            // labelY
            // 
            this.labelY.AutoSize = true;
            this.labelY.Location = new System.Drawing.Point(6, 27);
            this.labelY.Name = "labelY";
            this.labelY.Size = new System.Drawing.Size(105, 13);
            this.labelY.TabIndex = 15;
            this.labelY.Text = "Install selected build:";
            // 
            // txtBuildPath
            // 
            this.txtBuildPath.Location = new System.Drawing.Point(116, 24);
            this.txtBuildPath.Name = "txtBuildPath";
            this.txtBuildPath.Size = new System.Drawing.Size(503, 20);
            this.txtBuildPath.TabIndex = 14;
            // 
            // chkContinueAuto
            // 
            this.chkContinueAuto.AutoSize = true;
            this.chkContinueAuto.Checked = true;
            this.chkContinueAuto.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkContinueAuto.Location = new System.Drawing.Point(87, 62);
            this.chkContinueAuto.Name = "chkContinueAuto";
            this.chkContinueAuto.Size = new System.Drawing.Size(117, 17);
            this.chkContinueAuto.TabIndex = 12;
            this.chkContinueAuto.Text = "Continue automatic";
            this.chkContinueAuto.UseVisualStyleBackColor = true;
            // 
            // txtResult
            // 
            this.txtResult.AutoCompleteCustomSource.AddRange(new string[] {
            "Epia 4",
            "Etricc UI",
            "Etricc 5",
            "Ewms"});
            this.txtResult.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.txtResult.Location = new System.Drawing.Point(0, 267);
            this.txtResult.Multiline = true;
            this.txtResult.Name = "txtResult";
            this.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtResult.Size = new System.Drawing.Size(1092, 286);
            this.txtResult.TabIndex = 13;
            // 
            // labelX
            // 
            this.labelX.AutoSize = true;
            this.labelX.Location = new System.Drawing.Point(15, 243);
            this.labelX.Name = "labelX";
            this.labelX.Size = new System.Drawing.Size(75, 13);
            this.labelX.TabIndex = 14;
            this.labelX.Text = "Start Up Time:";
            // 
            // lbStartTime
            // 
            this.lbStartTime.AutoSize = true;
            this.lbStartTime.Location = new System.Drawing.Point(96, 243);
            this.lbStartTime.Name = "lbStartTime";
            this.lbStartTime.Size = new System.Drawing.Size(118, 13);
            this.lbStartTime.TabIndex = 15;
            this.lbStartTime.Text = "2008/May/01-09:09:09";
            // 
            // timStart
            // 
            this.timStart.Interval = 30000;
            this.timStart.Tick += new System.EventHandler(this.timStart_Tick);
            // 
            // btnConn
            // 
            this.btnConn.Location = new System.Drawing.Point(619, 238);
            this.btnConn.Name = "btnConn";
            this.btnConn.Size = new System.Drawing.Size(135, 23);
            this.btnConn.TabIndex = 17;
            this.btnConn.Text = "Test TFS Connection";
            this.btnConn.UseVisualStyleBackColor = true;
            this.btnConn.Click += new System.EventHandler(this.btnConn_Click);
            // 
            // folderBrowserDialog
            // 
            this.folderBrowserDialog.Description = "Select Install Scripts Folder";
            this.folderBrowserDialog.ShowNewFolderButton = false;
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            // 
            // btnConnTFS
            // 
            this.btnConnTFS.Location = new System.Drawing.Point(6, 92);
            this.btnConnTFS.Name = "btnConnTFS";
            this.btnConnTFS.Size = new System.Drawing.Size(136, 23);
            this.btnConnTFS.TabIndex = 18;
            this.btnConnTFS.Text = "Get Build Definitions:";
            this.btnConnTFS.UseVisualStyleBackColor = true;
            this.btnConnTFS.Click += new System.EventHandler(this.btnConnTFS_Click);
            // 
            // cmbProject
            // 
            this.cmbProject.FormattingEnabled = true;
            this.cmbProject.Items.AddRange(new object[] {
            "Epia 4",
            "Etricc 5",
            "Epia 3"});
            this.cmbProject.Location = new System.Drawing.Point(162, 11);
            this.cmbProject.Name = "cmbProject";
            this.cmbProject.Size = new System.Drawing.Size(121, 21);
            this.cmbProject.TabIndex = 18;
            this.cmbProject.Text = "Epia 4";
            this.cmbProject.SelectedIndexChanged += new System.EventHandler(this.cmbProject_SelectedIndexChanged);
            // 
            // lstBoxBuildDefinitions
            // 
            this.lstBoxBuildDefinitions.FormattingEnabled = true;
            this.lstBoxBuildDefinitions.Location = new System.Drawing.Point(6, 121);
            this.lstBoxBuildDefinitions.Name = "lstBoxBuildDefinitions";
            this.lstBoxBuildDefinitions.Size = new System.Drawing.Size(277, 108);
            this.lstBoxBuildDefinitions.TabIndex = 18;
            this.lstBoxBuildDefinitions.SelectedIndexChanged += new System.EventHandler(this.lstBoxBuildDefinitions_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkProtected);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cmbTargetPlatform);
            this.groupBox1.Controls.Add(this.cmbTestApp);
            this.groupBox1.Controls.Add(this.btnConnTFS);
            this.groupBox1.Controls.Add(this.cmbProject);
            this.groupBox1.Controls.Add(this.lstBoxBuildDefinitions);
            this.groupBox1.Location = new System.Drawing.Point(777, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(289, 234);
            this.groupBox1.TabIndex = 19;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "TFS Settings";
            // 
            // chkProtected
            // 
            this.chkProtected.AutoSize = true;
            this.chkProtected.Checked = true;
            this.chkProtected.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkProtected.Location = new System.Drawing.Point(211, 96);
            this.chkProtected.Name = "chkProtected";
            this.chkProtected.Size = new System.Drawing.Size(72, 17);
            this.chkProtected.TabIndex = 17;
            this.chkProtected.Text = "Protected";
            this.chkProtected.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(113, 66);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 13);
            this.label3.TabIndex = 23;
            this.label3.Text = "Target Platform:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(113, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 22;
            this.label2.Text = "Test App:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(113, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 21;
            this.label1.Text = "Project:";
            // 
            // cmbTargetPlatform
            // 
            this.cmbTargetPlatform.FormattingEnabled = true;
            this.cmbTargetPlatform.Items.AddRange(new object[] {
            "AnyCPU",
            "x86",
            "AnyCPU+x86"});
            this.cmbTargetPlatform.Location = new System.Drawing.Point(201, 65);
            this.cmbTargetPlatform.Name = "cmbTargetPlatform";
            this.cmbTargetPlatform.Size = new System.Drawing.Size(82, 21);
            this.cmbTargetPlatform.TabIndex = 20;
            this.cmbTargetPlatform.Text = "AnyCPU";
            this.cmbTargetPlatform.SelectedIndexChanged += new System.EventHandler(this.cmbTargetPlatform_SelectedIndexChanged);
            // 
            // cmbTestApp
            // 
            this.cmbTestApp.FormattingEnabled = true;
            this.cmbTestApp.Items.AddRange(new object[] {
            "Epia4",
            "Etricc5",
            "EtriccUI",
            "EtriccStatistics",
            "Ewms"});
            this.cmbTestApp.Location = new System.Drawing.Point(173, 38);
            this.cmbTestApp.Name = "cmbTestApp";
            this.cmbTestApp.Size = new System.Drawing.Size(110, 21);
            this.cmbTestApp.TabIndex = 19;
            this.cmbTestApp.Text = "Epia4";
            this.cmbTestApp.SelectedIndexChanged += new System.EventHandler(this.cmbTestApp_SelectedIndexChanged);
            // 
            // VMSwitchModeTimer
            // 
            this.VMSwitchModeTimer.Interval = 30000;
            this.VMSwitchModeTimer.Tick += new System.EventHandler(this.VMSwitchModeTimer_Tick);
            // 
            // ToolsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1092, 553);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnConn);
            this.Controls.Add(this.lbStartTime);
            this.Controls.Add(this.labelX);
            this.Controls.Add(this.txtResult);
            this.Controls.Add(this.btnClearLog);
            this.Controls.Add(this.gbManualTesting);
            this.Controls.Add(this.gbAutomaticTesting);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "ToolsForm";
            this.Text = "E\'pia Auto-Deployment Tool";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.gbAutomaticTesting.ResumeLayout(false);
            this.gbAutomaticTesting.PerformLayout();
            this.gbManualTesting.ResumeLayout(false);
            this.gbManualTesting.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem mnuFile;
        private System.Windows.Forms.ToolStripMenuItem mnuExit;
        private System.Windows.Forms.ToolStripMenuItem configToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mnuHelp;
        private System.Windows.Forms.ToolStripMenuItem mnuAbout;
        private System.Windows.Forms.Button btnStartAuto;
        private System.Windows.Forms.Button btnStopAuto;
        private System.Windows.Forms.Button btnConfig;
        private System.Windows.Forms.Button btnStartManual;
        private System.Windows.Forms.Button btnClearLog;
        private System.Windows.Forms.GroupBox gbAutomaticTesting;
        private System.Windows.Forms.GroupBox gbManualTesting;
        private System.Windows.Forms.CheckBox chkContinueAuto;
        private System.Windows.Forms.Label lbCurrentBuildInTesting;
        private System.Windows.Forms.TextBox txtResult;
        private System.Windows.Forms.Label labelX;
        private System.Windows.Forms.Label lbStartTime;
        private System.Windows.Forms.Label labelY;
        private System.Windows.Forms.TextBox txtBuildPath;
        private System.Windows.Forms.Timer timStart;
        private System.Windows.Forms.Label lbSelectedConfigFile;
        private System.Windows.Forms.Button btnConn;
        private System.Windows.Forms.Button button_buildSelector;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.ToolStripMenuItem selectToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewToolStripMenuItem;
        private System.Windows.Forms.Button btnConnTFS;
        private System.Windows.Forms.ComboBox cmbProject;
        private System.Windows.Forms.ListBox lstBoxBuildDefinitions;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbTargetPlatform;
        private System.Windows.Forms.ComboBox cmbTestApp;
        private System.Windows.Forms.ToolStripMenuItem computerInfoToolStripMenuItem;
        private System.Windows.Forms.CheckBox chkProtected;
        private System.Windows.Forms.Timer VMSwitchModeTimer;

       
    }
}

