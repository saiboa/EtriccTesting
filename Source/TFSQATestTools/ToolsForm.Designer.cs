namespace TFSQATestTools
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
            this.HostTestMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.computerInfoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnStartAuto = new System.Windows.Forms.Button();
            this.btnStopAuto = new System.Windows.Forms.Button();
            this.btnConfig = new System.Windows.Forms.Button();
            this.btnStartManual = new System.Windows.Forms.Button();
            this.gbAutomaticTesting = new System.Windows.Forms.GroupBox();
            this.lbSelectedConfigFile = new System.Windows.Forms.Label();
            this.btnConn = new System.Windows.Forms.Button();
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
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnConnTFS = new System.Windows.Forms.Button();
            this.lstBoxBuildDefinitions = new System.Windows.Forms.ListBox();
            this.VMSwitchModeTimer = new System.Windows.Forms.Timer(this.components);
            this.btnClear = new System.Windows.Forms.Button();
            this.cmbTestApp2 = new System.Windows.Forms.ComboBox();
            this.cmbTestApp3 = new System.Windows.Forms.ComboBox();
            this.cmbTestApp1 = new System.Windows.Forms.ComboBox();
            this.cmbTestApp1DefFileName = new System.Windows.Forms.ComboBox();
            this.cmbTestApp2DefFileName = new System.Windows.Forms.ComboBox();
            this.cmbTestApp3DefFileName = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.buildProjectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.gbAutomaticTesting.SuspendLayout();
            this.gbManualTesting.SuspendLayout();
            this.groupBox2.SuspendLayout();
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
            this.menuStrip1.Size = new System.Drawing.Size(1099, 24);
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
            this.mnuExit.Size = new System.Drawing.Size(93, 22);
            this.mnuExit.Text = "Exit";
            this.mnuExit.Click += new System.EventHandler(this.mnuExit_Click);
            // 
            // configToolStripMenuItem
            // 
            this.configToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.selectToolStripMenuItem,
            this.HostTestMenuItem,
            this.buildProjectToolStripMenuItem});
            this.configToolStripMenuItem.Name = "configToolStripMenuItem";
            this.configToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.configToolStripMenuItem.Text = "Extra";
            // 
            // selectToolStripMenuItem
            // 
            this.selectToolStripMenuItem.Name = "selectToolStripMenuItem";
            this.selectToolStripMenuItem.Size = new System.Drawing.Size(213, 22);
            this.selectToolStripMenuItem.Text = "Epia Release or HotFix Test";
            // 
            // HostTestMenuItem
            // 
            this.HostTestMenuItem.Name = "HostTestMenuItem";
            this.HostTestMenuItem.Size = new System.Drawing.Size(213, 22);
            this.HostTestMenuItem.Text = "HostTest";
            this.HostTestMenuItem.Click += new System.EventHandler(this.HostTestMenuItem_Click);
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
            // gbAutomaticTesting
            // 
            this.gbAutomaticTesting.Controls.Add(this.btnConfig);
            this.gbAutomaticTesting.Controls.Add(this.lbSelectedConfigFile);
            this.gbAutomaticTesting.Controls.Add(this.btnStopAuto);
            this.gbAutomaticTesting.Controls.Add(this.btnConn);
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
            // btnConn
            // 
            this.btnConn.Location = new System.Drawing.Point(589, 27);
            this.btnConn.Name = "btnConn";
            this.btnConn.Size = new System.Drawing.Size(135, 23);
            this.btnConn.TabIndex = 17;
            this.btnConn.Text = "Test TFS Connection";
            this.btnConn.UseVisualStyleBackColor = true;
            this.btnConn.Click += new System.EventHandler(this.btnConn_Click);
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
            this.button_buildSelector.Text = "Select Msi Folder";
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
            this.txtBuildPath.TextChanged += new System.EventHandler(this.txtBuildPath_TextChanged);
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
            this.txtResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.txtResult.AutoCompleteCustomSource.AddRange(new string[] {
            "Epia 4",
            "Etricc UI",
            "Etricc 5",
            "Ewms"});
            this.txtResult.Location = new System.Drawing.Point(12, 253);
            this.txtResult.Multiline = true;
            this.txtResult.Name = "txtResult";
            this.txtResult.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtResult.Size = new System.Drawing.Size(760, 300);
            this.txtResult.TabIndex = 13;
            // 
            // labelX
            // 
            this.labelX.AutoSize = true;
            this.labelX.Location = new System.Drawing.Point(807, 0);
            this.labelX.Name = "labelX";
            this.labelX.Size = new System.Drawing.Size(75, 13);
            this.labelX.TabIndex = 14;
            this.labelX.Text = "Start Up Time:";
            // 
            // lbStartTime
            // 
            this.lbStartTime.AutoSize = true;
            this.lbStartTime.Location = new System.Drawing.Point(888, 0);
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
            this.btnConnTFS.Location = new System.Drawing.Point(778, 197);
            this.btnConnTFS.Name = "btnConnTFS";
            this.btnConnTFS.Size = new System.Drawing.Size(115, 23);
            this.btnConnTFS.TabIndex = 18;
            this.btnConnTFS.Text = "Get Build Definitions:";
            this.btnConnTFS.UseVisualStyleBackColor = true;
            this.btnConnTFS.Click += new System.EventHandler(this.btnConnTFS_Click);
            // 
            // lstBoxBuildDefinitions
            // 
            this.lstBoxBuildDefinitions.FormattingEnabled = true;
            this.lstBoxBuildDefinitions.Location = new System.Drawing.Point(778, 231);
            this.lstBoxBuildDefinitions.Name = "lstBoxBuildDefinitions";
            this.lstBoxBuildDefinitions.Size = new System.Drawing.Size(249, 225);
            this.lstBoxBuildDefinitions.TabIndex = 18;
            this.lstBoxBuildDefinitions.SelectedIndexChanged += new System.EventHandler(this.lstBoxBuildDefinitions_SelectedIndexChanged);
            // 
            // VMSwitchModeTimer
            // 
            this.VMSwitchModeTimer.Interval = 30000;
            this.VMSwitchModeTimer.Tick += new System.EventHandler(this.VMSwitchModeTimer_Tick);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(778, 530);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(71, 23);
            this.btnClear.TabIndex = 27;
            this.btnClear.Text = "Clear Log";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // cmbTestApp2
            // 
            this.cmbTestApp2.FormattingEnabled = true;
            this.cmbTestApp2.Items.AddRange(new object[] {
            "EtriccUI",
            "EtriccStatistics"});
            this.cmbTestApp2.Location = new System.Drawing.Point(11, 62);
            this.cmbTestApp2.Name = "cmbTestApp2";
            this.cmbTestApp2.Size = new System.Drawing.Size(80, 21);
            this.cmbTestApp2.TabIndex = 30;
            this.cmbTestApp2.Text = "EtriccUI";
            this.cmbTestApp2.SelectedIndexChanged += new System.EventHandler(this.cmbTestApp2_SelectedIndexChanged);
            // 
            // cmbTestApp3
            // 
            this.cmbTestApp3.FormattingEnabled = true;
            this.cmbTestApp3.Items.AddRange(new object[] {
            "EtriccStatistics"});
            this.cmbTestApp3.Location = new System.Drawing.Point(11, 89);
            this.cmbTestApp3.Name = "cmbTestApp3";
            this.cmbTestApp3.Size = new System.Drawing.Size(80, 21);
            this.cmbTestApp3.TabIndex = 31;
            this.cmbTestApp3.Text = "EtriccStatistics";
            this.cmbTestApp3.SelectedIndexChanged += new System.EventHandler(this.cmbTestApp3_SelectedIndexChanged);
            // 
            // cmbTestApp1
            // 
            this.cmbTestApp1.FormattingEnabled = true;
            this.cmbTestApp1.Items.AddRange(new object[] {
            "Epia4",
            "EtriccUI",
            "EtriccStatistics",
            "EpiaNet45",
            "EtriccNet45"});
            this.cmbTestApp1.Location = new System.Drawing.Point(11, 35);
            this.cmbTestApp1.Name = "cmbTestApp1";
            this.cmbTestApp1.Size = new System.Drawing.Size(80, 21);
            this.cmbTestApp1.TabIndex = 32;
            this.cmbTestApp1.Text = "Epia4";
            this.cmbTestApp1.SelectedIndexChanged += new System.EventHandler(this.cmbTestApp1_SelectedIndexChanged);
            // 
            // cmbTestApp1DefFileName
            // 
            this.cmbTestApp1DefFileName.FormattingEnabled = true;
            this.cmbTestApp1DefFileName.Location = new System.Drawing.Point(97, 35);
            this.cmbTestApp1DefFileName.Name = "cmbTestApp1DefFileName";
            this.cmbTestApp1DefFileName.Size = new System.Drawing.Size(199, 21);
            this.cmbTestApp1DefFileName.TabIndex = 33;
            this.cmbTestApp1DefFileName.Text = "Epia4TestDefinition";
            // 
            // cmbTestApp2DefFileName
            // 
            this.cmbTestApp2DefFileName.FormattingEnabled = true;
            this.cmbTestApp2DefFileName.Location = new System.Drawing.Point(97, 62);
            this.cmbTestApp2DefFileName.Name = "cmbTestApp2DefFileName";
            this.cmbTestApp2DefFileName.Size = new System.Drawing.Size(199, 21);
            this.cmbTestApp2DefFileName.TabIndex = 34;
            this.cmbTestApp2DefFileName.Text = "EtriccUITestTypeDefinition";
            // 
            // cmbTestApp3DefFileName
            // 
            this.cmbTestApp3DefFileName.FormattingEnabled = true;
            this.cmbTestApp3DefFileName.Location = new System.Drawing.Point(97, 89);
            this.cmbTestApp3DefFileName.Name = "cmbTestApp3DefFileName";
            this.cmbTestApp3DefFileName.Size = new System.Drawing.Size(199, 21);
            this.cmbTestApp3DefFileName.TabIndex = 35;
            this.cmbTestApp3DefFileName.Text = "StatisticsTestTypeDefinition";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.cmbTestApp3DefFileName);
            this.groupBox2.Controls.Add(this.cmbTestApp1);
            this.groupBox2.Controls.Add(this.cmbTestApp2);
            this.groupBox2.Controls.Add(this.cmbTestApp2DefFileName);
            this.groupBox2.Controls.Add(this.cmbTestApp3);
            this.groupBox2.Controls.Add(this.cmbTestApp1DefFileName);
            this.groupBox2.Location = new System.Drawing.Point(778, 36);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(314, 145);
            this.groupBox2.TabIndex = 28;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Test Applications Selection";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(8, 19);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(86, 13);
            this.label6.TabIndex = 37;
            this.label6.Text = "Test Application:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(103, 19);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(78, 13);
            this.label4.TabIndex = 36;
            this.label4.Text = "Test Definition:";
            // 
            // buildProjectToolStripMenuItem
            // 
            this.buildProjectToolStripMenuItem.Name = "buildProjectToolStripMenuItem";
            this.buildProjectToolStripMenuItem.Size = new System.Drawing.Size(213, 22);
            this.buildProjectToolStripMenuItem.Text = "BuildProject";
            this.buildProjectToolStripMenuItem.Click += new System.EventHandler(this.buildProjectToolStripMenuItem_Click);
            // 
            // ToolsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1099, 565);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.lbStartTime);
            this.Controls.Add(this.labelX);
            this.Controls.Add(this.txtResult);
            this.Controls.Add(this.btnConnTFS);
            this.Controls.Add(this.gbManualTesting);
            this.Controls.Add(this.gbAutomaticTesting);
            this.Controls.Add(this.lstBoxBuildDefinitions);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "ToolsForm";
            this.Text = "E\'pia QA Test Tool";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.gbAutomaticTesting.ResumeLayout(false);
            this.gbAutomaticTesting.PerformLayout();
            this.gbManualTesting.ResumeLayout(false);
            this.gbManualTesting.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
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
        private System.Windows.Forms.ToolStripMenuItem HostTestMenuItem;
        private System.Windows.Forms.Button btnConnTFS;
        private System.Windows.Forms.ListBox lstBoxBuildDefinitions;
        private System.Windows.Forms.ToolStripMenuItem computerInfoToolStripMenuItem;
        private System.Windows.Forms.Timer VMSwitchModeTimer;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.ComboBox cmbTestApp2;
        private System.Windows.Forms.ComboBox cmbTestApp3;
        private System.Windows.Forms.ComboBox cmbTestApp1;
        private System.Windows.Forms.ComboBox cmbTestApp1DefFileName;
        private System.Windows.Forms.ComboBox cmbTestApp2DefFileName;
        private System.Windows.Forms.ComboBox cmbTestApp3DefFileName;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ToolStripMenuItem buildProjectToolStripMenuItem;
    }
}

