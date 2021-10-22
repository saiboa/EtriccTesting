namespace TFS2010AutoDeploymentTool
{
    partial class GetBuildDefinitionsForm
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
            this.btnConn = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lstBoxProject = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbTFSServer = new System.Windows.Forms.ComboBox();
            this.lstBoxDuildDefinition = new System.Windows.Forms.ListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbDateFilter = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.listBoxTestApp = new System.Windows.Forms.ListBox();
            this.chkBuildDefs = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnConn
            // 
            this.btnConn.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnConn.Location = new System.Drawing.Point(245, 392);
            this.btnConn.Name = "btnConn";
            this.btnConn.Size = new System.Drawing.Size(75, 23);
            this.btnConn.TabIndex = 0;
            this.btnConn.Text = "OK";
            this.btnConn.UseVisualStyleBackColor = true;
            this.btnConn.Click += new System.EventHandler(this.btnConn_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(338, 392);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(127, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Team Foundation Server:";
            // 
            // lstBoxProject
            // 
            this.lstBoxProject.FormattingEnabled = true;
            this.lstBoxProject.Location = new System.Drawing.Point(113, 49);
            this.lstBoxProject.Name = "lstBoxProject";
            this.lstBoxProject.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.lstBoxProject.Size = new System.Drawing.Size(314, 17);
            this.lstBoxProject.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Team Projects:";
            // 
            // cmbTFSServer
            // 
            this.cmbTFSServer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTFSServer.FormattingEnabled = true;
            this.cmbTFSServer.ImeMode = System.Windows.Forms.ImeMode.Disable;
            this.cmbTFSServer.Location = new System.Drawing.Point(155, 12);
            this.cmbTFSServer.Name = "cmbTFSServer";
            this.cmbTFSServer.Size = new System.Drawing.Size(272, 21);
            this.cmbTFSServer.TabIndex = 5;
            // 
            // lstBoxDuildDefinition
            // 
            this.lstBoxDuildDefinition.FormattingEnabled = true;
            this.lstBoxDuildDefinition.Location = new System.Drawing.Point(32, 113);
            this.lstBoxDuildDefinition.Name = "lstBoxDuildDefinition";
            this.lstBoxDuildDefinition.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lstBoxDuildDefinition.Size = new System.Drawing.Size(395, 199);
            this.lstBoxDuildDefinition.TabIndex = 7;
            this.lstBoxDuildDefinition.SelectedIndexChanged += new System.EventHandler(this.lstBoxDuildDefinition_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(29, 334);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Date filter:";
            // 
            // cmbDateFilter
            // 
            this.cmbDateFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDateFilter.FormattingEnabled = true;
            this.cmbDateFilter.Items.AddRange(new object[] {
            "Today",
            "Last 24 hours",
            "Last 48 hours",
            "Last 7 days",
            "Last 14 days",
            "Last 28 days",
            "<Any Time>"});
            this.cmbDateFilter.Location = new System.Drawing.Point(90, 331);
            this.cmbDateFilter.Name = "cmbDateFilter";
            this.cmbDateFilter.Size = new System.Drawing.Size(272, 21);
            this.cmbDateFilter.TabIndex = 9;
            this.cmbDateFilter.SelectedIndexChanged += new System.EventHandler(this.cmbDateFilter_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(29, 75);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(86, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Test Application:";
            // 
            // listBoxTestApp
            // 
            this.listBoxTestApp.FormattingEnabled = true;
            this.listBoxTestApp.Location = new System.Drawing.Point(113, 75);
            this.listBoxTestApp.Name = "listBoxTestApp";
            this.listBoxTestApp.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.listBoxTestApp.Size = new System.Drawing.Size(314, 17);
            this.listBoxTestApp.TabIndex = 11;
            // 
            // chkBuildDefs
            // 
            this.chkBuildDefs.AutoSize = true;
            this.chkBuildDefs.Location = new System.Drawing.Point(32, 96);
            this.chkBuildDefs.Name = "chkBuildDefs";
            this.chkBuildDefs.Size = new System.Drawing.Size(104, 17);
            this.chkBuildDefs.TabIndex = 12;
            this.chkBuildDefs.Text = "Build Definitions:";
            this.chkBuildDefs.UseVisualStyleBackColor = true;
            this.chkBuildDefs.CheckedChanged += new System.EventHandler(this.chkBuildDefs_CheckedChanged);
            // 
            // GetBuildDefinitionsForm
            // 
            this.AcceptButton = this.btnConn;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(449, 441);
            this.Controls.Add(this.chkBuildDefs);
            this.Controls.Add(this.listBoxTestApp);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cmbDateFilter);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lstBoxDuildDefinition);
            this.Controls.Add(this.cmbTFSServer);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lstBoxProject);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnConn);
            this.Name = "GetBuildDefinitionsForm";
            this.Text = "Connect to Team Project";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnConn;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstBoxProject;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbTFSServer;
        private System.Windows.Forms.ListBox lstBoxDuildDefinition;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbDateFilter;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ListBox listBoxTestApp;
        private System.Windows.Forms.CheckBox chkBuildDefs;
    }
}