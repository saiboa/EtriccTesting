namespace TFSQATestTools
{
    partial class GetBuildNumbersForm
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
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lstBoxProject = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbTFSServer = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lstBoxDuildNumber = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(245, 392);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnConn_Click);
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
            this.lstBoxProject.AllowDrop = true;
            this.lstBoxProject.FormattingEnabled = true;
            this.lstBoxProject.Location = new System.Drawing.Point(32, 79);
            this.lstBoxProject.MultiColumn = true;
            this.lstBoxProject.Name = "lstBoxProject";
            this.lstBoxProject.ScrollAlwaysVisible = true;
            this.lstBoxProject.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.lstBoxProject.Size = new System.Drawing.Size(381, 17);
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
            this.cmbTFSServer.FormattingEnabled = true;
            this.cmbTFSServer.Location = new System.Drawing.Point(155, 12);
            this.cmbTFSServer.Name = "cmbTFSServer";
            this.cmbTFSServer.Size = new System.Drawing.Size(272, 21);
            this.cmbTFSServer.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(32, 117);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Build Numbers";
            // 
            // lstBoxDuildNumber
            // 
            this.lstBoxDuildNumber.FormattingEnabled = true;
            this.lstBoxDuildNumber.Location = new System.Drawing.Point(32, 133);
            this.lstBoxDuildNumber.Name = "lstBoxDuildNumber";
            this.lstBoxDuildNumber.Size = new System.Drawing.Size(395, 238);
            this.lstBoxDuildNumber.TabIndex = 7;
            // 
            // GetBuildNumbersForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(449, 427);
            this.Controls.Add(this.lstBoxDuildNumber);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmbTFSServer);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lstBoxProject);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Name = "GetBuildNumbersForm";
            this.Text = "Select a build number";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstBoxProject;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbTFSServer;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox lstBoxDuildNumber;
    }
}