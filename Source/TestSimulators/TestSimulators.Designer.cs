namespace TestSimulators
{
    partial class TestSimulators
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
            this.btnForkLift = new System.Windows.Forms.Button();
            this.btnProngLift = new System.Windows.Forms.Button();
            this.BtnTestFile = new System.Windows.Forms.Button();
            this.TxtOutput = new System.Windows.Forms.TextBox();
            this.btnClearLogScreen = new System.Windows.Forms.Button();
            this.btnStop = new System.Windows.Forms.Button();
            this.chkFILempty = new System.Windows.Forms.CheckBox();
            this.btnSmallProngLift = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnForkLift
            // 
            this.btnForkLift.Location = new System.Drawing.Point(40, 12);
            this.btnForkLift.Name = "btnForkLift";
            this.btnForkLift.Size = new System.Drawing.Size(75, 23);
            this.btnForkLift.TabIndex = 0;
            this.btnForkLift.Text = "ForkLift";
            this.btnForkLift.UseVisualStyleBackColor = true;
            this.btnForkLift.Click += new System.EventHandler(this.btnForkLift_Click);
            // 
            // btnProngLift
            // 
            this.btnProngLift.Location = new System.Drawing.Point(40, 41);
            this.btnProngLift.Name = "btnProngLift";
            this.btnProngLift.Size = new System.Drawing.Size(75, 23);
            this.btnProngLift.TabIndex = 1;
            this.btnProngLift.Text = "ProngLift";
            this.btnProngLift.UseVisualStyleBackColor = true;
            this.btnProngLift.Click += new System.EventHandler(this.btnProngLift_Click);
            // 
            // BtnTestFile
            // 
            this.BtnTestFile.Location = new System.Drawing.Point(40, 114);
            this.BtnTestFile.Name = "BtnTestFile";
            this.BtnTestFile.Size = new System.Drawing.Size(111, 23);
            this.BtnTestFile.TabIndex = 2;
            this.BtnTestFile.Text = "CreateTestFile";
            this.BtnTestFile.UseVisualStyleBackColor = true;
            this.BtnTestFile.Click += new System.EventHandler(this.BtnTestFile_Click);
            // 
            // TxtOutput
            // 
            this.TxtOutput.Location = new System.Drawing.Point(4, 143);
            this.TxtOutput.Multiline = true;
            this.TxtOutput.Name = "TxtOutput";
            this.TxtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TxtOutput.Size = new System.Drawing.Size(509, 556);
            this.TxtOutput.TabIndex = 3;
            // 
            // btnClearLogScreen
            // 
            this.btnClearLogScreen.Location = new System.Drawing.Point(363, 88);
            this.btnClearLogScreen.Name = "btnClearLogScreen";
            this.btnClearLogScreen.Size = new System.Drawing.Size(136, 23);
            this.btnClearLogScreen.TabIndex = 4;
            this.btnClearLogScreen.Text = "Clear Log Screen";
            this.btnClearLogScreen.UseVisualStyleBackColor = true;
            this.btnClearLogScreen.Click += new System.EventHandler(this.btnClearLogScreen_Click);
            // 
            // btnStop
            // 
            this.btnStop.Location = new System.Drawing.Point(197, 114);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(75, 23);
            this.btnStop.TabIndex = 5;
            this.btnStop.Text = "Stop";
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            // 
            // chkFILempty
            // 
            this.chkFILempty.AutoSize = true;
            this.chkFILempty.Location = new System.Drawing.Point(140, 45);
            this.chkFILempty.Name = "chkFILempty";
            this.chkFILempty.Size = new System.Drawing.Size(132, 17);
            this.chkFILempty.TabIndex = 6;
            this.chkFILempty.Text = "Check drop loc. empty";
            this.chkFILempty.UseVisualStyleBackColor = true;
            this.chkFILempty.CheckedChanged += new System.EventHandler(this.chkFILempty_CheckedChanged);
            // 
            // btnSmallProngLift
            // 
            this.btnSmallProngLift.Location = new System.Drawing.Point(40, 70);
            this.btnSmallProngLift.Name = "btnSmallProngLift";
            this.btnSmallProngLift.Size = new System.Drawing.Size(75, 23);
            this.btnSmallProngLift.TabIndex = 7;
            this.btnSmallProngLift.Text = "ProngLift-2.5T";
            this.btnSmallProngLift.UseVisualStyleBackColor = true;
            this.btnSmallProngLift.Click += new System.EventHandler(this.btnSmallProngLift_Click);
            // 
            // TestSimulators
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(525, 696);
            this.Controls.Add(this.btnSmallProngLift);
            this.Controls.Add(this.chkFILempty);
            this.Controls.Add(this.btnStop);
            this.Controls.Add(this.btnClearLogScreen);
            this.Controls.Add(this.TxtOutput);
            this.Controls.Add(this.BtnTestFile);
            this.Controls.Add(this.btnProngLift);
            this.Controls.Add(this.btnForkLift);
            this.Name = "TestSimulators";
            this.Text = "TestSimulators";
            this.Load += new System.EventHandler(this.TestSimulators_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnForkLift;
        private System.Windows.Forms.Button btnProngLift;
        private System.Windows.Forms.Button BtnTestFile;
        private System.Windows.Forms.TextBox TxtOutput;
        private System.Windows.Forms.Button btnClearLogScreen;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.CheckBox chkFILempty;
        private System.Windows.Forms.Button btnSmallProngLift;
    }
}

