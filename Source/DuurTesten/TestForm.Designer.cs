namespace DuurTesten
{
    partial class TestForm
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
            this.btnStart = new System.Windows.Forms.Button();
            this.btnStop = new System.Windows.Forms.Button();
            this.cmbApplication = new System.Windows.Forms.ComboBox();
            this.lbDuration = new System.Windows.Forms.Label();
            this.txtBoxDuration = new System.Windows.Forms.TextBox();
            this.lbStart = new System.Windows.Forms.Label();
            this.lbStartTime = new System.Windows.Forms.Label();
            this.lbEnd = new System.Windows.Forms.Label();
            this.lbEndTime = new System.Windows.Forms.Label();
            this.lbApp = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(52, 89);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // btnStop
            // 
            this.btnStop.Location = new System.Drawing.Point(192, 89);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(75, 23);
            this.btnStop.TabIndex = 1;
            this.btnStop.Text = "Stop";
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            // 
            // cmbApplication
            // 
            this.cmbApplication.FormattingEnabled = true;
            this.cmbApplication.Items.AddRange(new object[] {
            "Etricc",
            "Ewms"});
            this.cmbApplication.Location = new System.Drawing.Point(112, 43);
            this.cmbApplication.Name = "cmbApplication";
            this.cmbApplication.Size = new System.Drawing.Size(121, 21);
            this.cmbApplication.TabIndex = 2;
            this.cmbApplication.Text = "Etricc";
            // 
            // lbDuration
            // 
            this.lbDuration.AutoSize = true;
            this.lbDuration.Location = new System.Drawing.Point(9, 42);
            this.lbDuration.Name = "lbDuration";
            this.lbDuration.Size = new System.Drawing.Size(76, 13);
            this.lbDuration.TabIndex = 3;
            this.lbDuration.Text = "Duration(Hour)";
            // 
            // txtBoxDuration
            // 
            this.txtBoxDuration.Location = new System.Drawing.Point(123, 39);
            this.txtBoxDuration.Name = "txtBoxDuration";
            this.txtBoxDuration.Size = new System.Drawing.Size(71, 20);
            this.txtBoxDuration.TabIndex = 4;
            this.txtBoxDuration.Text = "10";
            // 
            // lbStart
            // 
            this.lbStart.AutoSize = true;
            this.lbStart.Location = new System.Drawing.Point(27, 16);
            this.lbStart.Name = "lbStart";
            this.lbStart.Size = new System.Drawing.Size(58, 13);
            this.lbStart.TabIndex = 5;
            this.lbStart.Text = "Start Time:";
            // 
            // lbStartTime
            // 
            this.lbStartTime.AutoSize = true;
            this.lbStartTime.Location = new System.Drawing.Point(126, 16);
            this.lbStartTime.Name = "lbStartTime";
            this.lbStartTime.Size = new System.Drawing.Size(49, 13);
            this.lbStartTime.TabIndex = 6;
            this.lbStartTime.Text = "00:00:00";
            // 
            // lbEnd
            // 
            this.lbEnd.AutoSize = true;
            this.lbEnd.Location = new System.Drawing.Point(30, 75);
            this.lbEnd.Name = "lbEnd";
            this.lbEnd.Size = new System.Drawing.Size(55, 13);
            this.lbEnd.TabIndex = 7;
            this.lbEnd.Text = "End Time:";
            // 
            // lbEndTime
            // 
            this.lbEndTime.AutoSize = true;
            this.lbEndTime.Location = new System.Drawing.Point(126, 75);
            this.lbEndTime.Name = "lbEndTime";
            this.lbEndTime.Size = new System.Drawing.Size(49, 13);
            this.lbEndTime.TabIndex = 8;
            this.lbEndTime.Text = "10:00:00";
            // 
            // lbApp
            // 
            this.lbApp.AutoSize = true;
            this.lbApp.Location = new System.Drawing.Point(49, 50);
            this.lbApp.Name = "lbApp";
            this.lbApp.Size = new System.Drawing.Size(59, 13);
            this.lbApp.TabIndex = 9;
            this.lbApp.Text = "Application";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lbStartTime);
            this.groupBox1.Controls.Add(this.lbStart);
            this.groupBox1.Controls.Add(this.txtBoxDuration);
            this.groupBox1.Controls.Add(this.lbEndTime);
            this.groupBox1.Controls.Add(this.lbDuration);
            this.groupBox1.Controls.Add(this.lbEnd);
            this.groupBox1.Location = new System.Drawing.Point(377, 24);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(278, 100);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Time";
            // 
            // TestForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(699, 266);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lbApp);
            this.Controls.Add(this.cmbApplication);
            this.Controls.Add(this.btnStop);
            this.Controls.Add(this.btnStart);
            this.Name = "TestForm";
            this.Text = "DuurTesten";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.ComboBox cmbApplication;
        private System.Windows.Forms.Label lbDuration;
        private System.Windows.Forms.TextBox txtBoxDuration;
        private System.Windows.Forms.Label lbStart;
        private System.Windows.Forms.Label lbStartTime;
        private System.Windows.Forms.Label lbEnd;
        private System.Windows.Forms.Label lbEndTime;
        private System.Windows.Forms.Label lbApp;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}

