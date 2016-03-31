namespace AdjustmentAssistant
{
    partial class FormCoordinate
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
            this.txtMidLon = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cboGKNo = new System.Windows.Forms.ComboBox();
            this.lblGK = new System.Windows.Forms.Label();
            this.ckbGK = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cboOutputNo = new System.Windows.Forms.ComboBox();
            this.lblOutput = new System.Windows.Forms.Label();
            this.ckbCoordTransform = new System.Windows.Forms.CheckBox();
            this.txtOutputMidLon = new System.Windows.Forms.TextBox();
            this.lblOutputMidLon = new System.Windows.Forms.Label();
            this.btnSure = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "输入中央子午线";
            // 
            // txtMidLon
            // 
            this.txtMidLon.Location = new System.Drawing.Point(107, 18);
            this.txtMidLon.Name = "txtMidLon";
            this.txtMidLon.Size = new System.Drawing.Size(100, 21);
            this.txtMidLon.TabIndex = 1;
            this.txtMidLon.Text = "0";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cboGKNo);
            this.groupBox1.Controls.Add(this.lblGK);
            this.groupBox1.Controls.Add(this.ckbGK);
            this.groupBox1.Location = new System.Drawing.Point(14, 45);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(193, 58);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // cboGKNo
            // 
            this.cboGKNo.Enabled = false;
            this.cboGKNo.FormattingEnabled = true;
            this.cboGKNo.Items.AddRange(new object[] {
            "3度带",
            "6度带"});
            this.cboGKNo.Location = new System.Drawing.Point(111, 27);
            this.cboGKNo.Name = "cboGKNo";
            this.cboGKNo.Size = new System.Drawing.Size(76, 20);
            this.cboGKNo.TabIndex = 2;
            // 
            // lblGK
            // 
            this.lblGK.AutoSize = true;
            this.lblGK.Enabled = false;
            this.lblGK.Location = new System.Drawing.Point(6, 30);
            this.lblGK.Name = "lblGK";
            this.lblGK.Size = new System.Drawing.Size(101, 12);
            this.lblGK.TabIndex = 1;
            this.lblGK.Text = "高斯投影坐标分带";
            // 
            // ckbGK
            // 
            this.ckbGK.AutoSize = true;
            this.ckbGK.Location = new System.Drawing.Point(6, 0);
            this.ckbGK.Name = "ckbGK";
            this.ckbGK.Size = new System.Drawing.Size(84, 16);
            this.ckbGK.TabIndex = 0;
            this.ckbGK.Text = "高斯正反算";
            this.ckbGK.UseVisualStyleBackColor = true;
            this.ckbGK.CheckedChanged += new System.EventHandler(this.ckbGK_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cboOutputNo);
            this.groupBox2.Controls.Add(this.lblOutput);
            this.groupBox2.Controls.Add(this.ckbCoordTransform);
            this.groupBox2.Controls.Add(this.txtOutputMidLon);
            this.groupBox2.Controls.Add(this.lblOutputMidLon);
            this.groupBox2.Location = new System.Drawing.Point(15, 109);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(193, 87);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "groupBox1";
            // 
            // cboOutputNo
            // 
            this.cboOutputNo.Enabled = false;
            this.cboOutputNo.FormattingEnabled = true;
            this.cboOutputNo.Items.AddRange(new object[] {
            "3度带",
            "6度带"});
            this.cboOutputNo.Location = new System.Drawing.Point(89, 28);
            this.cboOutputNo.Name = "cboOutputNo";
            this.cboOutputNo.Size = new System.Drawing.Size(97, 20);
            this.cboOutputNo.TabIndex = 2;
            // 
            // lblOutput
            // 
            this.lblOutput.AutoSize = true;
            this.lblOutput.Enabled = false;
            this.lblOutput.Location = new System.Drawing.Point(6, 31);
            this.lblOutput.Name = "lblOutput";
            this.lblOutput.Size = new System.Drawing.Size(77, 12);
            this.lblOutput.TabIndex = 1;
            this.lblOutput.Text = "输出坐标分带";
            // 
            // ckbCoordTransform
            // 
            this.ckbCoordTransform.AutoSize = true;
            this.ckbCoordTransform.Location = new System.Drawing.Point(6, 0);
            this.ckbCoordTransform.Name = "ckbCoordTransform";
            this.ckbCoordTransform.Size = new System.Drawing.Size(96, 16);
            this.ckbCoordTransform.TabIndex = 0;
            this.ckbCoordTransform.Text = "坐标分带转换";
            this.ckbCoordTransform.UseVisualStyleBackColor = true;
            this.ckbCoordTransform.CheckedChanged += new System.EventHandler(this.ckbCoordTransform_CheckedChanged);
            // 
            // txtOutputMidLon
            // 
            this.txtOutputMidLon.Enabled = false;
            this.txtOutputMidLon.Location = new System.Drawing.Point(101, 54);
            this.txtOutputMidLon.Name = "txtOutputMidLon";
            this.txtOutputMidLon.Size = new System.Drawing.Size(85, 21);
            this.txtOutputMidLon.TabIndex = 1;
            this.txtOutputMidLon.Text = "0";
            // 
            // lblOutputMidLon
            // 
            this.lblOutputMidLon.AutoSize = true;
            this.lblOutputMidLon.Enabled = false;
            this.lblOutputMidLon.Location = new System.Drawing.Point(6, 57);
            this.lblOutputMidLon.Name = "lblOutputMidLon";
            this.lblOutputMidLon.Size = new System.Drawing.Size(89, 12);
            this.lblOutputMidLon.TabIndex = 0;
            this.lblOutputMidLon.Text = "输出中央子午线";
            // 
            // btnSure
            // 
            this.btnSure.Location = new System.Drawing.Point(32, 202);
            this.btnSure.Name = "btnSure";
            this.btnSure.Size = new System.Drawing.Size(64, 21);
            this.btnSure.TabIndex = 3;
            this.btnSure.Text = "确定";
            this.btnSure.UseVisualStyleBackColor = true;
            this.btnSure.Click += new System.EventHandler(this.btnSure_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(124, 202);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(64, 21);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // FormCoordinate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(221, 235);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSure);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.txtMidLon);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormCoordinate";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "高斯投影参数";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtMidLon;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cboGKNo;
        private System.Windows.Forms.Label lblGK;
        private System.Windows.Forms.CheckBox ckbGK;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox cboOutputNo;
        private System.Windows.Forms.Label lblOutput;
        private System.Windows.Forms.CheckBox ckbCoordTransform;
        private System.Windows.Forms.Button btnSure;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtOutputMidLon;
        private System.Windows.Forms.Label lblOutputMidLon;
    }
}