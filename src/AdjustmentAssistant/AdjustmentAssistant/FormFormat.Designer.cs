namespace AdjustmentAssistant
{
    partial class FormFormat
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
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cboDataType = new System.Windows.Forms.ComboBox();
            this.cboDirection = new System.Windows.Forms.ComboBox();
            this.txtDataNumber = new System.Windows.Forms.TextBox();
            this.txtBackCount = new System.Windows.Forms.TextBox();
            this.btnSure = new System.Windows.Forms.Button();
            this.btnCancle = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.cboLevelingWeight = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "数据类型";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(25, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "数据条数";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(25, 88);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "测回数";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(25, 121);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 0;
            this.label4.Text = "测角方向";
            // 
            // cboDataType
            // 
            this.cboDataType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboDataType.FormattingEnabled = true;
            this.cboDataType.Items.AddRange(new object[] {
            "闭附合导线",
            "支导线",
            "单面单程水准",
            "红黑双面水准"});
            this.cboDataType.Location = new System.Drawing.Point(98, 20);
            this.cboDataType.Name = "cboDataType";
            this.cboDataType.Size = new System.Drawing.Size(121, 20);
            this.cboDataType.TabIndex = 1;
            this.cboDataType.SelectedIndexChanged += new System.EventHandler(this.cboDataType_SelectedIndexChanged);
            // 
            // cboDirection
            // 
            this.cboDirection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboDirection.FormattingEnabled = true;
            this.cboDirection.Items.AddRange(new object[] {
            "左角",
            "右角"});
            this.cboDirection.Location = new System.Drawing.Point(98, 118);
            this.cboDirection.Name = "cboDirection";
            this.cboDirection.Size = new System.Drawing.Size(121, 20);
            this.cboDirection.TabIndex = 1;
            // 
            // txtDataNumber
            // 
            this.txtDataNumber.Location = new System.Drawing.Point(98, 52);
            this.txtDataNumber.Name = "txtDataNumber";
            this.txtDataNumber.Size = new System.Drawing.Size(121, 21);
            this.txtDataNumber.TabIndex = 2;
            // 
            // txtBackCount
            // 
            this.txtBackCount.Location = new System.Drawing.Point(98, 85);
            this.txtBackCount.Name = "txtBackCount";
            this.txtBackCount.Size = new System.Drawing.Size(121, 21);
            this.txtBackCount.TabIndex = 2;
            // 
            // btnSure
            // 
            this.btnSure.Location = new System.Drawing.Point(39, 184);
            this.btnSure.Name = "btnSure";
            this.btnSure.Size = new System.Drawing.Size(75, 23);
            this.btnSure.TabIndex = 3;
            this.btnSure.Text = "确定";
            this.btnSure.UseVisualStyleBackColor = true;
            this.btnSure.Click += new System.EventHandler(this.btnSure_Click);
            // 
            // btnCancle
            // 
            this.btnCancle.Location = new System.Drawing.Point(130, 184);
            this.btnCancle.Name = "btnCancle";
            this.btnCancle.Size = new System.Drawing.Size(75, 23);
            this.btnCancle.TabIndex = 3;
            this.btnCancle.Text = "取消";
            this.btnCancle.UseVisualStyleBackColor = true;
            this.btnCancle.Click += new System.EventHandler(this.btnCancle_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(25, 153);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 0;
            this.label5.Text = "水准观测权";
            // 
            // cboLevelingWeight
            // 
            this.cboLevelingWeight.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboLevelingWeight.FormattingEnabled = true;
            this.cboLevelingWeight.Items.AddRange(new object[] {
            "测站数",
            "距离"});
            this.cboLevelingWeight.Location = new System.Drawing.Point(98, 150);
            this.cboLevelingWeight.Name = "cboLevelingWeight";
            this.cboLevelingWeight.Size = new System.Drawing.Size(121, 20);
            this.cboLevelingWeight.TabIndex = 1;
            // 
            // FormFormat
            // 
            this.AcceptButton = this.btnSure;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(245, 227);
            this.Controls.Add(this.btnCancle);
            this.Controls.Add(this.btnSure);
            this.Controls.Add(this.txtBackCount);
            this.Controls.Add(this.txtDataNumber);
            this.Controls.Add(this.cboLevelingWeight);
            this.Controls.Add(this.cboDirection);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cboDataType);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormFormat";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "导线数据格式";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cboDataType;
        private System.Windows.Forms.ComboBox cboDirection;
        private System.Windows.Forms.TextBox txtDataNumber;
        private System.Windows.Forms.TextBox txtBackCount;
        private System.Windows.Forms.Button btnSure;
        private System.Windows.Forms.Button btnCancle;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cboLevelingWeight;
    }
}