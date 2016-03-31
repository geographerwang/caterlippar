using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AdjustmentAssistant
{
    public partial class FormAccuracy : Form
    {
        public string lblPointName;
        public double direction = 1;
        public double side = 0.005;
        public double sidePercent = 0.000001;
        public double angle = 0.7;
        public double pointX;
        public double pointY;

        public FormAccuracy()
        {
            InitializeComponent();
        }

        public void Init()
        {
            txtX.Text = pointX.ToString();
            txtY.Text = pointY.ToString();
            lblName.Text = lblPointName + lblName.Text;
        }

        private void btnSure_Click(object sender, EventArgs e)
        {
            try
            {
                direction = Convert.ToDouble(txtDirection.Text.Trim());
                side = Convert.ToDouble(txtSide.Text.Trim());
                sidePercent = Convert.ToDouble(txtSidePercent.Text.Trim());
                angle = Convert.ToDouble(txtAngle.Text.Trim());
                pointX = Convert.ToDouble(txtX.Text.Trim());
                pointY = Convert.ToDouble(txtY.Text.Trim());
                if (pointX < 0 || pointY < 0)
                {
                    MessageBox.Show("坐标值不能为负数！");
                }
            }
            catch
            {
                MessageBox.Show("输入的数据格式不正确\n请重新输入！");
            }
            this.DialogResult = DialogResult.OK;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
