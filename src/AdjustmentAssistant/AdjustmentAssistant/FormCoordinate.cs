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
    public partial class FormCoordinate : Form
    {
        public int gKOrCoord;
        public int gKNo;
        public int inputMidLon;
        public int outputMidLon;

        public FormCoordinate()
        {
            InitializeComponent();
        }

        private void ckbGK_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbGK.Checked == true)
            {
                lblGK.Enabled = true;
                cboGKNo.Enabled = true;
                ckbCoordTransform.Checked = false;
            }
            else
            {
                lblGK.Enabled = false;
                cboGKNo.Enabled = false;
            }
        }

        private void ckbCoordTransform_CheckedChanged(object sender, EventArgs e)
        {
            if (ckbCoordTransform.Checked == true)
            {
                lblOutput.Enabled = true;
                lblOutputMidLon.Enabled = true;
                cboOutputNo.Enabled = true;
                txtOutputMidLon.Enabled = true;
                ckbGK.Checked = false;
            }
            else
            {
                lblOutput.Enabled = false;
                lblOutputMidLon.Enabled = false;
                cboOutputNo.Enabled = false;
                txtOutputMidLon.Enabled = false;
            }
        }

        private void btnSure_Click(object sender, EventArgs e)
        {
            if (txtMidLon.Text.Trim() == "")
            {
                MessageBox.Show("输入坐标的中央经线错误");
                return;
            }
            else
            {
                inputMidLon = Convert.ToInt32(txtMidLon.Text.Trim());
            }
            if (ckbGK.Checked == true)
            {
                gKOrCoord = 1;
                if (cboGKNo.SelectedIndex == -1)
                {
                    MessageBox.Show("没有做出有效选择");
                    return;
                }
                if (cboGKNo.SelectedIndex == 0)
                {
                    gKNo = 3;
                }
                else
                {
                    gKNo = 6;
                }
                this.DialogResult = DialogResult.OK;
            }
            else if (ckbCoordTransform.Checked == true)
            {
                gKOrCoord = 2;
                if (txtOutputMidLon.Text.Trim() == "" || cboOutputNo.SelectedIndex == -1)
                {
                    MessageBox.Show("计算参数设置错误");
                    return;
                }
                if (cboOutputNo.SelectedIndex == 0)
                {
                    gKNo = 3;
                }
                else
                {
                    gKNo = 6;
                }
                outputMidLon = Convert.ToInt32(txtOutputMidLon.Text.Trim());
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("计算参数设置不正确");
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
