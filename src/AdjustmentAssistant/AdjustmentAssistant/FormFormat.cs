using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace AdjustmentAssistant
{
    public partial class FormFormat : Form
    {
        public FormFormat()
        {
            InitializeComponent();
        }



        public DataType.Data TraverseDataType
        {
            get;
            set;
        }

        public DataType.LeftOrRight LorR
        {
            get;
            set;
        }

        public DataType.LevelingWeight LvlWet
        {
            get;
            set;
        }

        public int DataCount
        {
            get;
            set;
        }

        public int BackCount
        {
            get;
            set;
        }

        private void cboDataType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboDataType.SelectedIndex > 1)
            {
                label3.Enabled = false;
                label4.Enabled = false;
                label5.Enabled = true;
                txtBackCount.Enabled = false;
                cboDirection.Enabled = false;
                cboLevelingWeight.Enabled = true;
            }
            else
            {
                label3.Enabled = true;
                label4.Enabled = true;
                label5.Enabled = false;
                txtBackCount.Enabled = true;
                cboDirection.Enabled = true;
                cboLevelingWeight.Enabled = false;
            }
        }

        private void btnSure_Click(object sender, EventArgs e)
        {
            try
            {
                if (cboDataType.SelectedIndex == 0 && !string.IsNullOrEmpty(txtDataNumber.Text.Trim()))
                {
                    TraverseDataType = DataType.Data.ConnectingTraverse;
                    DataCount = Convert.ToInt32(txtDataNumber.Text.Trim());
                    BackCount = Convert.ToInt32(txtBackCount.Text.Trim());
                    LorR = (DataType.LeftOrRight)Enum.Parse(typeof(DataType.LeftOrRight), cboDirection.SelectedIndex.ToString());
                    if (DataCount % (BackCount * 2) != 0)
                    {
                        MessageBox.Show("错误信息", "数据总条数不能按测回分配!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (cboDataType.SelectedIndex == 1 && !string.IsNullOrEmpty(txtDataNumber.Text.Trim()))
                {
                    TraverseDataType = DataType.Data.OpenTraverse;
                    DataCount = Convert.ToInt32(txtDataNumber.Text.Trim());
                    BackCount = Convert.ToInt32(txtBackCount.Text.Trim());
                    LorR = (DataType.LeftOrRight)Enum.Parse(typeof(DataType.LeftOrRight), cboDirection.SelectedIndex.ToString());
                    if (DataCount % (BackCount * 2) != 0)
                    {
                        MessageBox.Show("错误信息", "数据总条数不能按测回分配!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else if (cboDataType.SelectedIndex == 2 && !string.IsNullOrEmpty(txtDataNumber.Text.Trim()))
                {
                    TraverseDataType = DataType.Data.SingleRule;
                    DataCount = Convert.ToInt32(txtDataNumber.Text.Trim());
                    LvlWet = (DataType.LevelingWeight)Enum.Parse(typeof(DataType.LevelingWeight), cboLevelingWeight.SelectedIndex.ToString());
                }
                else if (cboDataType.SelectedIndex == 3 && !string.IsNullOrEmpty(txtDataNumber.Text.Trim()))
                {
                    TraverseDataType = DataType.Data.DoubleRule;
                    DataCount = Convert.ToInt32(txtDataNumber.Text.Trim());
                    LvlWet = (DataType.LevelingWeight)Enum.Parse(typeof(DataType.LevelingWeight), cboLevelingWeight.SelectedIndex.ToString());
                }
                DialogResult = DialogResult.OK;
            }
            catch
            {
                MessageBox.Show("错误信息", "输入的参数有误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancle_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
