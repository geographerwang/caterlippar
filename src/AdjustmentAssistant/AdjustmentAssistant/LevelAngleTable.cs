using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;

namespace AdjustmentAssistant
{
    public class LevelAngleTable
    {
        public string strTitle;
        public string strProjectName;
        public string strInstrument;
        public string strWeather;
        public string strObserver;
        public string strRecorder;
        public string strDate;
        public string strCalculate;
        public string strAssessment;

        internal void DrawTable(Panel pnlResult, int dataCount, int backCount, int dataType, TableLayoutPanel tableLayoutPanel)
        {
            if (dataCount < 1)
            {
                return;
            }

            //绘制标题记录项
            pnlResult.Controls.Clear();
            TextBox txtTitle = new TextBox();
            txtTitle.Location = new Point(45, 45);
            txtTitle.BorderStyle = BorderStyle.None;
            txtTitle.Width = 606;
            txtTitle.Text = "水平角观测记录";
            txtTitle.TextAlign = HorizontalAlignment.Center;
            txtTitle.Font = new Font("宋体", 12, FontStyle.Bold);
            TextBox txtProjectName = new TextBox();
            txtProjectName.Location = new Point(45, 76);
            txtProjectName.BorderStyle = BorderStyle.None;
            txtProjectName.Width = 194;
            txtProjectName.Text = "工程名称：";
            TextBox txtInstrument = new TextBox();
            txtInstrument.Location = new Point(246, 76);
            txtInstrument.BorderStyle = BorderStyle.None;
            txtInstrument.Width = 194;
            txtInstrument.Text = "仪器：";
            TextBox txtWeather = new TextBox();
            txtWeather.Location = new Point(447, 76);
            txtWeather.BorderStyle = BorderStyle.None;
            txtWeather.Width = 194;
            txtWeather.Text = "天气：";
            TextBox txtObserver = new TextBox();
            txtObserver.Location = new Point(45, 97);
            txtObserver.BorderStyle = BorderStyle.None;
            txtObserver.Width = 194;
            txtObserver.Text = "观测者：";
            TextBox txtRecorder = new TextBox();
            txtRecorder.Location = new Point(246, 97);
            txtRecorder.BorderStyle = BorderStyle.None;
            txtRecorder.Width = 194;
            txtRecorder.Text = "记录者：";
            TextBox txtDate = new TextBox();
            txtDate.Location = new Point(447, 97);
            txtDate.BorderStyle = BorderStyle.None;
            txtDate.Width = 194;
            txtDate.Text = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");

            //绘制表格总体布局
            tableLayoutPanel.RowCount = dataCount + 3;
            tableLayoutPanel.ColumnCount = 7;
            tableLayoutPanel.Location = new Point(45, 120);
            tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            tableLayoutPanel.Width = 608;
            tableLayoutPanel.Height = tableLayoutPanel.RowCount * 22 + 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 56f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 56f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 56f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108f));
            //测站点
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i += (2 * backCount))
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol0 = new TextBox();
                txtBoxCol0.Multiline = true;
                txtBoxCol0.Margin = new Padding(0);
                txtBoxCol0.Name = "txtBoxCol1" + i;
                txtBoxCol0.Width = 56;
                txtBoxCol0.Height = backCount * 44;
                txtBoxCol0.TextAlign = HorizontalAlignment.Center;
                txtBoxCol0.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol0, 0, i);
                Control[] ctrlTxtBoxCol0 = tableLayoutPanel.Controls.Find(txtBoxCol0.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol0[0], backCount * 2);
            }
            //照准点
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol = new TextBox();
                txtBoxCol.Margin = new Padding(0);
                txtBoxCol.Width = 56;
                txtBoxCol.Height = 21;
                txtBoxCol.TextAlign = HorizontalAlignment.Center;
                txtBoxCol.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol, 1, i);
            }
            //盘左
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol = new TextBox();
                txtBoxCol.Margin = new Padding(0);
                txtBoxCol.Width = 108;
                txtBoxCol.Height = 21;
                txtBoxCol.TextAlign = HorizontalAlignment.Center;
                txtBoxCol.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol, 2, i);
            }
            //盘右
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol = new TextBox();
                txtBoxCol.Margin = new Padding(0);
                txtBoxCol.Width = 108;
                txtBoxCol.Height = 21;
                txtBoxCol.TextAlign = HorizontalAlignment.Center;
                txtBoxCol.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol, 3, i);
            }
            //根据数据类型决定已知点的个数并生成相应的单元格
            if (dataType == 3)//闭附和导线
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol6f = new TextBox();
                txtBoxCol6f.Name = "txtBoxCol6f";
                txtBoxCol6f.Multiline = true;
                txtBoxCol6f.Margin = new Padding(0);
                txtBoxCol6f.Width = 108;
                txtBoxCol6f.Height = backCount * 21;
                txtBoxCol6f.TextAlign = HorizontalAlignment.Center;
                txtBoxCol6f.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol6f, 6, 2);
                Control[] ctrlTxtBxoCol6f = tableLayoutPanel.Controls.Find(txtBoxCol6f.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBxoCol6f[0], backCount);
                TextBox txtBoxCol6l = new TextBox();
                txtBoxCol6l.Name = "txtBoxCol6l";
                txtBoxCol6l.Multiline = true;
                txtBoxCol6l.Margin = new Padding(0);
                txtBoxCol6l.Width = 108;
                txtBoxCol6l.Height = backCount * 44;
                txtBoxCol6l.TextAlign = HorizontalAlignment.Center;
                txtBoxCol6l.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol6l, 6, tableLayoutPanel.RowCount - 1 - backCount * 2);
                Control[] ctrlTxtBoxCol6l = tableLayoutPanel.Controls.Find(txtBoxCol6l.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol6l[0], backCount * 2);
            }
            else if (dataType == 4)//支导线
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol6 = new TextBox();
                txtBoxCol6.Name = "txtBoxCol6";
                txtBoxCol6.Multiline = true;
                txtBoxCol6.Margin = new Padding(0);
                txtBoxCol6.Width = 108;
                txtBoxCol6.Height = 22 * backCount;
                txtBoxCol6.TextAlign = HorizontalAlignment.Center;
                txtBoxCol6.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol6, 6, 2);
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                Control[] ctrlTxtBoxCol6 = tableLayoutPanel.Controls.Find(txtBoxCol6.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol6[0], backCount);
            }
            //绘制表头
            Label lblPointName = new Label();
            lblPointName.Name = "lblPointName";
            lblPointName.Margin = new Padding(0);
            lblPointName.Width = 56;
            lblPointName.Height = 42;
            lblPointName.Text = "测站名";
            lblPointName.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblPointName, 0, 0);
            Control[] ctrlLblPointName = tableLayoutPanel.Controls.Find("lblPointName", false);
            tableLayoutPanel.SetRowSpan(ctrlLblPointName[0], 2);
            Label lblAimPoint = new Label();
            lblAimPoint.Name = "lblAimPoint";
            lblAimPoint.Margin = new Padding(0);
            lblAimPoint.Width = 56;
            lblAimPoint.Height = 42;
            lblAimPoint.Text = "照准点";
            lblAimPoint.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblAimPoint, 1, 0);
            Control[] ctrlLblAimPoint = tableLayoutPanel.Controls.Find("lblAimPoint", false);
            tableLayoutPanel.SetRowSpan(ctrlLblAimPoint[0], 2);
            Label lblHorizontal = new Label();
            lblHorizontal.Name = "lblHorizontal";
            lblHorizontal.Margin = new Padding(0);
            lblHorizontal.Width = 488;
            lblHorizontal.Height = 21;
            lblHorizontal.Text = "水平角(° ' \")";
            lblHorizontal.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblHorizontal, 2, 0);
            Control[] ctrlLblHorizontal = tableLayoutPanel.Controls.Find("lblHorizontal", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblHorizontal[0], 5);
            Label lblLeft = new Label();
            lblLeft.Margin = new Padding(0);
            lblLeft.Width = 108;
            lblLeft.Height = 21;
            lblLeft.Text = "盘左读数";
            lblLeft.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblLeft, 2, 1);
            Label lblRight = new Label();
            lblRight.Margin = new Padding(0);
            lblRight.Width = 108;
            lblRight.Height = 21;
            lblRight.Text = "盘右读数";
            lblRight.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblRight, 3, 1);
            Label lbl2C = new Label();
            lbl2C.Margin = new Padding(0);
            lbl2C.Width = 56;
            lbl2C.Height = 21;
            lbl2C.Text = "2C(\")";
            lbl2C.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lbl2C, 4, 1);
            Label lblAngle = new Label();
            lblAngle.Margin = new Padding(0);
            lblAngle.Width = 108;
            lblAngle.Height = 21;
            lblAngle.Text = "角值";
            lblAngle.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblAngle, 5, 1);
            Label lblAzimuth = new Label();
            lblAzimuth.Width = 108;
            lblAzimuth.Height = 21;
            lblAzimuth.Text = "方位角";
            lblAzimuth.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblAzimuth, 6, 1);

            //表末尾备注
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
            Label lblRemark = new Label();
            lblRemark.Width = 56;
            lblRemark.Height = 21;
            lblRemark.Text = "备注";
            lblRemark.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblRemark, 0, tableLayoutPanel.RowCount - 1);
            TextBox txtCalculate = new TextBox();
            txtCalculate.Location = new Point(48, tableLayoutPanel.Height + 125);
            txtCalculate.Width = 195;
            txtCalculate.Height = 21;
            txtCalculate.Text = "计算者：";
            txtCalculate.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtCalculate);
            TextBox txtAssessment = new TextBox();
            txtAssessment.Location = new Point(246, tableLayoutPanel.Height + 125);
            txtAssessment.Width = 195;
            txtAssessment.Height = 21;
            txtAssessment.Text = "审核者：";
            txtAssessment.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtAssessment);
            TextBox txtDataEnd = new TextBox();
            txtDataEnd.Location = new Point(442, tableLayoutPanel.Height + 125);
            txtDataEnd.Width = 195;
            txtDataEnd.Height = 21;
            txtDataEnd.Text = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            txtDataEnd.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtDataEnd);
            Label lbl = new Label();
            lbl.Location = new Point(45, 171 + tableLayoutPanel.Height);

            //将控件添加到panel
            pnlResult.Controls.Add(txtTitle);
            pnlResult.Controls.Add(txtProjectName);
            pnlResult.Controls.Add(txtInstrument);
            pnlResult.Controls.Add(txtWeather);
            pnlResult.Controls.Add(txtObserver);
            pnlResult.Controls.Add(txtRecorder);
            pnlResult.Controls.Add(txtDate);
            pnlResult.Controls.Add(tableLayoutPanel);
            pnlResult.Controls.Add(lbl);

            //将表格附属记录储存起来
            strTitle = txtTitle.Text;
            strProjectName = txtProjectName.Text;
            strInstrument = txtInstrument.Text;
            strWeather = txtWeather.Text;
            strObserver = txtObserver.Text;
            strRecorder = txtRecorder.Text;
            strDate = txtDate.Text;
            strCalculate = txtCalculate.Text;
            strAssessment = txtAssessment.Text;
        }
    }
}
