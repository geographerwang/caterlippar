using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace AdjustmentAssistant
{
    class XYToXYTable
    {
        internal void DrawTable(Panel pnlResult, int dataCount, ref TableLayoutPanel tableLayoutPanel)
        {
            if (dataCount == 0)
            {
                return;
            }
            pnlResult.Controls.Clear();
            TextBox txtTitle = new TextBox();
            txtTitle.Location = new Point(45, 45);
            txtTitle.BorderStyle = BorderStyle.None;
            txtTitle.Width = 505;
            txtTitle.Text = "坐标换带计算";
            txtTitle.TextAlign = HorizontalAlignment.Center;
            txtTitle.Font = new Font("宋体", 12, FontStyle.Bold);
            TextBox txtProjectName = new TextBox();
            txtProjectName.Location = new Point(45, 76);
            txtProjectName.BorderStyle = BorderStyle.None;
            txtProjectName.Width = 200;
            txtProjectName.Text = "工程名称：";
            TextBox txtDate = new TextBox();
            txtDate.Location = new Point(345, 76);
            txtDate.BorderStyle = BorderStyle.None;
            txtDate.Width = 194;
            txtDate.Text = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            //表格整体布局
            tableLayoutPanel.RowCount = dataCount + 2;
            tableLayoutPanel.ColumnCount = 5;
            tableLayoutPanel.Location = new Point(45, 99);
            tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            tableLayoutPanel.Width = 505;
            tableLayoutPanel.Height = tableLayoutPanel.RowCount * 22 + 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105f));
            //表头
            Label lblPointName = new Label();
            lblPointName.Name = "lblPointName";
            lblPointName.Margin = new Padding(0);
            lblPointName.Width = 79;
            lblPointName.Height = 42;
            lblPointName.Text = "点名";
            lblPointName.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblPointName, 0, 0);
            Control[] ctrlLblPointName = tableLayoutPanel.Controls.Find("lblPointName", false);
            tableLayoutPanel.SetRowSpan(ctrlLblPointName[0], 2);
            Label lblFromXY = new Label();
            lblFromXY.Name = "lblFromXY";
            lblFromXY.Margin = new Padding(0);
            lblFromXY.Width = 210;
            lblFromXY.Height = 21;
            lblFromXY.Text = "转换前的坐标(m)";
            lblFromXY.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblFromXY, 1, 0);
            Control[] ctrlLblXY = tableLayoutPanel.Controls.Find("lblFromXY", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblXY[0], 2);
            Label lblX = new Label();
            lblX.Margin = new Padding(0);
            lblX.Width = 103;
            lblX.Text = "X";
            lblX.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblX, 1, 1);
            Label lblY = new Label();
            lblY.Margin = new Padding(0);
            lblY.Width = 103;
            lblY.Text = "Y";
            lblY.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblY, 2, 1);
            Label lblToXY = new Label();
            lblToXY.Name = "lblToXY";
            lblToXY.Margin = new Padding(0);
            lblToXY.Width = 210;
            lblToXY.Text = "转换后的坐标(m)";
            lblToXY.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblToXY, 3, 0);
            Control[] ctrlLblBL = tableLayoutPanel.Controls.Find("lblToXY", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblBL[0], 2);
            Label lblX2 = new Label();
            lblX2.Margin = new Padding(0);
            lblX2.Width = 103;
            lblX2.Text = "X";
            lblX2.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblX2, 3, 1);
            Label lblY2 = new Label();
            lblY2.Margin = new Padding(0);
            lblY2.Width = 103;
            lblY2.Text = "Y";
            lblY2.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblY2, 4, 1);

            TextBox txtCalculate = new TextBox();
            txtCalculate.Location = new Point(48, tableLayoutPanel.Height + 104);
            txtCalculate.Width = 150;
            txtCalculate.Height = 21;
            txtCalculate.Text = "计算：";
            txtCalculate.BorderStyle = BorderStyle.None;
            TextBox txtAssessment = new TextBox();
            txtAssessment.Location = new Point(200, tableLayoutPanel.Height + 104);
            txtAssessment.Width = 150;
            txtAssessment.Height = 21;
            txtAssessment.Text = "复核：";
            txtAssessment.BorderStyle = BorderStyle.None;
            TextBox txtDataEnd = new TextBox();
            txtDataEnd.Location = new Point(355, tableLayoutPanel.Height + 104);
            txtDataEnd.Width = 150;
            txtDataEnd.Height = 21;
            txtDataEnd.Text = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            txtDataEnd.BorderStyle = BorderStyle.None;
            Label lbl = new Label();
            lbl.Location = new Point(45, 130 + tableLayoutPanel.Height);

            pnlResult.Controls.Add(txtTitle);
            pnlResult.Controls.Add(txtProjectName);
            pnlResult.Controls.Add(txtDate);
            pnlResult.Controls.Add(txtCalculate);
            pnlResult.Controls.Add(txtAssessment);
            pnlResult.Controls.Add(txtDataEnd);
            pnlResult.Controls.Add(tableLayoutPanel);
            pnlResult.Controls.Add(txtCalculate);
            pnlResult.Controls.Add(txtAssessment);
            pnlResult.Controls.Add(txtDataEnd);
            pnlResult.Controls.Add(lbl);
        }
    }
}
