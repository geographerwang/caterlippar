using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace AdjustmentAssistant
{
    public class LevelingRecordTable
    {
        internal static void DrawTable(Panel pnlResult, int dataCount, ref TableLayoutPanel tableLayoutPanel)
        {
            if (dataCount < 1)
            {
                return;
            }
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
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
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
            pnlResult.Controls.Add(tableLayoutPanel);
        }
    }
}
