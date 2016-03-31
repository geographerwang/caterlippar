using System;
using System.Windows.Forms;
using System.Drawing;

namespace AdjustmentAssistant
{
    public class TraverseRecordTable
    {
        public static void DrawTable(Panel panel, int dataCount, ref TableLayoutPanel tableLayoutPanel)
        {
            if (dataCount < 1)
            {
                return;
            }
            //绘制表格总体布局
            tableLayoutPanel.RowCount = dataCount + 3;
            tableLayoutPanel.ColumnCount = 11;
            tableLayoutPanel.Location = new Point(45, 123);
            tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            tableLayoutPanel.Width = 606;
            tableLayoutPanel.Height = tableLayoutPanel.RowCount * 22 + 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 48f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 53f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 53f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 52f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 52f));
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
            }
            //绘制表头
            Label lblPointName = new Label();
            lblPointName.Name = "lblPointName";
            lblPointName.Margin = new Padding(0);
            lblPointName.Width = 48;
            lblPointName.Height = 42;
            lblPointName.Text = "测站名";
            lblPointName.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblPointName, 0, 0);
            Control[] ctrlLblPointName = tableLayoutPanel.Controls.Find("lblPointName", false);
            tableLayoutPanel.SetRowSpan(ctrlLblPointName[0], 2);
            Label lblAimPoint = new Label();
            lblAimPoint.Name = "lblAimPoint";
            lblAimPoint.Margin = new Padding(0);
            lblAimPoint.Width = 48;
            lblAimPoint.Height = 42;
            lblAimPoint.Text = "照准点";
            lblAimPoint.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblAimPoint, 1, 0);
            Control[] ctrlLblAimPoint = tableLayoutPanel.Controls.Find("lblAimPoint", false);
            tableLayoutPanel.SetRowSpan(ctrlLblAimPoint[0], 2);
            Label lblHorizontal = new Label();
            lblHorizontal.Name = "lblHorizontal";
            lblHorizontal.Margin = new Padding(0);
            lblHorizontal.Width = 288;
            lblHorizontal.Height = 21;
            lblHorizontal.Text = "水平角(° ' \")";
            lblHorizontal.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblHorizontal, 2, 0);
            Control[] ctrlLblHorizontal = tableLayoutPanel.Controls.Find("lblHorizontal", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblHorizontal[0], 5);
            Label lblLeft = new Label();
            lblLeft.Margin = new Padding(0);
            lblLeft.Width = 60;
            lblLeft.Height = 21;
            lblLeft.Text = "盘左读数";
            lblLeft.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblLeft, 2, 1);
            Label lblRight = new Label();
            lblRight.Margin = new Padding(0);
            lblRight.Width = 60;
            lblRight.Height = 21;
            lblRight.Text = "盘右读数";
            lblRight.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblRight, 3, 1);
            Label lbl2C = new Label();
            lbl2C.Margin = new Padding(0);
            lbl2C.Width = 60;
            lbl2C.Height = 21;
            lbl2C.Text = "2C(\")";
            lbl2C.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lbl2C, 4, 1);
            Label lblAngle = new Label();
            lblAngle.Margin = new Padding(0);
            lblAngle.Width = 60;
            lblAngle.Height = 21;
            lblAngle.Text = "角值";
            lblAngle.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblAngle, 5, 1);
            Label lblAzimuth = new Label();
            lblAzimuth.Width = 60;
            lblAzimuth.Height = 21;
            lblAzimuth.Text = "方位角";
            lblAzimuth.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblAzimuth, 6, 1);
            Label lblDistance = new Label();
            lblDistance.Name = "lblDistance";
            lblDistance.Margin = new Padding(0);
            lblDistance.Width = 106;
            lblDistance.Height = 21;
            lblDistance.Text = "距离(m)";
            lblDistance.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblDistance, 7, 0);
            Control[] ctrlLblDistance = tableLayoutPanel.Controls.Find("lblDistance", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblDistance[0], 2);
            Label lblHDistance = new Label();
            lblHDistance.Margin = new Padding(0);
            lblHDistance.Width = 53;
            lblHDistance.Height = 21;
            lblHDistance.Text = "平距";
            lblHDistance.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblHDistance, 7, 1);
            Label lblMDistance = new Label();
            lblMDistance.Margin = new Padding(0);
            lblMDistance.Width = 53;
            lblMDistance.Height = 21;
            lblMDistance.Text = "平均距离";
            lblMDistance.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblMDistance, 8, 1);
            Label lblCoordinate = new Label();
            lblCoordinate.Name = "lblCoordinate";
            lblCoordinate.Margin = new Padding(0);
            lblCoordinate.Width = 104;
            lblCoordinate.Height = 21;
            lblCoordinate.Text = "坐标值(m)";
            lblCoordinate.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblCoordinate, 9, 0);
            Control[] ctrlLblCoordinate = tableLayoutPanel.Controls.Find("lblCoordinate", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblCoordinate[0], 2);
            Label lblX = new Label();
            lblX.Margin = new Padding(0);
            lblX.Width = 52;
            lblX.Height = 21;
            lblX.Text = "X(m)";
            lblX.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblX, 9, 1);
            Label lblY = new Label();
            lblY.Margin = new Padding(0);
            lblY.Width = 52;
            lblY.Height = 21;
            lblY.Text = "Y(m)";
            lblY.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblY, 10, 1);

            //表末尾备注
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
            Label lblRemark = new Label();
            lblRemark.Margin = new Padding(0);
            lblRemark.Width = 48;
            lblRemark.Height = 21;
            lblRemark.Text = "备注";
            lblRemark.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblRemark, 0, tableLayoutPanel.RowCount - 1);
        }
    }
}
