using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace AdjustmentAssistant
{
    class ResultTable
    {
        internal void DrawTable(string filePath, Panel pnlResult, TableLayoutPanel tableLayoutPanel, List<string> col0, List<string> col14, List<string> col15, List<string> col16, List<string> col17, List<string> col18)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return;
            }
            //绘制标题记录项
            pnlResult.Controls.Clear();
            TextBox txtTitle = new TextBox();
            txtTitle.Location = new Point(45, 45);
            txtTitle.BorderStyle = BorderStyle.None;
            txtTitle.Width = 580;
            txtTitle.Text = "平差计算成果表";
            txtTitle.TextAlign = HorizontalAlignment.Center;
            txtTitle.Font = new Font("宋体", 12, FontStyle.Bold);
            TextBox txtProjectName = new TextBox();
            txtProjectName.Location = new Point(45, 76);
            txtProjectName.BorderStyle = BorderStyle.None;
            txtProjectName.Width = 200;
            txtProjectName.Text = "工程名称：";
            TextBox txtCalculate = new TextBox();
            txtCalculate.Location = new Point(425, 76);
            txtCalculate.BorderStyle = BorderStyle.None;
            txtCalculate.Width = 200;
            txtCalculate.Text = "计算者：";
            //绘制表格总体布局
            tableLayoutPanel.RowCount = col0.Count + 2;
            tableLayoutPanel.ColumnCount = 9;
            tableLayoutPanel.Location = new Point(45, 98);
            tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            tableLayoutPanel.Width = 590;
            tableLayoutPanel.Height = tableLayoutPanel.RowCount * 22 + 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 30f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 30f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50f));
            Label lblPointName = new Label();
            lblPointName.Name = "lblPointName";
            lblPointName.Margin = new Padding(0);
            lblPointName.Width = 30;
            lblPointName.Height = 42;
            lblPointName.Text = "点名";
            lblPointName.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblPointName, 0, 0);
            Control[] ctrlLblPointName = tableLayoutPanel.Controls.Find("lblPointName", false);
            tableLayoutPanel.SetRowSpan(ctrlLblPointName[0], 2);
            Label lblCoordinate = new Label();
            lblCoordinate.Name = "lblCoordinate";
            lblCoordinate.Margin = new Padding(0);
            lblCoordinate.Width = 160;
            lblCoordinate.Height = 21;
            lblCoordinate.Text = "坐标(m)";
            lblCoordinate.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblCoordinate, 1, 0);
            Control[] ctrlLblCoordinate = tableLayoutPanel.Controls.Find("lblCoordinate", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblCoordinate[0], 2);
            Label lblX = new Label();
            lblX.Margin = new Padding(0);
            lblX.Width = 80;
            lblX.Height = 21;
            lblX.Text = "X";
            lblX.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblX, 1, 1);
            Label lblY = new Label();
            lblY.Margin = new Padding(0);
            lblY.Width = 80;
            lblY.Height = 21;
            lblY.Text = "Y";
            lblY.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblY, 2, 1);
            Label lblDem = new Label();
            lblDem.Name = "lblDem";
            lblDem.Margin = new Padding(0);
            lblDem.Width = 50;
            lblDem.Height = 42;
            lblDem.Text = "高程(m)";
            lblDem.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblDem, 3, 0);
            Control[] ctrlLblDem = tableLayoutPanel.Controls.Find("lblDem", false);
            tableLayoutPanel.SetRowSpan(ctrlLblDem[0], 2);
            Label lblAngle = new Label();
            lblAngle.Name = "lblAngle";
            lblAngle.Margin = new Padding(0);
            lblAngle.Width = 100;
            lblAngle.Height = 42;
            lblAngle.Text = "角度平差值\n(° ' \")";
            lblAngle.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblAngle, 4, 0);
            Control[] ctrlLblAngle = tableLayoutPanel.Controls.Find("lblAngle", false);
            tableLayoutPanel.SetRowSpan(ctrlLblAngle[0], 2);
            Label lblToP = new Label();
            lblToP.Name = "lblToP";
            lblToP.Margin = new Padding(0);
            lblToP.Width = 30;
            lblToP.Height = 42;
            lblToP.Text = "至点";
            lblToP.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblToP, 5, 0);
            Control[] ctrlLblToP = tableLayoutPanel.Controls.Find("lblToP", false);
            tableLayoutPanel.SetRowSpan(ctrlLblToP[0], 2);
            Label lblDir = new Label();
            lblDir.Name = "lblDir";
            lblDir.Margin = new Padding(0);
            lblDir.Width = 100;
            lblDir.Height = 42;
            lblDir.Text = "方位角\n(° ' \")";
            lblDir.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblDir, 6, 0);
            Control[] ctrlLblDir = tableLayoutPanel.Controls.Find("lblDir", false);
            tableLayoutPanel.SetRowSpan(ctrlLblDir[0], 2);
            Label lblSide = new Label();
            lblSide.Name = "lblSide";
            lblSide.Margin = new Padding(0);
            lblSide.Width = 60;
            lblSide.Height = 42;
            lblSide.Text = "边长平差值(m)";
            lblSide.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSide, 7, 0);
            Control[] ctrlLblSide = tableLayoutPanel.Controls.Find("lblSide", false);
            tableLayoutPanel.SetRowSpan(ctrlLblSide[0], 2);
            Label lblDemAdj = new Label();
            lblDemAdj.Name = "lblDemAdj";
            lblDemAdj.Margin = new Padding(0);
            lblDemAdj.Width = 50;
            lblDemAdj.Height = 42;
            lblDemAdj.Text = "高差平差值(m)";
            lblDemAdj.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblDemAdj, 8, 0);
            Control[] ctrlLblDemAdj = tableLayoutPanel.Controls.Find("lblDemAdj", false);
            tableLayoutPanel.SetRowSpan(ctrlLblDemAdj[0], 2);
            //测站点
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol0 = new TextBox();
                txtBoxCol0.Text = col0[i - 2];
                txtBoxCol0.Margin = new Padding(0);
                txtBoxCol0.Width = 30;
                txtBoxCol0.Height = 21;
                txtBoxCol0.TextAlign = HorizontalAlignment.Center;
                txtBoxCol0.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol0, 0, i);
            }
            //坐标X
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol1 = new TextBox();
                txtBoxCol1.Text = col14[i - 2];
                txtBoxCol1.Margin = new Padding(0);
                txtBoxCol1.Width = 80;
                txtBoxCol1.Height = 21;
                txtBoxCol1.TextAlign = HorizontalAlignment.Center;
                txtBoxCol1.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol1, 1, i);
            }
            //坐标Y
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol2 = new TextBox();
                txtBoxCol2.Text = col15[i - 2];
                txtBoxCol2.Margin = new Padding(0);
                txtBoxCol2.Width = 80;
                txtBoxCol2.Height = 21;
                txtBoxCol2.TextAlign = HorizontalAlignment.Center;
                txtBoxCol2.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol2, 2, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol3 = new TextBox();
                txtBoxCol3.Text = "0";
                txtBoxCol3.Margin = new Padding(0);
                txtBoxCol3.Width = 50;
                txtBoxCol3.Height = 21;
                txtBoxCol3.TextAlign = HorizontalAlignment.Center;
                txtBoxCol3.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol3, 3, i);
            }
            for (int i = 3; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol4 = new TextBox();
                txtBoxCol4.Text = col18[i - 3];
                txtBoxCol4.Margin = new Padding(0);
                txtBoxCol4.Width = 100;
                txtBoxCol4.Height = 21;
                txtBoxCol4.TextAlign = HorizontalAlignment.Center;
                txtBoxCol4.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol4, 4, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol5 = new TextBox();
                txtBoxCol5.Text = col0[i - 1];
                txtBoxCol5.Margin = new Padding(0);
                txtBoxCol5.Width = 30;
                txtBoxCol5.Height = 21;
                txtBoxCol5.TextAlign = HorizontalAlignment.Center;
                txtBoxCol5.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol5, 5, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol6 = new TextBox();
                txtBoxCol6.Text = col16[i - 2];
                txtBoxCol6.Margin = new Padding(0);
                txtBoxCol6.Width = 100;
                txtBoxCol6.Height = 21;
                txtBoxCol6.TextAlign = HorizontalAlignment.Center;
                txtBoxCol6.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol6, 6, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol7 = new TextBox();
                txtBoxCol7.Text = col17[i - 2];
                txtBoxCol7.Margin = new Padding(0);
                txtBoxCol7.Width = 60;
                txtBoxCol7.Height = 21;
                txtBoxCol7.TextAlign = HorizontalAlignment.Center;
                txtBoxCol7.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol7, 7, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol8 = new TextBox();
                txtBoxCol8.Text = "0";
                txtBoxCol8.Margin = new Padding(0);
                txtBoxCol8.Width = 50;
                txtBoxCol8.Height = 21;
                txtBoxCol8.TextAlign = HorizontalAlignment.Center;
                txtBoxCol8.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol8, 8, i);
            }
            TextBox txtAssessment = new TextBox();
            txtAssessment.Location = new Point(48, tableLayoutPanel.Height + 103);
            txtAssessment.Width = 200;
            txtAssessment.Height = 21;
            txtAssessment.Text = "校核者：";
            txtAssessment.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtAssessment);
            TextBox txtDataEnd = new TextBox();
            txtDataEnd.Location = new Point(425, tableLayoutPanel.Height + 103);
            txtDataEnd.Width = 200;
            txtDataEnd.Height = 21;
            txtDataEnd.Text = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            txtDataEnd.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtDataEnd);
            Label lbl = new Label();
            lbl.Location = new Point(45, 149 + tableLayoutPanel.Height);

            pnlResult.Controls.Add(txtTitle);
            pnlResult.Controls.Add(txtProjectName);
            pnlResult.Controls.Add(txtCalculate);
            pnlResult.Controls.Add(tableLayoutPanel);
            pnlResult.Controls.Add(lbl);
        }
    }
}
