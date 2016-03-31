using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace AdjustmentAssistant
{
    class AccuracyTable
    {
        internal void DrawTable(string filePath, Panel pnlResult, TableLayoutPanel tableLayoutPanel, double unitError, List<string> col0, List<string> col6, List<string> col7, List<string> col8, List<string> col9, List<string> col10, List<string> col11, List<string> col12, List<string> col13)
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
            txtTitle.Text = "精度评定表";
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
            tableLayoutPanel.RowCount = col0.Count + 3;
            tableLayoutPanel.ColumnCount = 11;
            tableLayoutPanel.Location = new Point(45, 98);
            tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            tableLayoutPanel.Width = 592;
            tableLayoutPanel.Height = tableLayoutPanel.RowCount * 22 + 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 40f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            Label lblPointName = new Label();
            lblPointName.Name = "lblPointName";
            lblPointName.Margin = new Padding(0);
            lblPointName.Width = 40;
            lblPointName.Height = 42;
            lblPointName.Text = "点名";
            lblPointName.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblPointName, 0, 0);
            Control[] ctrlLblPointName = tableLayoutPanel.Controls.Find("lblPointName", false);
            tableLayoutPanel.SetRowSpan(ctrlLblPointName[0], 2);
            Label lblPoint = new Label();
            lblPoint.Name = "lblPoint";
            lblPoint.Margin = new Padding(0);
            lblPoint.Width = 150;
            lblPoint.Height = 21;
            lblPoint.Text = "点位中误差(m)";
            lblPoint.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblPoint, 1, 0);
            Control[] ctrlLblPoint = tableLayoutPanel.Controls.Find("lblPoint", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblPoint[0], 3);
            Label lblMx = new Label();
            lblMx.Margin = new Padding(0);
            lblMx.Width = 50;
            lblMx.Height = 21;
            lblMx.Text = "Mx";
            lblMx.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblMx, 1, 1);
            Label lblMy = new Label();
            lblMy.Margin = new Padding(0);
            lblMy.Width = 50;
            lblMy.Height = 21;
            lblMy.Text = "My";
            lblMy.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblMy, 2, 1);
            Label lblM = new Label();
            lblM.Margin = new Padding(0);
            lblM.Width = 50;
            lblM.Height = 21;
            lblM.Text = "M";
            lblM.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblM, 3, 1);
            Label lblEllipses = new Label();
            lblEllipses.Name = "lblEllipses";
            lblEllipses.Margin = new Padding(0);
            lblEllipses.Width = 150;
            lblEllipses.Height = 21;
            lblEllipses.Text = "误差椭圆";
            lblEllipses.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblEllipses, 4, 0);
            Control[] ctrlLblEllipses = tableLayoutPanel.Controls.Find("lblEllipses", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblEllipses[0], 3);
            Label lblA = new Label();
            lblA.Margin = new Padding(0);
            lblA.Width = 50;
            lblA.Height = 21;
            lblA.Text = "A(m)";
            lblA.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblA, 4, 1);
            Label lblB = new Label();
            lblB.Margin = new Padding(0);
            lblB.Width = 50;
            lblB.Height = 21;
            lblB.Text = "B(m)";
            lblB.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblB, 5, 1);
            Label lblF = new Label();
            lblF.Margin = new Padding(0);
            lblF.Width = 80;
            lblF.Height = 21;
            lblF.Text = "F(° ' \")";
            lblF.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblF, 6, 1);
            Label lblDem = new Label();
            lblDem.Name = "lblDem";
            lblDem.Margin = new Padding(0);
            lblDem.Width = 50;
            lblDem.Height = 42;
            lblDem.Text = "高程中误差(m)";
            lblDem.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblDem, 7, 0);
            Control[] ctrlLblDem = tableLayoutPanel.Controls.Find("lblDem", false);
            tableLayoutPanel.SetRowSpan(ctrlLblDem[0], 2);
            Label lblToP = new Label();
            lblToP.Name = "lblToP";
            lblToP.Margin = new Padding(0);
            lblToP.Width = 40;
            lblToP.Height = 42;
            lblToP.Text = "至点";
            lblToP.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblToP, 8, 0);
            Control[] ctrlLblToP = tableLayoutPanel.Controls.Find("lblToP", false);
            tableLayoutPanel.SetRowSpan(ctrlLblToP[0], 2);
            Label lblDir = new Label();
            lblDir.Name = "lblDir";
            lblDir.Margin = new Padding(0);
            lblDir.Width = 60;
            lblDir.Height = 42;
            lblDir.Text = "方位角中误差(\")";
            lblDir.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblDir, 9, 0);
            Control[] ctrlLblDir = tableLayoutPanel.Controls.Find("lblDir", false);
            tableLayoutPanel.SetRowSpan(ctrlLblDir[0], 2);
            Label lblSide = new Label();
            lblSide.Name = "lblSide";
            lblSide.Margin = new Padding(0);
            lblSide.Width = 60;
            lblSide.Height = 42;
            lblSide.Text = "边长中误差(m)";
            lblSide.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSide, 10, 0);
            Control[] ctrlLblSide = tableLayoutPanel.Controls.Find("lblSide", false);
            tableLayoutPanel.SetRowSpan(ctrlLblSide[0], 2);
            //测站点
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol0 = new TextBox();
                txtBoxCol0.Text = col0[i - 2];
                txtBoxCol0.Margin = new Padding(0);
                txtBoxCol0.Width = 40;
                txtBoxCol0.Height = 21;
                txtBoxCol0.TextAlign = HorizontalAlignment.Center;
                txtBoxCol0.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol0, 0, i);
            }
            //Mx
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol1 = new TextBox();
                txtBoxCol1.Text = col6[i - 2];
                txtBoxCol1.Margin = new Padding(0);
                txtBoxCol1.Width = 50;
                txtBoxCol1.Height = 21;
                txtBoxCol1.TextAlign = HorizontalAlignment.Center;
                txtBoxCol1.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol1, 1, i);
            }
            //My
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol2 = new TextBox();
                txtBoxCol2.Text = col7[i - 2];
                txtBoxCol2.Margin = new Padding(0);
                txtBoxCol2.Width = 50;
                txtBoxCol2.Height = 21;
                txtBoxCol2.TextAlign = HorizontalAlignment.Center;
                txtBoxCol2.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol2, 2, i);
            }
            //M
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol3 = new TextBox();
                txtBoxCol3.Text = col8[i - 2];
                txtBoxCol3.Margin = new Padding(0);
                txtBoxCol3.Width = 50;
                txtBoxCol3.Height = 21;
                txtBoxCol3.TextAlign = HorizontalAlignment.Center;
                txtBoxCol3.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol3, 3, i);
            }
            //A
            for (int i = 4; i < tableLayoutPanel.RowCount - 3; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol4 = new TextBox();
                txtBoxCol4.Text = col11[i - 4];
                txtBoxCol4.Margin = new Padding(0);
                txtBoxCol4.Width = 50;
                txtBoxCol4.Height = 21;
                txtBoxCol4.TextAlign = HorizontalAlignment.Center;
                txtBoxCol4.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol4, 4, i);
            }
            //B
            for (int i = 4; i < tableLayoutPanel.RowCount - 3; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol5 = new TextBox();
                txtBoxCol5.Text = col12[i - 4];
                txtBoxCol5.Margin = new Padding(0);
                txtBoxCol5.Width = 50;
                txtBoxCol5.Height = 21;
                txtBoxCol5.TextAlign = HorizontalAlignment.Center;
                txtBoxCol5.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol5, 5, i);
            }
            //F
            for (int i = 4; i < tableLayoutPanel.RowCount - 3; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol6 = new TextBox();
                txtBoxCol6.Text = col13[i - 4];
                txtBoxCol6.Margin = new Padding(0);
                txtBoxCol6.Width = 80;
                txtBoxCol6.Height = 21;
                txtBoxCol6.TextAlign = HorizontalAlignment.Center;
                txtBoxCol6.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol6, 6, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount - 2; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol7 = new TextBox();
                txtBoxCol7.Text = "0";
                txtBoxCol7.Margin = new Padding(0);
                txtBoxCol7.Width = 50;
                txtBoxCol7.Height = 21;
                txtBoxCol7.TextAlign = HorizontalAlignment.Center;
                txtBoxCol7.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol7, 7, i);
            }
            //ToPoint
            for (int i = 2; i < tableLayoutPanel.RowCount - 2; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol7 = new TextBox();
                txtBoxCol7.Text = col0[i - 1];
                txtBoxCol7.Margin = new Padding(0);
                txtBoxCol7.Width = 40;
                txtBoxCol7.Height = 21;
                txtBoxCol7.TextAlign = HorizontalAlignment.Center;
                txtBoxCol7.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol7, 8, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount - 2; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol8 = new TextBox();
                txtBoxCol8.Text = col10[i - 2];
                txtBoxCol8.Margin = new Padding(0);
                txtBoxCol8.Width = 60;
                txtBoxCol8.Height = 21;
                txtBoxCol8.TextAlign = HorizontalAlignment.Center;
                txtBoxCol8.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol8, 9, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount - 2; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol9 = new TextBox();
                txtBoxCol9.Text = col9[i - 2];
                txtBoxCol9.Margin = new Padding(0);
                txtBoxCol9.Width = 60;
                txtBoxCol9.Height = 21;
                txtBoxCol9.TextAlign = HorizontalAlignment.Center;
                txtBoxCol9.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol9, 10, i);
            }
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
            TextBox txtMark = new TextBox();
            txtMark.Text = "备注";
            txtMark.Margin = new Padding(0);
            txtMark.Width = 60;
            txtMark.Height = 21;
            txtMark.TextAlign = HorizontalAlignment.Center;
            txtMark.BorderStyle = BorderStyle.None;
            tableLayoutPanel.Controls.Add(txtMark, 0, tableLayoutPanel.RowCount - 1);
            TextBox txtMarkCont = new TextBox();
            txtMarkCont.Name = "txtMarkCont";
            txtMarkCont.Text = "单位权中误差 = " + unitError;
            txtMarkCont.Multiline = true;
            txtMarkCont.Margin = new Padding(0);
            txtMarkCont.Width = 520;
            txtMarkCont.Height = 21;
            txtMarkCont.TextAlign = HorizontalAlignment.Center;
            txtMarkCont.BorderStyle = BorderStyle.None;
            tableLayoutPanel.Controls.Add(txtMarkCont, 1, tableLayoutPanel.RowCount - 1);
            Control[] ctrlLblMarkCont = tableLayoutPanel.Controls.Find("txtMarkCont", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblMarkCont[0], 10);

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
