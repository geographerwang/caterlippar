using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace AdjustmentAssistant
{
    class ControlPointTable
    {
        internal void DrawTable(string filePath, Panel pnlResult, TableLayoutPanel tableLayoutPanel, List<string> col0, List<string> col14, List<string> col15)
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
            txtTitle.Width = 340;
            txtTitle.Text = "控制点成果表";
            txtTitle.TextAlign = HorizontalAlignment.Center;
            txtTitle.Font = new Font("宋体", 12, FontStyle.Bold);
            TextBox txtProjectName = new TextBox();
            txtProjectName.Location = new Point(45, 76);
            txtProjectName.BorderStyle = BorderStyle.None;
            txtProjectName.Width = 150;
            txtProjectName.Text = "工程名称：";
            TextBox txtCalculate = new TextBox();
            txtCalculate.Location = new Point(265, 76);
            txtCalculate.BorderStyle = BorderStyle.None;
            txtCalculate.Width = 150;
            txtCalculate.Text = "计算者：";
            //绘制表格总体布局
            tableLayoutPanel.RowCount = col0.Count + 2;
            tableLayoutPanel.ColumnCount = 4;
            tableLayoutPanel.Location = new Point(45, 98);
            tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            tableLayoutPanel.Width = 345;
            tableLayoutPanel.Height = tableLayoutPanel.RowCount * 22 + 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            Label lblPointName = new Label();
            lblPointName.Name = "lblPointName";
            lblPointName.Margin = new Padding(0);
            lblPointName.Width = 60;
            lblPointName.Height = 42;
            lblPointName.Text = "点名";
            lblPointName.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblPointName, 0, 0);
            Control[] ctrlLblPointName = tableLayoutPanel.Controls.Find("lblPointName", false);
            tableLayoutPanel.SetRowSpan(ctrlLblPointName[0], 2);
            Label lblCoordinate = new Label();
            lblCoordinate.Name = "lblCoordinate";
            lblCoordinate.Margin = new Padding(0);
            lblCoordinate.Width = 200;
            lblCoordinate.Height = 21;
            lblCoordinate.Text = "坐标(m)";
            lblCoordinate.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblCoordinate, 1, 0);
            Control[] ctrlLblCoordinate = tableLayoutPanel.Controls.Find("lblCoordinate", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblCoordinate[0], 2);
            Label lblX = new Label();
            lblX.Margin = new Padding(0);
            lblX.Width = 100;
            lblX.Height = 21;
            lblX.Text = "X";
            lblX.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblX, 1, 1);
            Label lblY = new Label();
            lblY.Margin = new Padding(0);
            lblY.Width = 100;
            lblY.Height = 21;
            lblY.Text = "Y";
            lblY.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblY, 2, 1);
            Label lblDem = new Label();
            lblDem.Name = "lblDem";
            lblDem.Margin = new Padding(0);
            lblDem.Width = 80;
            lblDem.Height = 42;
            lblDem.Text = "高程(m)";
            lblDem.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblDem, 3, 0);
            Control[] ctrlLblDem = tableLayoutPanel.Controls.Find("lblDem", false);
            tableLayoutPanel.SetRowSpan(ctrlLblDem[0], 2);
            //测站点
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol0 = new TextBox();
                txtBoxCol0.Text = col0[i - 2];
                txtBoxCol0.Margin = new Padding(0);
                txtBoxCol0.Width = 60;
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
                txtBoxCol1.Width = 100;
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
                txtBoxCol2.Width = 100;
                txtBoxCol2.Height = 21;
                txtBoxCol2.TextAlign = HorizontalAlignment.Center;
                txtBoxCol2.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol2, 2, i);
            }
            //高程（待增加）
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol3 = new TextBox();
                txtBoxCol3.Text = "0";
                txtBoxCol3.Margin = new Padding(0);
                txtBoxCol3.Width = 80;
                txtBoxCol3.Height = 21;
                txtBoxCol3.TextAlign = HorizontalAlignment.Center;
                txtBoxCol3.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol3, 3, i);
            }

            TextBox txtAssessment = new TextBox();
            txtAssessment.Location = new Point(48, tableLayoutPanel.Height + 103);
            txtAssessment.Width = 150;
            txtAssessment.Height = 21;
            txtAssessment.Text = "校核者：";
            txtAssessment.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtAssessment);
            TextBox txtDataEnd = new TextBox();
            txtDataEnd.Location = new Point(265, tableLayoutPanel.Height + 103);
            txtDataEnd.Width = 150;
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
