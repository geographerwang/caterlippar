using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace AdjustmentAssistant
{
    public class TraverseAdjustment
    {
        internal void DrawTable(Panel pnlResult, List<string> col0, List<string> col1, List<string> col2, List<string> col3, TableLayoutPanel tableLayoutPanel)
        {
            if (col0.Count == 0)
            {
                return;
            }
            //绘制标题记录项
            pnlResult.Controls.Clear();
            TextBox txtTitle = new TextBox();
            txtTitle.Location = new Point(45, 45);
            txtTitle.BorderStyle = BorderStyle.None;
            txtTitle.Width = 570;
            txtTitle.Text = "导线平差计算";
            txtTitle.TextAlign = HorizontalAlignment.Center;
            txtTitle.Font = new Font("宋体", 12, FontStyle.Bold);
            TextBox txtProjectName = new TextBox();
            txtProjectName.Location = new Point(45, 76);
            txtProjectName.BorderStyle = BorderStyle.None;
            txtProjectName.Width = 180;
            txtProjectName.Text = "工程名称：";
            TextBox txtInstrument = new TextBox();
            txtInstrument.Location = new Point(235, 76);
            txtInstrument.BorderStyle = BorderStyle.None;
            txtInstrument.Width = 180;
            txtInstrument.Text = "仪器：";
            TextBox txtWeather = new TextBox();
            txtWeather.Location = new Point(425, 76);
            txtWeather.BorderStyle = BorderStyle.None;
            txtWeather.Width = 180;
            txtWeather.Text = "天气：";
            TextBox txtObserver = new TextBox();
            txtObserver.Location = new Point(45, 97);
            txtObserver.BorderStyle = BorderStyle.None;
            txtObserver.Width = 180;
            txtObserver.Text = "观测者：";
            TextBox txtRecorder = new TextBox();
            txtRecorder.Location = new Point(235, 97);
            txtRecorder.BorderStyle = BorderStyle.None;
            txtRecorder.Width = 180;
            txtRecorder.Text = "记录者：";
            TextBox txtDate = new TextBox();
            txtDate.Location = new Point(425, 97);
            txtDate.BorderStyle = BorderStyle.None;
            txtDate.Width = 180;
            txtDate.Text = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");

            //绘制表格总体布局
            tableLayoutPanel.RowCount = col0.Count + 2;
            tableLayoutPanel.ColumnCount = 8;
            tableLayoutPanel.Location = new Point(45, 120);
            tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            tableLayoutPanel.Width = 580;
            tableLayoutPanel.Height = tableLayoutPanel.RowCount * 22 + 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 30f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95f));

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
            Label lblSimilarCoordinate = new Label();
            lblSimilarCoordinate.Name = "lblSimilarCoordinate";
            lblSimilarCoordinate.Margin = new Padding(0);
            lblSimilarCoordinate.Width = 160;
            lblSimilarCoordinate.Height = 21;
            lblSimilarCoordinate.Text = "近似坐标(m)";
            lblSimilarCoordinate.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSimilarCoordinate, 1, 0);
            Control[] ctrlLblSimilar = tableLayoutPanel.Controls.Find("lblSimilarCoordinate", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblSimilar[0], 2);
            Label lblSimilarX = new Label();
            lblSimilarX.Margin = new Padding(0);
            lblSimilarX.Width = 80;
            lblSimilarX.Height = 21;
            lblSimilarX.Text = "X";
            lblSimilarX.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSimilarX, 1, 1);
            Label lblSimilarY = new Label();
            lblSimilarY.Margin = new Padding(0);
            lblSimilarY.Width = 80;
            lblSimilarY.Height = 21;
            lblSimilarY.Text = "Y";
            lblSimilarY.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSimilarY, 2, 1);
            Label lblSimilarSide = new Label();
            lblSimilarSide.Name = "lblSimilarSide";
            lblSimilarSide.Margin = new Padding(0);
            lblSimilarSide.Width = 60;
            lblSimilarSide.Height = 42;
            lblSimilarSide.Text = "近似边长\n(m)";
            lblSimilarSide.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSimilarSide, 3, 0);
            Control[] ctrlLblSimilarSide = tableLayoutPanel.Controls.Find("lblSimilarSide", false);
            tableLayoutPanel.SetRowSpan(ctrlLblSimilarSide[0], 2);
            Label lblSimilarDirection = new Label();
            lblSimilarDirection.Name = "lblSimilarDirection";
            lblSimilarDirection.Margin = new Padding(0);
            lblSimilarDirection.Width = 60;
            lblSimilarDirection.Height = 42;
            lblSimilarDirection.Text = "近似方位角\n(° ' \")";
            lblSimilarDirection.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSimilarDirection, 4, 0);
            Control[] ctrlLblSimilarDirection = tableLayoutPanel.Controls.Find("lblSimilarDirection", false);
            tableLayoutPanel.SetRowSpan(ctrlLblSimilarDirection[0], 2);
            Label lblAccuracy = new Label();
            lblAccuracy.Name = "lblAccuracy";
            lblAccuracy.Margin = new Padding(0);
            lblAccuracy.Width = 70;
            lblAccuracy.Height = 42;
            lblAccuracy.Text = "点位中误差\n(m)";
            lblAccuracy.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblAccuracy, 5, 0);
            Control[] ctrlLblAccuracy = tableLayoutPanel.Controls.Find("lblAccuracy", false);
            tableLayoutPanel.SetRowSpan(ctrlLblAccuracy[0], 2);
            Label lblCoordinate = new Label();
            lblCoordinate.Name = "lblCoordinate";
            lblCoordinate.Margin = new Padding(0);
            lblCoordinate.Width = 190;
            lblCoordinate.Height = 21;
            lblCoordinate.Text = "坐标平差值(m)";
            lblCoordinate.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblCoordinate, 6, 0);
            Control[] ctrlLblCoordinate = tableLayoutPanel.Controls.Find("lblCoordinate", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblCoordinate[0], 2);
            Label lblCoordinationX = new Label();
            lblCoordinationX.Margin = new Padding(0);
            lblCoordinationX.Width = 95;
            lblCoordinationX.Height = 21;
            lblCoordinationX.Text = "X";
            lblCoordinationX.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblCoordinationX, 6, 1);
            Label lblCoordinationY = new Label();
            lblCoordinationY.Margin = new Padding(0);
            lblCoordinationY.Width = 95;
            lblCoordinationY.Height = 21;
            lblCoordinationY.Text = "Y";
            lblCoordinationY.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblCoordinationY, 7, 1);
            //测站点
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol1 = new TextBox();
                txtBoxCol1.Text = col0[i - 2];
                txtBoxCol1.Margin = new Padding(0);
                txtBoxCol1.Width = 30;
                txtBoxCol1.Height = 21;
                txtBoxCol1.TextAlign = HorizontalAlignment.Center;
                txtBoxCol1.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol1, 0, i);
            }
            //近似坐标X
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol2 = new TextBox();
                txtBoxCol2.Text = col1[i - 2];
                txtBoxCol2.Margin = new Padding(0);
                txtBoxCol2.Width = 80;
                txtBoxCol2.Height = 21;
                txtBoxCol2.TextAlign = HorizontalAlignment.Center;
                txtBoxCol2.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol2, 1, i);
            }
            //近似坐标Y
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol3 = new TextBox();
                txtBoxCol3.Text = col2[i - 2];
                txtBoxCol3.Margin = new Padding(0);
                txtBoxCol3.Width = 80;
                txtBoxCol3.Height = 21;
                txtBoxCol3.TextAlign = HorizontalAlignment.Center;
                txtBoxCol3.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol3, 2, i);
            }
            TextBox txtCalculate = new TextBox();
            txtCalculate.Location = new Point(48, tableLayoutPanel.Height + 125);
            txtCalculate.Width = 180;
            txtCalculate.Height = 21;
            txtCalculate.Text = "计算者：";
            txtCalculate.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtCalculate);
            TextBox txtAssessment = new TextBox();
            txtAssessment.Location = new Point(238, tableLayoutPanel.Height + 125);
            txtAssessment.Width = 180;
            txtAssessment.Height = 21;
            txtAssessment.Text = "审核者：";
            txtAssessment.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtAssessment);
            TextBox txtDataEnd = new TextBox();
            txtDataEnd.Location = new Point(428, tableLayoutPanel.Height + 125);
            txtDataEnd.Width = 180;
            txtDataEnd.Height = 21;
            txtDataEnd.Text = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            txtDataEnd.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtDataEnd);
            Label lbl = new Label();
            lbl.Location = new Point(45, 171 + tableLayoutPanel.Height);

            pnlResult.Controls.Add(txtTitle);
            pnlResult.Controls.Add(txtProjectName);
            pnlResult.Controls.Add(txtInstrument);
            pnlResult.Controls.Add(txtWeather);
            pnlResult.Controls.Add(txtObserver);
            pnlResult.Controls.Add(txtRecorder);
            pnlResult.Controls.Add(txtDate);
            pnlResult.Controls.Add(tableLayoutPanel);
            pnlResult.Controls.Add(lbl);
        }

        internal void DrawTable(Panel pnlResult, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5, List<string> col8, List<string> col14, List<string> col15, TableLayoutPanel tableLayoutPanel)
        {
            if (col0.Count == 0)
            {
                return;
            }
            //绘制标题记录项
            pnlResult.Controls.Clear();
            TextBox txtTitle = new TextBox();
            txtTitle.Location = new Point(45, 45);
            txtTitle.BorderStyle = BorderStyle.None;
            txtTitle.Width = 570;
            txtTitle.Text = "导线平差计算";
            txtTitle.TextAlign = HorizontalAlignment.Center;
            txtTitle.Font = new Font("宋体", 12, FontStyle.Bold);
            TextBox txtProjectName = new TextBox();
            txtProjectName.Location = new Point(45, 76);
            txtProjectName.BorderStyle = BorderStyle.None;
            txtProjectName.Width = 180;
            txtProjectName.Text = "工程名称：";
            TextBox txtInstrument = new TextBox();
            txtInstrument.Location = new Point(235, 76);
            txtInstrument.BorderStyle = BorderStyle.None;
            txtInstrument.Width = 180;
            txtInstrument.Text = "仪器：";
            TextBox txtWeather = new TextBox();
            txtWeather.Location = new Point(425, 76);
            txtWeather.BorderStyle = BorderStyle.None;
            txtWeather.Width = 180;
            txtWeather.Text = "天气：";
            TextBox txtObserver = new TextBox();
            txtObserver.Location = new Point(45, 97);
            txtObserver.BorderStyle = BorderStyle.None;
            txtObserver.Width = 180;
            txtObserver.Text = "观测者：";
            TextBox txtRecorder = new TextBox();
            txtRecorder.Location = new Point(235, 97);
            txtRecorder.BorderStyle = BorderStyle.None;
            txtRecorder.Width = 180;
            txtRecorder.Text = "记录者：";
            TextBox txtDate = new TextBox();
            txtDate.Location = new Point(425, 97);
            txtDate.BorderStyle = BorderStyle.None;
            txtDate.Width = 180;
            txtDate.Text = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");

            //绘制表格总体布局
            tableLayoutPanel.RowCount = col0.Count + 2;
            tableLayoutPanel.ColumnCount = 8;
            tableLayoutPanel.Location = new Point(45, 120);
            tableLayoutPanel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
            tableLayoutPanel.Width = 580;
            tableLayoutPanel.Height = tableLayoutPanel.RowCount * 22 + 1;
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 30f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 60f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95f));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 95f));

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
            Label lblSimilarCoordinate = new Label();
            lblSimilarCoordinate.Name = "lblSimilarCoordinate";
            lblSimilarCoordinate.Margin = new Padding(0);
            lblSimilarCoordinate.Width = 160;
            lblSimilarCoordinate.Height = 21;
            lblSimilarCoordinate.Text = "近似坐标(m)";
            lblSimilarCoordinate.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSimilarCoordinate, 1, 0);
            Control[] ctrlLblSimilar = tableLayoutPanel.Controls.Find("lblSimilarCoordinate", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblSimilar[0], 2);
            Label lblSimilarX = new Label();
            lblSimilarX.Margin = new Padding(0);
            lblSimilarX.Width = 80;
            lblSimilarX.Height = 21;
            lblSimilarX.Text = "X";
            lblSimilarX.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSimilarX, 1, 1);
            Label lblSimilarY = new Label();
            lblSimilarY.Margin = new Padding(0);
            lblSimilarY.Width = 80;
            lblSimilarY.Height = 21;
            lblSimilarY.Text = "Y";
            lblSimilarY.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSimilarY, 2, 1);
            Label lblSimilarSide = new Label();
            lblSimilarSide.Name = "lblSimilarSide";
            lblSimilarSide.Margin = new Padding(0);
            lblSimilarSide.Width = 60;
            lblSimilarSide.Height = 42;
            lblSimilarSide.Text = "近似边长\n(m)";
            lblSimilarSide.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSimilarSide, 3, 0);
            Control[] ctrlLblSimilarSide = tableLayoutPanel.Controls.Find("lblSimilarSide", false);
            tableLayoutPanel.SetRowSpan(ctrlLblSimilarSide[0], 2);
            Label lblSimilarDirection = new Label();
            lblSimilarDirection.Name = "lblSimilarDirection";
            lblSimilarDirection.Margin = new Padding(0);
            lblSimilarDirection.Width = 60;
            lblSimilarDirection.Height = 42;
            lblSimilarDirection.Text = "近似方位角\n(° ' \")";
            lblSimilarDirection.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblSimilarDirection, 4, 0);
            Control[] ctrlLblSimilarDirection = tableLayoutPanel.Controls.Find("lblSimilarDirection", false);
            tableLayoutPanel.SetRowSpan(ctrlLblSimilarDirection[0], 2);
            Label lblAccuracy = new Label();
            lblAccuracy.Name = "lblAccuracy";
            lblAccuracy.Margin = new Padding(0);
            lblAccuracy.Width = 70;
            lblAccuracy.Height = 42;
            lblAccuracy.Text = "点位中误差\n(m)";
            lblAccuracy.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblAccuracy, 5, 0);
            Control[] ctrlLblAccuracy = tableLayoutPanel.Controls.Find("lblAccuracy", false);
            tableLayoutPanel.SetRowSpan(ctrlLblAccuracy[0], 2);
            Label lblCoordinate = new Label();
            lblCoordinate.Name = "lblCoordinate";
            lblCoordinate.Margin = new Padding(0);
            lblCoordinate.Width = 190;
            lblCoordinate.Height = 21;
            lblCoordinate.Text = "坐标平差值(m)";
            lblCoordinate.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblCoordinate, 6, 0);
            Control[] ctrlLblCoordinate = tableLayoutPanel.Controls.Find("lblCoordinate", false);
            tableLayoutPanel.SetColumnSpan(ctrlLblCoordinate[0], 2);
            Label lblCoordinationX = new Label();
            lblCoordinationX.Margin = new Padding(0);
            lblCoordinationX.Width = 95;
            lblCoordinationX.Height = 21;
            lblCoordinationX.Text = "X";
            lblCoordinationX.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblCoordinationX, 6, 1);
            Label lblCoordinationY = new Label();
            lblCoordinationY.Margin = new Padding(0);
            lblCoordinationY.Width = 95;
            lblCoordinationY.Height = 21;
            lblCoordinationY.Text = "Y";
            lblCoordinationY.TextAlign = ContentAlignment.MiddleCenter;
            tableLayoutPanel.Controls.Add(lblCoordinationY, 7, 1);
            //测站点
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol1 = new TextBox();
                txtBoxCol1.Text = col0[i - 2];
                txtBoxCol1.Margin = new Padding(0);
                txtBoxCol1.Width = 30;
                txtBoxCol1.Height = 21;
                txtBoxCol1.TextAlign = HorizontalAlignment.Center;
                txtBoxCol1.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol1, 0, i);
            }
            //近似坐标X
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol2 = new TextBox();
                txtBoxCol2.Text = col1[i - 2];
                txtBoxCol2.Margin = new Padding(0);
                txtBoxCol2.Width = 80;
                txtBoxCol2.Height = 21;
                txtBoxCol2.TextAlign = HorizontalAlignment.Center;
                txtBoxCol2.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol2, 1, i);
            }
            //近似坐标Y
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol3 = new TextBox();
                txtBoxCol3.Text = col2[i - 2];
                txtBoxCol3.Margin = new Padding(0);
                txtBoxCol3.Width = 80;
                txtBoxCol3.Height = 21;
                txtBoxCol3.TextAlign = HorizontalAlignment.Center;
                txtBoxCol3.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol3, 2, i);
            }
            for (int i = 0; i < col0.Count - 1; i++)
            {
                TextBox txtSide = new TextBox();
                txtSide.Text = col4[i];
                txtSide.Margin = new Padding(0);
                txtSide.Width = 60;
                txtSide.Height = 21;
                txtSide.TextAlign = HorizontalAlignment.Center;
                txtSide.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtSide, 3, i + 2);
            }
            for (int i = 0; i < col0.Count - 1; i++)
            {
                TextBox txtAngle = new TextBox();
                txtAngle.Text = col5[i];
                txtAngle.Margin = new Padding(0);
                txtAngle.Width = 60;
                txtAngle.Height = 21;
                txtAngle.TextAlign = HorizontalAlignment.Center;
                txtAngle.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtAngle, 4, i + 2);
            }
            for (int i = 0; i < col0.Count; i++)
            {
                TextBox txtErr = new TextBox();
                txtErr.Text = col8[i];
                txtErr.Margin = new Padding(0);
                txtErr.Width = 70;
                txtErr.Height = 21;
                txtErr.TextAlign = HorizontalAlignment.Center;
                txtErr.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtErr, 5, i + 2);
                TextBox txtX = new TextBox();
                txtX.Text = col14[i];
                txtX.Margin = new Padding(0);
                txtX.Width = 95;
                txtX.Height = 21;
                txtX.TextAlign = HorizontalAlignment.Center;
                txtX.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtX, 6, i + 2);
                TextBox txtY = new TextBox();
                txtY.Text = col15[i];
                txtY.Margin = new Padding(0);
                txtY.Width = 95;
                txtY.Height = 21;
                txtY.TextAlign = HorizontalAlignment.Center;
                txtY.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtY, 7, i + 2);
            }
            TextBox txtCalculate = new TextBox();
            txtCalculate.Location = new Point(48, tableLayoutPanel.Height + 125);
            txtCalculate.Width = 180;
            txtCalculate.Height = 21;
            txtCalculate.Text = "计算者：";
            txtCalculate.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtCalculate);
            TextBox txtAssessment = new TextBox();
            txtAssessment.Location = new Point(238, tableLayoutPanel.Height + 125);
            txtAssessment.Width = 180;
            txtAssessment.Height = 21;
            txtAssessment.Text = "审核者：";
            txtAssessment.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtAssessment);
            TextBox txtDataEnd = new TextBox();
            txtDataEnd.Location = new Point(428, tableLayoutPanel.Height + 125);
            txtDataEnd.Width = 180;
            txtDataEnd.Height = 21;
            txtDataEnd.Text = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            txtDataEnd.BorderStyle = BorderStyle.None;
            pnlResult.Controls.Add(txtDataEnd);
            Label lbl = new Label();
            lbl.Location = new Point(45, 171 + tableLayoutPanel.Height);

            pnlResult.Controls.Add(txtTitle);
            pnlResult.Controls.Add(txtProjectName);
            pnlResult.Controls.Add(txtInstrument);
            pnlResult.Controls.Add(txtWeather);
            pnlResult.Controls.Add(txtObserver);
            pnlResult.Controls.Add(txtRecorder);
            pnlResult.Controls.Add(txtDate);
            pnlResult.Controls.Add(tableLayoutPanel);
            pnlResult.Controls.Add(lbl);
        }
    }
}
