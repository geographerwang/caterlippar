using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;

namespace AdjustmentAssistant
{
    public class InputTraverseRecordTable
    {
        public static void GetData(Panel panel, int dataCount, int backCount, DataType.Data ApproximateDataType, ref TableLayoutPanel tableLayoutPanel)
        {
            if (dataCount < 1)
            {
                return;
            }
            panel.Controls.Clear();
            //绘制标题记录项
            panel.Controls.Clear();
            TextBox txtTitle = new TextBox();
            txtTitle.Location = new Point(45, 45);
            txtTitle.BorderStyle = BorderStyle.None;
            txtTitle.Width = 606;
            txtTitle.Text = "导线观测记录";
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
            TextBox txtCalculate = new TextBox();
            txtCalculate.Location = new Point(48, tableLayoutPanel.Height + 125);
            txtCalculate.Width = 195;
            txtCalculate.Height = 21;
            txtCalculate.Text = "计算者：";
            txtCalculate.BorderStyle = BorderStyle.None;
            TextBox txtAssessment = new TextBox();
            txtAssessment.Location = new Point(246, tableLayoutPanel.Height + 125);
            txtAssessment.Width = 195;
            txtAssessment.Height = 21;
            txtAssessment.Text = "审核者：";
            txtAssessment.BorderStyle = BorderStyle.None;
            TextBox txtDataEnd = new TextBox();
            txtDataEnd.Location = new Point(442, tableLayoutPanel.Height + 125);
            txtDataEnd.Width = 195;
            txtDataEnd.Height = 21;
            txtDataEnd.Text = "日期：" + DateTime.Now.ToString("yyyy-MM-dd");
            txtDataEnd.BorderStyle = BorderStyle.None;
            Label lbl = new Label();
            lbl.Location = new Point(45, 171 + tableLayoutPanel.Height);
            //测站点
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i += (2 * backCount))
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol0 = new TextBox();
                txtBoxCol0.Multiline = true;
                txtBoxCol0.Margin = new Padding(0);
                txtBoxCol0.Name = "txtBoxCol1" + i;
                txtBoxCol0.Width = 48;
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
                txtBoxCol.Width = 48;
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
                txtBoxCol.Width = 60;
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
                txtBoxCol.Width = 60;
                txtBoxCol.Height = 21;
                txtBoxCol.TextAlign = HorizontalAlignment.Center;
                txtBoxCol.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol, 3, i);
            }
            //根据数据类型决定已知点的个数并生成相应的单元格
            if (ApproximateDataType == DataType.Data.ConnectingTraverse)//闭附和导线
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol6f = new TextBox();
                txtBoxCol6f.Name = "txtBoxCol6f";
                txtBoxCol6f.Multiline = true;
                txtBoxCol6f.Margin = new Padding(0);
                txtBoxCol6f.Width = 60;
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
                txtBoxCol6l.Width = 60;
                txtBoxCol6l.Height = backCount * 44;
                txtBoxCol6l.TextAlign = HorizontalAlignment.Center;
                txtBoxCol6l.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol6l, 6, tableLayoutPanel.RowCount - 1 - backCount * 2);
                Control[] ctrlTxtBoxCol6l = tableLayoutPanel.Controls.Find(txtBoxCol6l.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol6l[0], backCount * 2);
                //平距
                for (int i = 2; i < tableLayoutPanel.RowCount - 1; i += 2)
                {
                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                    TextBox txtBoxCol7 = new TextBox();
                    txtBoxCol7.Multiline = true;
                    txtBoxCol7.Name = "txtBoxCol7" + i;
                    txtBoxCol7.Margin = new Padding(0);
                    txtBoxCol7.Width = 53;
                    txtBoxCol7.Height = 44;
                    txtBoxCol7.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol7.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol7, 7, i);
                    Control[] ctrlTxtBoxCol = tableLayoutPanel.Controls.Find(txtBoxCol7.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol[0], 2);
                }
                for (int i = 2; i < tableLayoutPanel.RowCount - 1; i += (tableLayoutPanel.RowCount - 1 - (2 * backCount) - 2))
                {
                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                    TextBox txtBoxCol9 = new TextBox();
                    txtBoxCol9.Multiline = true;
                    txtBoxCol9.Margin = new Padding(0);
                    txtBoxCol9.Name = "txtBoxCol9" + i;
                    txtBoxCol9.Width = 52;
                    txtBoxCol9.Height = backCount * 44;
                    txtBoxCol9.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol9.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol9, 9, i);
                    Control[] ctrlTxtBoxCol9 = tableLayoutPanel.Controls.Find(txtBoxCol9.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol9[0], backCount * 2);
                }
                for (int i = 2; i < tableLayoutPanel.RowCount - 1; i += (tableLayoutPanel.RowCount - 1 - (2 * backCount) - 2))
                {
                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                    TextBox txtBoxCol10 = new TextBox();
                    txtBoxCol10.Multiline = true;
                    txtBoxCol10.Margin = new Padding(0);
                    txtBoxCol10.Name = "txtBoxCol10" + i;
                    txtBoxCol10.Width = 52;
                    txtBoxCol10.Height = backCount * 44;
                    txtBoxCol10.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol10.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol10, 10, i);
                    Control[] ctrlTxtBoxCol10 = tableLayoutPanel.Controls.Find(txtBoxCol10.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol10[0], backCount * 2);
                }
            }
            else if (ApproximateDataType == DataType.Data.OpenTraverse)//支导线
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol6 = new TextBox();
                txtBoxCol6.Name = "txtBoxCol6";
                txtBoxCol6.Multiline = true;
                txtBoxCol6.Margin = new Padding(0);
                txtBoxCol6.Width = 60;
                txtBoxCol6.Height = 22 * backCount;
                txtBoxCol6.TextAlign = HorizontalAlignment.Center;
                txtBoxCol6.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol6, 6, 2);
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                Control[] ctrlTxtBoxCol6 = tableLayoutPanel.Controls.Find(txtBoxCol6.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol6[0], backCount);
                //平距
                for (int i = 2; i < tableLayoutPanel.RowCount - 1; i += 2)
                {
                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                    TextBox txtBoxCol7 = new TextBox();
                    txtBoxCol7.Multiline = true;
                    txtBoxCol7.Name = "txtBoxCol7" + i;
                    txtBoxCol7.Margin = new Padding(0);
                    txtBoxCol7.Width = 53;
                    txtBoxCol7.Height = 44;
                    txtBoxCol7.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol7.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol7, 7, i);
                    Control[] ctrlTxtBoxCol = tableLayoutPanel.Controls.Find(txtBoxCol7.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol[0], 2);
                }
                TextBox txtBoxCol9 = new TextBox();
                txtBoxCol9.Multiline = true;
                txtBoxCol9.Margin = new Padding(0);
                txtBoxCol9.Name = "txtBoxCol9" + 2;
                txtBoxCol9.Width = 52;
                txtBoxCol9.Height = backCount * 44;
                txtBoxCol9.TextAlign = HorizontalAlignment.Center;
                txtBoxCol9.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol9, 9, 2);
                Control[] ctrlTxtBoxCol9 = tableLayoutPanel.Controls.Find(txtBoxCol9.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol9[0], backCount * 2);
                TextBox txtBoxCol10 = new TextBox();
                txtBoxCol10.Multiline = true;
                txtBoxCol10.Margin = new Padding(0);
                txtBoxCol10.Name = "txtBoxCol10" + 2;
                txtBoxCol10.Width = 52;
                txtBoxCol10.Height = backCount * 44;
                txtBoxCol10.TextAlign = HorizontalAlignment.Center;
                txtBoxCol10.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol10, 10, 2);
                Control[] ctrlTxtBoxCol10 = tableLayoutPanel.Controls.Find(txtBoxCol10.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol10[0], backCount * 2);
            }

            //将控件添加到panel
            panel.Controls.Add(txtTitle);
            panel.Controls.Add(txtProjectName);
            panel.Controls.Add(txtInstrument);
            panel.Controls.Add(txtWeather);
            panel.Controls.Add(txtObserver);
            panel.Controls.Add(txtRecorder);
            panel.Controls.Add(txtDate);
            panel.Controls.Add(tableLayoutPanel);
            panel.Controls.Add(txtCalculate);
            panel.Controls.Add(txtAssessment);
            panel.Controls.Add(txtDataEnd);
            panel.Controls.Add(tableLayoutPanel);
            panel.Controls.Add(lbl);
        }
    }
}
