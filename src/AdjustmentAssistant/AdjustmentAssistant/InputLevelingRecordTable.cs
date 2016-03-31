using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace AdjustmentAssistant
{
    public class InputLevelingRecordTable
    {
        internal static void GetData(System.Windows.Forms.Panel pnlResult, int dataCount, int backCount, DataType.Data ApproximateDataType, ref System.Windows.Forms.TableLayoutPanel tableLayoutPanel)
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

            //测站点
            for (int i = 2; i < tableLayoutPanel.RowCount - 1; i += (2 * backCount))
            {
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
                TextBox txtBoxCol = new TextBox();
                txtBoxCol.Margin = new Padding(0);
                txtBoxCol.Width = 108;
                txtBoxCol.Height = 21;
                txtBoxCol.TextAlign = HorizontalAlignment.Center;
                txtBoxCol.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol, 3, i);
            }
            //根据数据类型决定已知点的个数并生成相应的单元格
            if (ApproximateDataType == DataType.Data.ConnectingTraverse)//闭附和导线
            {
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
            else if (ApproximateDataType == DataType.Data.OpenTraverse)//支导线
            {
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
            //表末尾备注
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
        }
    }
}
