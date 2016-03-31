using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AdjustmentAssistant
{
    class InputBLToXY
    {
        internal static void GetData(List<string> col0, List<string> col1, List<string> col2, ref TableLayoutPanel tableLayoutPanel)
        {
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol = new TextBox();
                txtBoxCol.Margin = new Padding(0);
                txtBoxCol.Width = 78;
                txtBoxCol.TextAlign = HorizontalAlignment.Center;
                txtBoxCol.Text = col0[i - 2];
                txtBoxCol.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol, 0, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol = new TextBox();
                txtBoxCol.Margin = new Padding(0);
                txtBoxCol.Width = 108;
                txtBoxCol.TextAlign = HorizontalAlignment.Center;
                txtBoxCol.Text = col1[i - 2];
                txtBoxCol.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol, 1, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol = new TextBox();
                txtBoxCol.Margin = new Padding(0);
                txtBoxCol.Width = 108;
                txtBoxCol.TextAlign = HorizontalAlignment.Center;
                txtBoxCol.Text = col2[i - 2];
                txtBoxCol.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol, 2, i);
            }
        }
    }
}
