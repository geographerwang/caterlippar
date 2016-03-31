using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AdjustmentAssistant
{
    class OutputXYToXY
    {
        internal static void GetData(ref TableLayoutPanel tableLayoutPanel, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                TextBox txtX = new TextBox();
                txtX.Margin = new Padding(0);
                txtX.Text = col4[i - 2];
                txtX.BorderStyle = BorderStyle.None;
                txtX.TextAlign = HorizontalAlignment.Center;
                txtX.Width = 98;
                txtX.Height = 21;
                tableLayoutPanel.Controls.Add(txtX, 3, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                TextBox txtY = new TextBox();
                txtY.Margin = new Padding(0);
                txtY.Text = col5[i - 2];
                txtY.BorderStyle = BorderStyle.None;
                txtY.TextAlign = HorizontalAlignment.Center;
                txtY.Width = 98;
                txtY.Height = 21;
                tableLayoutPanel.Controls.Add(txtY, 4, i);
            }
        }
    }
}
