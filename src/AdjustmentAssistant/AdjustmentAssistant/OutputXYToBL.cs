using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AdjustmentAssistant
{
    class OutputXYToBL
    {
        internal static void GetData(ref TableLayoutPanel tableLayoutPanel, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5)
        {
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                TextBox txtB = new TextBox();
                txtB.Margin = new Padding(0);
                txtB.Text = col4[i - 2];
                txtB.BorderStyle = BorderStyle.None;
                txtB.TextAlign = HorizontalAlignment.Center;
                txtB.Width = 108;
                txtB.Height = 21;
                tableLayoutPanel.Controls.Add(txtB, 3, i);
            }
            for (int i = 2; i < tableLayoutPanel.RowCount; i++)
            {
                TextBox txtL = new TextBox();
                txtL.Margin = new Padding(0);
                txtL.Text = col5[i - 2];
                txtL.BorderStyle = BorderStyle.None;
                txtL.TextAlign = HorizontalAlignment.Center;
                txtL.Width = 108;
                txtL.Height = 21;
                tableLayoutPanel.Controls.Add(txtL, 4, i);
            }
        }
    }
}
