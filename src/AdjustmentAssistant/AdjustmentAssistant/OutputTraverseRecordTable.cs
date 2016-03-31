using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace AdjustmentAssistant
{
    public class OutputTraverseRecordTable
    {
        internal static void GetData(string filePath, int backCount, TableLayoutPanel tableLayoutPanel, ref double coordinateCloseError, ref double angleCloseError, ref double k, ref List<string> col0, ref List<string> col1, ref List<string> col2, ref List<string> col3, ref List<string> col4, ref List<string> col5, ref List<string> col6, ref List<string> col7, ref List<string> col8, ref List<string> col9, ref List<string> col10)
        {
            string[] strCol0;
            string[] strCol1;
            string[] strCol2;
            string[] strCol3;
            string[] strCol4;
            string[] strCol5;
            string[] strCol6;
            string[] strCol7;
            string[] strCol8;
            string[] strCol9;
            string[] strCol10;
            DataType.Data dataType;
            string[] split = { "," };
            using (StreamReader sr = new StreamReader(filePath, Encoding.Default))
            {
                string strDataType = sr.ReadLine().Trim();
                dataType = (DataType.Data)Enum.Parse(typeof(DataType.Data), strDataType.Split(split, StringSplitOptions.RemoveEmptyEntries)[0], false);
                coordinateCloseError = Convert.ToDouble(sr.ReadLine().Trim());
                angleCloseError = Convert.ToDouble(sr.ReadLine().Trim());
                strCol0 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
                strCol1 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
                strCol2 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
                strCol3 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
                strCol4 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
                strCol5 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
                strCol6 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
                strCol7 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
                strCol8 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
                strCol9 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
                strCol10 = sr.ReadLine().Trim().Split(split, StringSplitOptions.RemoveEmptyEntries);
            }
            GetList(strCol0, ref col0);
            GetList(strCol1, ref col1);
            GetList(strCol2, ref col2);
            GetList(strCol3, ref col3);
            GetList(strCol4, ref col4);
            GetList(strCol5, ref col5);
            GetList(strCol6, ref col6);
            GetList(strCol7, ref col7);
            GetList(strCol8, ref col8);
            GetList(strCol9, ref col9);
            GetList(strCol10, ref col10);
            for (int i = 0; i < (tableLayoutPanel.RowCount - 3) / 2; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol4 = new TextBox();
                txtBoxCol4.Name = "txtBoxCol4" + i;
                txtBoxCol4.Multiline = true;
                txtBoxCol4.Margin = new Padding(0);
                txtBoxCol4.Text = col4[i];
                txtBoxCol4.Width = 48;
                txtBoxCol4.Height = 44;
                txtBoxCol4.TextAlign = HorizontalAlignment.Center;
                txtBoxCol4.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol4, 4, i * 2 + 2);
                Control[] ctrlTxtBoxCol4 = tableLayoutPanel.Controls.Find(txtBoxCol4.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol4[0], 2);
            }
            for (int i = 0; i < (tableLayoutPanel.RowCount - 3) / backCount / 2; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol5 = new TextBox();
                txtBoxCol5.Name = "txtBoxCol5" + i;
                txtBoxCol5.Multiline = true;
                txtBoxCol5.Margin = new Padding(0);
                txtBoxCol5.Text = col5[i];
                txtBoxCol5.Width = 60;
                txtBoxCol5.Height = backCount * 44;
                txtBoxCol5.TextAlign = HorizontalAlignment.Center;
                txtBoxCol5.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol5, 5, i * 2 * backCount + 2);
                Control[] ctrlTxtBoxCol5 = tableLayoutPanel.Controls.Find(txtBoxCol5.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol5[0], backCount * 2);
            }
            //填充方位角
            if (dataType == DataType.Data.ConnectingTraverse)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol = new TextBox();
                txtBoxCol.Name = "txtBoxCol62";
                txtBoxCol.Multiline = true;
                txtBoxCol.Margin = new Padding(0);
                txtBoxCol.Text = col6[2];
                txtBoxCol.Width = 60;
                txtBoxCol.Height = backCount * 21;
                txtBoxCol.TextAlign = HorizontalAlignment.Center;
                txtBoxCol.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol, 6, backCount + 2);
                Control[] ctrlTxtBoxCol = tableLayoutPanel.Controls.Find(txtBoxCol.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol[0], backCount);
                for (int i = 3; i < (tableLayoutPanel.RowCount - 3) / backCount / 2 + 1; i++)
                {
                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                    TextBox txtBoxCol6 = new TextBox();
                    txtBoxCol6.Name = "txtBoxCol6" + i;
                    txtBoxCol6.Multiline = true;
                    txtBoxCol6.Margin = new Padding(0);
                    txtBoxCol6.Text = col6[i];
                    txtBoxCol6.Width = 60;
                    txtBoxCol6.Height = backCount * 44;
                    txtBoxCol6.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol6.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol6, 6, (i - 2) * 2 * backCount + 2);
                    Control[] ctrlTxtBoxCol6 = tableLayoutPanel.Controls.Find(txtBoxCol6.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol6[0], backCount * 2);
                }
            }
            else
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol = new TextBox();
                txtBoxCol.Name = "txtBoxCol61";
                txtBoxCol.Multiline = true;
                txtBoxCol.Margin = new Padding(0);
                txtBoxCol.Text = col6[1];
                txtBoxCol.Width = 60;
                txtBoxCol.Height = backCount * 21;
                txtBoxCol.TextAlign = HorizontalAlignment.Center;
                txtBoxCol.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol, 6, backCount + 2);
                Control[] ctrlTxtBoxCol = tableLayoutPanel.Controls.Find(txtBoxCol.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol[0], backCount);
                for (int i = 2; i < (tableLayoutPanel.RowCount - 3) / backCount / 2 + 1; i++)
                {
                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                    TextBox txtBoxCol6 = new TextBox();
                    txtBoxCol6.Name = "txtBoxCol6" + i;
                    txtBoxCol6.Multiline = true;
                    txtBoxCol6.Margin = new Padding(0);
                    txtBoxCol6.Text = col6[i];
                    txtBoxCol6.Width = 60;
                    txtBoxCol6.Height = backCount * 44;
                    txtBoxCol6.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol6.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol6, 6, (i - 1) * 2 * backCount + 2);
                    Control[] ctrlTxtBoxCol6 = tableLayoutPanel.Controls.Find(txtBoxCol6.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol6[0], backCount * 2);
                }
            }
            for (int i = 0; i < (tableLayoutPanel.RowCount - 3) / backCount / 2; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol8 = new TextBox();
                txtBoxCol8.Name = "txtBoxCol8" + i;
                txtBoxCol8.Multiline = true;
                txtBoxCol8.Margin = new Padding(0);
                txtBoxCol8.Text = col8[i];
                txtBoxCol8.Width = 53;
                txtBoxCol8.Height = backCount * 44;
                txtBoxCol8.TextAlign = HorizontalAlignment.Center;
                txtBoxCol8.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtBoxCol8, 8, i * 2 * backCount + 2);
                Control[] ctrlTxtBoxCol8 = tableLayoutPanel.Controls.Find(txtBoxCol8.Name, false);
                tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol8[0], backCount * 2);
            }
            //填充XY坐标值
            if (dataType == DataType.Data.ConnectingTraverse)
            {
                for (int i = backCount * 2 + 2; i < tableLayoutPanel.RowCount - backCount * 2 - 1; i += (backCount * 2))
                {
                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                    TextBox txtBoxCol9 = new TextBox();
                    txtBoxCol9.Name = "txtBoxCol9" + i;
                    txtBoxCol9.Multiline = true;
                    txtBoxCol9.Margin = new Padding(0);
                    txtBoxCol9.Text = col9[(i - 2) / backCount / 2 + 1].ToString();
                    txtBoxCol9.Width = 52;
                    txtBoxCol9.Height = backCount * 44;
                    txtBoxCol9.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol9.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol9, 9, i);
                    Control[] ctrlTxtBoxCol9 = tableLayoutPanel.Controls.Find(txtBoxCol9.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol9[0], backCount * 2);
                    TextBox txtBoxCol10 = new TextBox();
                    txtBoxCol10.Name = "txtBoxCol10" + i;
                    txtBoxCol10.Multiline = true;
                    txtBoxCol10.Margin = new Padding(0);
                    txtBoxCol10.Text = col10[(i - 2) / backCount / 2 + 1].ToString();
                    txtBoxCol10.Width = 52;
                    txtBoxCol10.Height = backCount * 44;
                    txtBoxCol10.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol10.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol10, 10, i);
                    Control[] ctrlTxtBoxCol10 = tableLayoutPanel.Controls.Find(txtBoxCol10.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol10[0], backCount * 2);
                }
                double d = 0;
                for (int i = 0; i < col8.Count; i++)
                {
                    d += double.Parse(col8[i]);
                }
                k = Math.Round(d * (1 / coordinateCloseError), 3);
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtRemark = new TextBox();
                txtRemark.Name = "txtRemark";
                txtRemark.Multiline = true;
                txtRemark.Margin = new Padding(0);
                txtRemark.Width = 550;
                txtRemark.Height = 21;
                txtRemark.Text = " 类型:闭附和线路 角度闭合差:" + angleCloseError + " 坐标增量闭合差:±" + coordinateCloseError + " K≈" + "1/" + k;
                txtRemark.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtRemark, 1, tableLayoutPanel.RowCount - 1);
                Control[] ctrlTxtRemark = tableLayoutPanel.Controls.Find(txtRemark.Name, false);
                tableLayoutPanel.SetColumnSpan(ctrlTxtRemark[0], 10);
            }
            else
            {
                for (int i = backCount * 2 + 2; i < tableLayoutPanel.RowCount - 1; i += (backCount * 2))
                {
                    tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                    TextBox txtBoxCol9 = new TextBox();
                    txtBoxCol9.Name = "txtBoxCol9" + i;
                    txtBoxCol9.Multiline = true;
                    txtBoxCol9.Margin = new Padding(0);
                    txtBoxCol9.Text = col9[(i - 2) / backCount / 2].ToString();
                    txtBoxCol9.Width = 52;
                    txtBoxCol9.Height = backCount * 44;
                    txtBoxCol9.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol9.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol9, 9, i);
                    Control[] ctrlTxtBoxCol9 = tableLayoutPanel.Controls.Find(txtBoxCol9.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol9[0], backCount * 2);
                    TextBox txtBoxCol10 = new TextBox();
                    txtBoxCol10.Name = "txtBoxCol10" + i;
                    txtBoxCol10.Multiline = true;
                    txtBoxCol10.Margin = new Padding(0);
                    txtBoxCol10.Text = col10[(i - 2) / backCount / 2].ToString();
                    txtBoxCol10.Width = 52;
                    txtBoxCol10.Height = backCount * 44;
                    txtBoxCol10.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol10.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol10, 10, i);
                    Control[] ctrlTxtBoxCol10 = tableLayoutPanel.Controls.Find(txtBoxCol10.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol10[0], backCount * 2);
                }
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtRemark = new TextBox();
                txtRemark.Name = "txtRemark";
                txtRemark.Multiline = true;
                txtRemark.Margin = new Padding(0);
                txtRemark.Width = 555;
                txtRemark.Height = 21;
                txtRemark.Text = " 类型:支导线路线";
                txtRemark.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtRemark, 1, tableLayoutPanel.RowCount - 1);
                Control[] ctrlTxtRemark = tableLayoutPanel.Controls.Find(txtRemark.Name, false);
                tableLayoutPanel.SetColumnSpan(ctrlTxtRemark[0], 10);
            }
        }

        private static void GetList(string[] strCol, ref List<string> col)
        {
            foreach (string strItem in strCol)
            {
                col.Add(strItem);
            }
        }
    }
}
