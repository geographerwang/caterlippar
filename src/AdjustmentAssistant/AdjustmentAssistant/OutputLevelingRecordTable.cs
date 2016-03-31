using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace AdjustmentAssistant
{
    public class OutputLevelingRecordTable
    {
        internal static void GetData(string filePath, int backCount, System.Windows.Forms.TableLayoutPanel tableLayoutPanel, ref List<string> col0, ref List<string> col1, ref List<string> col2, ref List<string> col3, ref List<string> col4, ref List<string> col5, ref List<string> col6)
        {
            string[] strCol0;
            string[] strCol1;
            string[] strCol2;
            string[] strCol3;
            string[] strCol4;
            string[] strCol5;
            string[] strCol6;
            DataType.Data dataType;
            double coordinateCloseError;
            double angleCloseError;
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
            }
            GetList(strCol0, ref col0);
            GetList(strCol1, ref col1);
            GetList(strCol2, ref col2);
            GetList(strCol3, ref col3);
            GetList(strCol4, ref col4);
            GetList(strCol5, ref col5);
            GetList(strCol6, ref col6);
            for (int i = 0; i < (tableLayoutPanel.RowCount - 3) / 2; i++)
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol4 = new TextBox();
                txtBoxCol4.Name = "txtBoxCol4" + i;
                txtBoxCol4.Multiline = true;
                txtBoxCol4.Margin = new Padding(0);
                txtBoxCol4.Text = col4[i];
                txtBoxCol4.Width = 56;
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
                txtBoxCol5.Width = 108;
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
                txtBoxCol.Width = 108;
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
                    txtBoxCol6.Width = 108;
                    txtBoxCol6.Height = backCount * 44;
                    txtBoxCol6.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol6.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol6, 6, (i - 2) * 2 * backCount + 2);
                    Control[] ctrlTxtBoxCol6 = tableLayoutPanel.Controls.Find(txtBoxCol6.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol6[0], backCount * 2);
                }
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtRemark = new TextBox();
                txtRemark.Name = "txtRemark";
                txtRemark.Multiline = true;
                txtRemark.Margin = new Padding(0);
                txtRemark.Width = 550;
                txtRemark.Height = 21;
                txtRemark.BorderStyle = BorderStyle.None;
                txtRemark.Text = " 类型:闭附合路线   角度闭合差:" + angleCloseError;
                tableLayoutPanel.Controls.Add(txtRemark, 1, tableLayoutPanel.RowCount - 1);
                Control[] ctrlTxtRemark = tableLayoutPanel.Controls.Find(txtRemark.Name, false);
                tableLayoutPanel.SetColumnSpan(ctrlTxtRemark[0], 6);
            }
            else
            {
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtBoxCol = new TextBox();
                txtBoxCol.Name = "txtBoxCol61";
                txtBoxCol.Multiline = true;
                txtBoxCol.Margin = new Padding(0);
                txtBoxCol.Text = col6[1];
                txtBoxCol.Width = 108;
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
                    txtBoxCol6.Width = 108;
                    txtBoxCol6.Height = backCount * 44;
                    txtBoxCol6.TextAlign = HorizontalAlignment.Center;
                    txtBoxCol6.BorderStyle = BorderStyle.None;
                    tableLayoutPanel.Controls.Add(txtBoxCol6, 6, (i - 1) * 2 * backCount + 2);
                    Control[] ctrlTxtBoxCol6 = tableLayoutPanel.Controls.Find(txtBoxCol6.Name, false);
                    tableLayoutPanel.SetRowSpan(ctrlTxtBoxCol6[0], backCount * 2);
                }
                tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 21f));
                TextBox txtRemark = new TextBox();
                txtRemark.Name = "txtRemark";
                txtRemark.Multiline = true;
                txtRemark.Margin = new Padding(0);
                txtRemark.Width = 550;
                txtRemark.Height = 21;
                txtRemark.Text = " 类型:支导线路线";
                txtRemark.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtRemark, 1, tableLayoutPanel.RowCount - 1);
                Control[] ctrlTxtRemark = tableLayoutPanel.Controls.Find(txtRemark.Name, false);
                tableLayoutPanel.SetColumnSpan(ctrlTxtRemark[0], 6);
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
