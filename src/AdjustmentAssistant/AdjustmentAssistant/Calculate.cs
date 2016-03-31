using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Adjustment;

namespace AdjustmentAssistant
{
    public static class Calculate
    {
        public static string TraverseRecord(int dataCount, int backCount, DataType.Data approximateDataType, DataType.LeftOrRight LorR, ref TableLayoutPanel tableLayoutPanel)
        {
            List<string> list = new List<string>();
            List<string> col0 = new List<string>();
            List<string> col1 = new List<string>();
            List<string> col2 = new List<string>();
            List<string> col3 = new List<string>();
            List<string> col6 = new List<string>();
            List<string> col7 = new List<string>();
            List<string> col9 = new List<string>();
            List<string> col10 = new List<string>();
            if (tableLayoutPanel == null)
            {
                return null;
            }
            foreach (Control item in tableLayoutPanel.Controls)
            {
                TextBox txtBox = item as TextBox;
                if (txtBox != null)
                {
                    if (txtBox.Text != "")
                    {
                        list.Add(txtBox.Text);
                    }
                    else
                    {
                        MessageBox.Show("请将数据填写完整！");
                        return null;
                    }
                }
            }
            for (int i = 0; i < (tableLayoutPanel.RowCount - 3) / (backCount * 2); i++)
            {
                col0.Add(list[i]);
            }
            for (int i = col0.Count; i < col0.Count + tableLayoutPanel.RowCount - 3; i++)
            {
                col1.Add(list[i]);
            }
            for (int i = col0.Count + col1.Count; i < col0.Count + col1.Count + tableLayoutPanel.RowCount - 3; i++)
            {
                col2.Add(list[i]);
            }
            for (int i = col0.Count + col1.Count + col2.Count; i < col0.Count + col1.Count + col2.Count + tableLayoutPanel.RowCount - 3; i++)
            {
                col3.Add(list[i]);
            }
            if (approximateDataType == DataType.Data.ConnectingTraverse)
            {
                for (int i = col0.Count + col1.Count + col2.Count + col3.Count + 2; i < col0.Count + col1.Count + col2.Count + col3.Count + 2 + (tableLayoutPanel.RowCount - 3) / 2; i++)
                {
                    col7.Add(list[i]);
                }
                int j = col0.Count + col1.Count + col2.Count + col3.Count;
                col6.Add(list[j]);
                col6.Add(list[j + 1]);//col6第二个数存放的是最后一个数据
                col9.Add(list[j + col6.Count + col7.Count]);
                col9.Add(list[j + col6.Count + col7.Count + 1]);//col9第二个数据存放的是最后一个数据
                col10.Add(list[j + col6.Count + col7.Count + col9.Count]);
                col10.Add(list[j + col6.Count + col7.Count + col9.Count + 1]);//col10第二个数据存放的是最后一个数据
            }
            else if (approximateDataType == DataType.Data.OpenTraverse)
            {
                for (int i = col0.Count + col1.Count + col2.Count + col3.Count + 1; i < col0.Count + col1.Count + col2.Count + col3.Count + 1 + (tableLayoutPanel.RowCount - 3) / 2; i++)
                {
                    col7.Add(list[i]);
                }
                int j = col0.Count + col1.Count + col2.Count + col3.Count;
                col6.Add(list[j]);
                col9.Add(list[j + col6.Count + col7.Count]);
                col10.Add(list[j + col6.Count + col7.Count + col9.Count]);
            }
            return ApproximateAdjustment.GetTraverse(col0, col1, col2, col3, col6, col7, col9, col10, dataCount, backCount, approximateDataType, LorR);
        }

        public static string LevelingRecord(int dataCount, int backCount, DataType.Data approximateDataType, DataType.LeftOrRight LorR, ref TableLayoutPanel tableLayoutPanel)
        {
            List<string> list = new List<string>();
            List<string> col0 = new List<string>();
            List<string> col1 = new List<string>();
            List<string> col2 = new List<string>();
            List<string> col3 = new List<string>();
            List<string> col6 = new List<string>();
            if (tableLayoutPanel == null)
            {
                return null;
            }
            foreach (Control item in tableLayoutPanel.Controls)
            {
                TextBox txtBox = item as TextBox;
                if (txtBox != null)
                {
                    if (txtBox.Text != "")
                    {
                        list.Add(txtBox.Text);
                    }
                    else
                    {
                        MessageBox.Show("请将数据填写完整！");
                        return null;
                    }
                }
            }
            for (int i = 0; i < (tableLayoutPanel.RowCount - 3) / (backCount * 2); i++)
            {
                col0.Add(list[i]);
            }
            for (int i = col0.Count; i < col0.Count + tableLayoutPanel.RowCount - 3; i++)
            {
                col1.Add(list[i]);
            }
            for (int i = col0.Count + col1.Count; i < col0.Count + col1.Count + tableLayoutPanel.RowCount - 3; i++)
            {
                col2.Add(list[i]);
            }
            for (int i = col0.Count + col1.Count + col2.Count; i < col0.Count + col1.Count + col2.Count + tableLayoutPanel.RowCount - 3; i++)
            {
                col3.Add(list[i]);
            }
            if (approximateDataType == DataType.Data.ConnectingTraverse)
            {
                int j = col0.Count + col1.Count + col2.Count + col3.Count;
                col6.Add(list[j]);
                col6.Add(list[j + 1]);//col6第二个数存放的是最后一个数据
            }
            else
            {
                int j = col0.Count + col1.Count + col2.Count + col3.Count;
                col6.Add(list[j]);
            }
            return ApproximateAdjustment.GetLeveling(col0, col1, col2, col3, col6, dataCount, backCount, approximateDataType, LorR);
        }

        internal static string ParameterAdjustment(List<string> col0, List<string> col1, List<string> col2, ref List<string> col4, ref List<string> col5, double[] accuracy, DataType.Data approximateDataType, DataType.LeftOrRight lorR, ref TableLayoutPanel tableLayoutPanel)
        {
            List<double> azimuth = new List<double>();
            for (int i = 0; i < col0.Count - 1; i++)
            {
                double deltaX = Convert.ToDouble(col1[i + 1]) - Convert.ToDouble(col1[i]);
                double deltaY = Convert.ToDouble(col2[i + 1]) - Convert.ToDouble(col2[i]);
                col4.Add(Math.Sqrt(Math.Abs(deltaX) * Math.Abs(deltaX) + Math.Abs(deltaY) * Math.Abs(deltaY)).ToString("#.000000"));
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
                double xf = Convert.ToDouble(col1[i]);
                double xa = Convert.ToDouble(col1[i + 1]);
                double yf = Convert.ToDouble(col2[i]);
                double ya = Convert.ToDouble(col2[i + 1]);
                double deltaX = Math.Round(xa - xf, 3);
                double deltaY = Math.Round(ya - yf, 3);
                double rian = Math.Atan2(deltaY, deltaX);
                if (rian >= 0)
                {
                    azimuth.Add(rian);
                    col5.Add(ConvertAngle.DegreeToString(rian * 180 / Math.PI));
                }
                else
                {
                    azimuth.Add(2 * Math.PI + rian);
                    col5.Add(ConvertAngle.DegreeToString(360 + rian * 180 / Math.PI));
                }
                TextBox txtAngle = new TextBox();
                txtAngle.Text = col5[i];
                txtAngle.Margin = new Padding(0);
                txtAngle.Width = 60;
                txtAngle.Height = 21;
                txtAngle.TextAlign = HorizontalAlignment.Center;
                txtAngle.BorderStyle = BorderStyle.None;
                tableLayoutPanel.Controls.Add(txtAngle, 4, i + 2);
            }
            string tempFile = Path.GetTempFileName();
            using (StreamWriter sw = new StreamWriter(tempFile))
            {
                sw.WriteLine("{0} {1} {2} {3} {4} {5}", col0.Count, 4, col0.Count - 2, (col0.Count - 2) * 2, col0.Count - 3, 1);
                sw.WriteLine("{0} {1} {2} {3}", accuracy[0], accuracy[1], accuracy[2], accuracy[3]);
                sw.WriteLine("{0} {1} {2}", col0[0], col1[0], col2[0]);
                sw.WriteLine("{0} {1} {2}", col0[1], col1[1], col2[1]);
                sw.WriteLine("{0} {1} {2}", col0[col0.Count - 2], accuracy[4], accuracy[5]);
                sw.WriteLine("{0} {1} {2}", col0[col0.Count - 1], col1[col1.Count - 1], col2[col2.Count - 1]);
                for (int i = 1; i < col0.Count - 1; i++)
                {
                    sw.WriteLine("{0} {1}", col0[i], 2);
                    sw.WriteLine("{0} {1}", col0[i - 1], 0);
                    sw.WriteLine("{0} {1}", col0[i + 1], Math.PI - azimuth[i - 1] + azimuth[i]);
                }
                for (int i = 1; i < col0.Count - 2; i++)
                {
                    sw.WriteLine("{0} {1} {2}", col0[i], col0[i + 1], col4[i]);
                }
                sw.WriteLine("{0} {1} {2}", col0[2], col0[3], azimuth[2]);
            }
            return Adjustment.ParameterAdjustment.Calculate(tempFile);
        }
    }
}
