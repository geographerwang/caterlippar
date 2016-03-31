using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Adjustment;
using System.Windows.Forms;

namespace AdjustmentAssistant
{
    class OutputParameterAdjustment
    {
        internal static void GetData(string filePath, TableLayoutPanel tableLayoutPanel, ref double unitError, List<string> col0, List<string> col1, List<string> col2, List<string> col4, List<string> col5, ref List<string> col6, ref List<string> col7, ref List<string> col8, ref List<string> col9, ref List<string> col10, ref List<string> col11, ref List<string> col12, ref List<string> col13, ref List<string> col14, ref List<string> col15, ref List<string> col16, ref List<string> col17, ref List<string> col18)
        {
            using (StreamReader sr = new StreamReader(filePath))
            {
                string[] strSplit = { " " };
                string[] strArrCol6 = new string[2];
                string[] strArrCol7 = new string[2];
                string[] strArrCol8 = new string[2];
                string[] strArrCol15 = new string[2];
                string[] strArrCol16 = new string[2];
                unitError = Convert.ToDouble(sr.ReadLine().Trim());//平差结果的第一行是单位权中误差，暂时不需要
                //这部分读取的是坐标平差值及其精度
                for (int i = 0; i < 2; i++)
                {
                    string strAdjCoord = sr.ReadLine();
                    string[] strArrAdjCoord = strAdjCoord.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                    col14.Add(strArrAdjCoord[2]);
                    col15.Add(strArrAdjCoord[3]);
                    col6.Add(strArrAdjCoord[4]);
                    col7.Add(strArrAdjCoord[5]);
                    col8.Add(strArrAdjCoord[6]);
                }
                for (int i = 0; i < 2; i++)
                {
                    string strAdjCoord = sr.ReadLine();
                    string[] strArrAdjCoord = strAdjCoord.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                    strArrCol15[i] = strArrAdjCoord[2];
                    strArrCol16[i] = strArrAdjCoord[3];
                    strArrCol6[i] = strArrAdjCoord[4];
                    strArrCol7[i] = strArrAdjCoord[5];
                    strArrCol8[i] = strArrAdjCoord[6];
                }
                for (int i = 4; i < col0.Count; i++)
                {
                    string strAdjCoord = sr.ReadLine();
                    string[] strArrAdjCoord = strAdjCoord.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                    col14.Add(strArrAdjCoord[2]);
                    col15.Add(strArrAdjCoord[3]);
                    col6.Add(strArrAdjCoord[4]);
                    col7.Add(strArrAdjCoord[5]);
                    col8.Add(strArrAdjCoord[6]);
                }
                for (int i = 0; i < 2; i++)
                {
                    col14.Add(strArrCol15[i]);
                    col15.Add(strArrCol16[i]);
                    col6.Add(strArrCol6[i]);
                    col7.Add(strArrCol7[i]);
                    col8.Add(strArrCol8[i]);
                }
                for (int i = 0; i < col0.Count - 2; i++)
                {
                    string strFront = sr.ReadLine();
                    string[] strArrFront = strFront.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                    string strAfter = sr.ReadLine();
                    string[] strArrAfter = strAfter.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                    double douAngle;
                    double douFront = Convert.ToDouble(strArrFront[4]);
                    double douAfter = Convert.ToDouble(strArrAfter[3]);
                    if (i == 0)
                    {
                        if (douFront > Math.PI)
                        {
                            col16.Add(ConvertAngle.RealDegreeToString((douFront - Math.PI) / Math.PI * 180));
                        }
                        else
                        {
                            col16.Add(ConvertAngle.RealDegreeToString((douFront + Math.PI) / Math.PI * 180));
                        }
                        col17.Add(Math.Round(Convert.ToDouble(strArrFront[6]), 4).ToString());
                        col9.Add("0");
                        col10.Add("0");
                    }
                    col16.Add(ConvertAngle.RealDegreeToString(douAfter / Math.PI * 180));
                    col17.Add(Math.Round(Convert.ToDouble(strArrAfter[5]), 4).ToString());
                    if (douAfter > douFront)
                    {
                        douAngle = douAfter - douFront;
                    }
                    else
                    {
                        douAngle = douAfter - douFront + 2 * Math.PI;
                    }
                    col18.Add(ConvertAngle.RealDegreeToString(douAngle / Math.PI * 180));
                    col9.Add(Math.Round(Convert.ToDouble(strArrAfter[6]), 4).ToString());
                    col10.Add(Math.Round(Convert.ToDouble(strArrAfter[4]), 3).ToString());
                }
                for (int i = 0; i < col0.Count - 2; i++)
                {
                    sr.ReadLine();
                }
                for (int i = 0; i < col0.Count - 4; i++)
                {
                    string strEllipses = sr.ReadLine();
                    string[] strArrEllipses = strEllipses.Split(strSplit, StringSplitOptions.RemoveEmptyEntries);
                    col11.Add(strArrEllipses[1]);
                    col12.Add(strArrEllipses[2]);
                    col13.Add(ConvertAngle.DegreeToString(Convert.ToDouble(strArrEllipses[3]) / Math.PI * 180));
                }
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
        }
    }
}
