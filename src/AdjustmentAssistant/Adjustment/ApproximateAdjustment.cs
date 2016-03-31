using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Adjustment
{
    public class ApproximateAdjustment
    {
        public static string GetTraverse(List<string> col0, List<string> col1, List<string> col2, List<string> col3, List<string> col6, List<string> col7, List<string> col9, List<string> col10, int dataCount, int backCount, DataType.Data approximateDataType, DataType.LeftOrRight LorR)
        {
            List<string> col4 = new List<string>();
            List<string> col5 = new List<string>();
            List<string> col8 = new List<string>();
            double coordinateCloseError = 0;
            double angleCloseError = 0;
            for (int i = 0; i < dataCount; i += 2)
            {
                double left;
                double right;
                if (ConvertAngle.SecondFromString(col2[i + 1]) >= ConvertAngle.SecondFromString(col2[i]) && ConvertAngle.SecondFromString(col3[i + 1]) >= ConvertAngle.SecondFromString(col3[i]))
                {
                    left = ConvertAngle.SecondFromString(col2[i + 1]) - ConvertAngle.SecondFromString(col2[i]);
                    right = ConvertAngle.SecondFromString(col3[i + 1]) - ConvertAngle.SecondFromString(col3[i]);
                    col4.Add(Math.Abs(left - right).ToString());
                }
                else if (ConvertAngle.SecondFromString(col2[i + 1]) >= ConvertAngle.SecondFromString(col2[i]) && ConvertAngle.SecondFromString(col3[i + 1]) < ConvertAngle.SecondFromString(col3[i]))
                {
                    left = ConvertAngle.SecondFromString(col2[i + 1]) - ConvertAngle.SecondFromString(col2[i]);
                    right = ConvertAngle.SecondFromString(col3[i + 1]) - ConvertAngle.SecondFromString(col3[i]) + 360 * 3600;
                    col4.Add(Math.Abs(left - right).ToString());
                }
                else if (ConvertAngle.SecondFromString(col2[i + 1]) < ConvertAngle.SecondFromString(col2[i]) && ConvertAngle.SecondFromString(col3[i + 1]) >= ConvertAngle.SecondFromString(col3[i]))
                {
                    left = ConvertAngle.SecondFromString(col2[i + 1]) - ConvertAngle.SecondFromString(col2[i]) + 360 * 3600;
                    right = ConvertAngle.SecondFromString(col3[i + 1]) - ConvertAngle.SecondFromString(col3[i]);
                    col4.Add(Math.Abs(left - right).ToString());
                }
                else
                {
                    left = ConvertAngle.SecondFromString(col2[i + 1]) - ConvertAngle.SecondFromString(col2[i]) + 360 * 3600;
                    right = ConvertAngle.SecondFromString(col3[i + 1]) - ConvertAngle.SecondFromString(col3[i]) + 360 * 3600;
                    col4.Add(Math.Abs(left - right).ToString());
                }
            }
            for (int i = 0; i < dataCount; i += (backCount * 2))
            {
                double allSecond = 0;
                for (int j = 0; j < backCount * 2; j += 2)
                {
                    double left;
                    double right;
                    if (ConvertAngle.SecondFromString(col2[i + j + 1]) >= ConvertAngle.SecondFromString(col2[i + j]) && ConvertAngle.SecondFromString(col3[i + j + 1]) >= ConvertAngle.SecondFromString(col3[i + j]))
                    {
                        left = ConvertAngle.SecondFromString(col2[i + j + 1]) - ConvertAngle.SecondFromString(col2[i + j]);
                        right = ConvertAngle.SecondFromString(col3[i + j + 1]) - ConvertAngle.SecondFromString(col3[i + j]);
                        allSecond += (left + right);
                    }
                    else if (ConvertAngle.SecondFromString(col2[i + j + 1]) >= ConvertAngle.SecondFromString(col2[i + j]) && ConvertAngle.SecondFromString(col3[i + j + 1]) < ConvertAngle.SecondFromString(col3[i + j]))
                    {
                        left = ConvertAngle.SecondFromString(col2[i + j + 1]) - ConvertAngle.SecondFromString(col2[i + j]);
                        right = ConvertAngle.SecondFromString(col3[i + j + 1]) - ConvertAngle.SecondFromString(col3[i + j]) + 360 * 3600;
                        allSecond += (left + right);
                    }
                    else if (ConvertAngle.SecondFromString(col2[i + j + 1]) < ConvertAngle.SecondFromString(col2[i + j]) && ConvertAngle.SecondFromString(col3[i + j + 1]) >= ConvertAngle.SecondFromString(col3[i + j]))
                    {
                        left = ConvertAngle.SecondFromString(col2[i + j + 1]) - ConvertAngle.SecondFromString(col2[i + j]) + 360 * 3600;
                        right = ConvertAngle.SecondFromString(col3[i + j + 1]) - ConvertAngle.SecondFromString(col3[i + j]);
                        allSecond += (left + right);
                    }
                    else
                    {
                        left = ConvertAngle.SecondFromString(col2[i + j + 1]) - ConvertAngle.SecondFromString(col2[i + j]) + 360 * 3600;
                        right = ConvertAngle.SecondFromString(col3[i + j + 1]) - ConvertAngle.SecondFromString(col3[i + j]) + 360 * 3600;
                        allSecond += (left + right);
                    }
                }
                col5.Add(ConvertAngle.SecondToString(allSecond / (backCount * 2)));
            }
            if (LorR == DataType.LeftOrRight.Left)//1代表左角
            {
                if ((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) + 180 * 3600 <= 360 * 3600)
                {
                    col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) + 180 * 3600));
                }
                else
                {
                    col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) - 180 * 3600));
                }
                if (approximateDataType == DataType.Data.ConnectingTraverse)//闭附和导线
                {
                    for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                    {
                        if ((ConvertAngle.SecondFromString(col6[i + 2]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600 <= 360 * 3600)
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600));
                        }
                        else
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) + ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600));
                        }
                    }
                    //还要进行平差，左角平差，将闭合差相加
                    angleCloseError = ConvertAngle.SecondFromString(col6[1]) - ConvertAngle.SecondFromString(col6[col6.Count - 1]);
                    for (int i = 0; i < Math.Abs(angleCloseError); i++)
                    {
                        if (angleCloseError < 0)
                        {
                            if (ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1 < 360 * 3600)
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1);
                            }
                            else
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1 - 360 * 3600);
                            }
                        }
                        if (angleCloseError > 0)
                        {
                            if (ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1 > 0)
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1);
                            }
                            else
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1 + 360 * 3600);
                            }
                        }
                    }
                    if ((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) + 180 * 3600 <= 360 * 3600)
                    {
                        col6[2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) + 180 * 3600);
                    }
                    else
                    {
                        col6[2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) - 180 * 3600);
                    }
                    for (int i = 1; i < col6.Count - 3; i++)
                    {
                        if ((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i])) + 180 * 3600 <= 360 * 3600)
                        {
                            col6[i + 2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i])) + 180 * 3600);
                        }
                        else
                        {
                            col6[i + 2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i])) - 180 * 3600);
                        }
                    }
                }
                else if (approximateDataType == DataType.Data.OpenTraverse)//支导线
                {
                    for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                    {
                        if ((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600 <= 360 * 3600)
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600));
                        }
                        else
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600));
                        }
                    }
                }
            }
            else if (LorR == DataType.LeftOrRight.Right)//2代表右角
            {
                if ((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) - 180 * 3600 > 0)
                {
                    col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) - 180 * 3600));
                }
                else if ((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) + 180 * 3600 > 0)
                {
                    col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) + 180 * 3600));
                }
                else
                {
                    col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) + 540 * 3600));
                }
                if (approximateDataType == DataType.Data.ConnectingTraverse)//闭附和导线
                {
                    for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                    {
                        if ((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600 >= 0)
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600));
                        }
                        else if ((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) > -180 * 3600)
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600));
                        }
                        else
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) + 540 * 3600));
                        }
                    }
                    //右角平差将闭合差相减
                    angleCloseError = ConvertAngle.SecondFromString(col6[1]) - ConvertAngle.SecondFromString(col6[col6.Count - 1]);
                    for (int i = 0; i < Math.Abs(angleCloseError); i++)
                    {
                        if (angleCloseError > 0)
                        {
                            if (ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1 > 0)
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1);
                            }
                            else
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1 + 360 * 3600);
                            }
                        }
                        if (angleCloseError < 0)
                        {
                            if (ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1 < 360 * 3600)
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1);
                            }
                            else
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1 - 360 * 3600);
                            }
                        }
                    }
                    if ((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) - 180 * 3600 > 0)
                    {
                        col6[2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) - 180 * 3600);
                    }
                    else if ((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) + 180 * 3600 > 0)
                    {
                        col6[2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) + 180 * 3600);
                    }
                    else
                    {
                        col6[2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) + 540 * 3600);
                    }
                    for (int i = 1; i < col6.Count - 3; i++)
                    {
                        if ((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600 >= 0)
                        {
                            col6[i + 2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 1]) - ConvertAngle.SecondFromString(col5[i])) - 180 * 3600);
                        }
                        else if ((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) > -180 * 3600)
                        {
                            col6[i + 2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 1]) - ConvertAngle.SecondFromString(col5[i])) + 180 * 3600);
                        }
                        else
                        {
                            col6[i + 2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) + 540 * 3600);
                        }
                    }
                }
                else//支导线
                {
                    for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                    {
                        if ((ConvertAngle.SecondFromString(col6[i]) - ConvertAngle.SecondFromString(col5[i])) - 180 * 3600 >= 0)
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i]) - ConvertAngle.SecondFromString(col5[i])) - 180 * 3600));
                        }
                        else
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i]) - ConvertAngle.SecondFromString(col5[i])) + 180 * 3600));
                        }
                    }
                }
            }
            //计算平均距离
            for (int i = 0; i < dataCount / 2; i += backCount)
            {
                double hDistance = 0;
                for (int j = 0; j < backCount; j++)
                {
                    hDistance += Convert.ToDouble(col7[i + j]);
                }
                col8.Add((hDistance / backCount).ToString());
            }
            //计算坐标
            if (approximateDataType == DataType.Data.ConnectingTraverse)
            {
                for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                {
                    //平差后，最后一个值用于计算闭合差
                    col9.Add(Math.Round((Convert.ToDouble(col8[i]) * Math.Cos(ConvertAngle.DegreeToRadian(col6[i + 2])) + double.Parse(col9[i])), 3).ToString());
                    col10.Add(Math.Round((Convert.ToDouble(col8[i]) * Math.Sin(ConvertAngle.DegreeToRadian(col6[i + 2])) + double.Parse(col10[i])), 3).ToString());
                }
                double x = Math.Round(double.Parse(col9[1]) - double.Parse(col9[col9.Count - 1]), 3);
                double y = Math.Round(double.Parse(col10[1]) - double.Parse(col10[col10.Count - 1]), 3);
                coordinateCloseError = Math.Round(Math.Sqrt(x * x + y * y), 3);
                for (int i = 0; i < (int)Math.Abs(x * 1000); i++)
                {
                    if (x > 0)
                    {
                        col9[(i % (dataCount / backCount / 2 - 2)) + 2] = (Convert.ToDouble(col9[(i % (dataCount / backCount / 2 - 2)) + 2]) - 0.001).ToString();
                    }
                    else if (x < 0)
                    {
                        col9[(i % (dataCount / backCount / 2 - 2)) + 2] = (Convert.ToDouble(col9[(i % (dataCount / backCount / 2 - 2)) + 2]) + 0.001).ToString();
                    }
                }
                for (int i = 0; i < (int)Math.Abs(y * 1000); i++)
                {
                    if (y > 0)
                    {
                        col10[(i % (dataCount / backCount / 2 - 2)) + 2] = (Convert.ToDouble(col10[(i % (dataCount / backCount / 2 - 2)) + 2]) - 0.001).ToString();
                    }
                    else if (y < 0)
                    {
                        col10[(i % (dataCount / backCount / 2 - 2)) + 2] = (Convert.ToDouble(col10[(i % (dataCount / backCount / 2 - 2)) + 2]) + 0.001).ToString();
                    }
                }
            }
            else if (approximateDataType == DataType.Data.OpenTraverse)
            {
                for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                {
                    //最后一个值必须保留
                    col9.Add(Math.Round((Convert.ToDouble(col8[i]) * Math.Cos(ConvertAngle.DegreeToRadian(col6[i + 1])) + double.Parse(col9[i])), 3).ToString());
                    col10.Add(Math.Round((Convert.ToDouble(col8[i]) * Math.Sin(ConvertAngle.DegreeToRadian(col6[i + 1])) + double.Parse(col10[i])), 3).ToString());
                }
            }
            return PrintResult(col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, approximateDataType, coordinateCloseError, angleCloseError);
        }

        public static string PrintResult(List<string> col0, List<string> col1, List<string> col2, List<string> col3, List<string> col4, List<string> col5, List<string> col6, List<string> col7, List<string> col8, List<string> col9, List<string> col10, DataType.Data approximateDataType, double coordinateCloseError, double angleCloseError)
        {
            string dataType = null;
            if (approximateDataType == DataType.Data.ConnectingTraverse)
            {
                dataType = "ConnectingTraverse";
            }
            else if (approximateDataType == DataType.Data.OpenTraverse)
            {
                dataType = "OpenTraverse";
            }
            string filePath = Path.GetTempFileName();
            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.Default))
            {
                sw.WriteLine(dataType);
                sw.WriteLine(coordinateCloseError.ToString());
                sw.WriteLine(angleCloseError.ToString());
                foreach (string strItem in col0)
                {
                    sw.Write("{0},", strItem);
                }
                sw.WriteLine();
                foreach (string strItem in col1)
                {
                    sw.Write("{0},", strItem);
                }
                sw.WriteLine();
                foreach (string strItem in col2)
                {
                    sw.Write("{0},", strItem);
                }
                sw.WriteLine();
                foreach (string strItem in col3)
                {
                    sw.Write("{0},", strItem);
                }
                sw.WriteLine();
                foreach (string strItem in col4)
                {
                    sw.Write("{0},", strItem);
                }
                sw.WriteLine();
                foreach (string strItem in col5)
                {
                    sw.Write("{0},", strItem);
                }
                sw.WriteLine();
                foreach (string strItem in col6)
                {
                    sw.Write("{0},", strItem);
                }
                if (col7 != null)
                {
                    sw.WriteLine();
                    foreach (string strItem in col7)
                    {
                        sw.Write("{0},", strItem);
                    }
                    sw.WriteLine();
                    foreach (string strItem in col8)
                    {
                        sw.Write("{0},", strItem);
                    }
                    sw.WriteLine();
                    foreach (string strItem in col9)
                    {
                        sw.Write("{0},", strItem);
                    }
                    sw.WriteLine();
                    foreach (string strItem in col10)
                    {
                        sw.Write("{0},", strItem);
                    }
                }
            }
            return filePath;
        }

        public static string GetLeveling(List<string> col0, List<string> col1, List<string> col2, List<string> col3, List<string> col6, int dataCount, int backCount, DataType.Data approximateDataType, DataType.LeftOrRight LorR)
        {
            List<string> col4 = new List<string>();
            List<string> col5 = new List<string>();
            double coordinateCloseError = 0;
            double angleCloseError = 0;
            for (int i = 0; i < dataCount; i += 2)
            {
                double left;
                double right;
                if (ConvertAngle.SecondFromString(col2[i + 1]) >= ConvertAngle.SecondFromString(col2[i]) && ConvertAngle.SecondFromString(col3[i + 1]) >= ConvertAngle.SecondFromString(col3[i]))
                {
                    left = ConvertAngle.SecondFromString(col2[i + 1]) - ConvertAngle.SecondFromString(col2[i]);
                    right = ConvertAngle.SecondFromString(col3[i + 1]) - ConvertAngle.SecondFromString(col3[i]);
                    col4.Add(Math.Abs(left - right).ToString());
                }
                else if (ConvertAngle.SecondFromString(col2[i + 1]) >= ConvertAngle.SecondFromString(col2[i]) && ConvertAngle.SecondFromString(col3[i + 1]) < ConvertAngle.SecondFromString(col3[i]))
                {
                    left = ConvertAngle.SecondFromString(col2[i + 1]) - ConvertAngle.SecondFromString(col2[i]);
                    right = ConvertAngle.SecondFromString(col3[i + 1]) - ConvertAngle.SecondFromString(col3[i]) + 360 * 3600;
                    col4.Add(Math.Abs(left - right).ToString());
                }
                else if (ConvertAngle.SecondFromString(col2[i + 1]) < ConvertAngle.SecondFromString(col2[i]) && ConvertAngle.SecondFromString(col3[i + 1]) >= ConvertAngle.SecondFromString(col3[i]))
                {
                    left = ConvertAngle.SecondFromString(col2[i + 1]) - ConvertAngle.SecondFromString(col2[i]) + 360 * 3600;
                    right = ConvertAngle.SecondFromString(col3[i + 1]) - ConvertAngle.SecondFromString(col3[i]);
                    col4.Add(Math.Abs(left - right).ToString());
                }
                else
                {
                    left = ConvertAngle.SecondFromString(col2[i + 1]) - ConvertAngle.SecondFromString(col2[i]) + 360 * 3600;
                    right = ConvertAngle.SecondFromString(col3[i + 1]) - ConvertAngle.SecondFromString(col3[i]) + 360 * 3600;
                    col4.Add(Math.Abs(left - right).ToString());
                }
            }
            for (int i = 0; i < dataCount; i += (backCount * 2))
            {
                double allSecond = 0;
                for (int j = 0; j < backCount * 2; j += 2)
                {
                    double left;
                    double right;
                    if (ConvertAngle.SecondFromString(col2[i + j + 1]) >= ConvertAngle.SecondFromString(col2[i + j]) && ConvertAngle.SecondFromString(col3[i + j + 1]) >= ConvertAngle.SecondFromString(col3[i + j]))
                    {
                        left = ConvertAngle.SecondFromString(col2[i + j + 1]) - ConvertAngle.SecondFromString(col2[i + j]);
                        right = ConvertAngle.SecondFromString(col3[i + j + 1]) - ConvertAngle.SecondFromString(col3[i + j]);
                        allSecond += (left + right);
                    }
                    else if (ConvertAngle.SecondFromString(col2[i + j + 1]) >= ConvertAngle.SecondFromString(col2[i + j]) && ConvertAngle.SecondFromString(col3[i + j + 1]) < ConvertAngle.SecondFromString(col3[i + j]))
                    {
                        left = ConvertAngle.SecondFromString(col2[i + j + 1]) - ConvertAngle.SecondFromString(col2[i + j]);
                        right = ConvertAngle.SecondFromString(col3[i + j + 1]) - ConvertAngle.SecondFromString(col3[i + j]) + 360 * 3600;
                        allSecond += (left + right);
                    }
                    else if (ConvertAngle.SecondFromString(col2[i + j + 1]) < ConvertAngle.SecondFromString(col2[i + j]) && ConvertAngle.SecondFromString(col3[i + j + 1]) >= ConvertAngle.SecondFromString(col3[i + j]))
                    {
                        left = ConvertAngle.SecondFromString(col2[i + j + 1]) - ConvertAngle.SecondFromString(col2[i + j]) + 360 * 3600;
                        right = ConvertAngle.SecondFromString(col3[i + j + 1]) - ConvertAngle.SecondFromString(col3[i + j]);
                        allSecond += (left + right);
                    }
                    else
                    {
                        left = ConvertAngle.SecondFromString(col2[i + j + 1]) - ConvertAngle.SecondFromString(col2[i + j]) + 360 * 3600;
                        right = ConvertAngle.SecondFromString(col3[i + j + 1]) - ConvertAngle.SecondFromString(col3[i + j]) + 360 * 3600;
                        allSecond += (left + right);
                    }
                }
                col5.Add(ConvertAngle.SecondToString(allSecond / (backCount * 2)));
            }
            if (LorR == DataType.LeftOrRight.Left)//1代表左角
            {
                if ((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) + 180 * 3600 <= 360 * 3600)
                {
                    col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) + 180 * 3600));
                }
                else
                {
                    col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) - 180 * 3600));
                }
                if (approximateDataType == DataType.Data.ConnectingTraverse)//闭附和导线
                {
                    for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                    {
                        if ((ConvertAngle.SecondFromString(col6[i + 2]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600 <= 360 * 3600)
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600));
                        }
                        else
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) + ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600));
                        }
                    }
                    //还要进行平差，左角平差，将闭合差相加
                    angleCloseError = ConvertAngle.SecondFromString(col6[1]) - ConvertAngle.SecondFromString(col6[col6.Count - 1]);
                    for (int i = 0; i < Math.Abs(angleCloseError); i++)
                    {
                        if (angleCloseError < 0)
                        {
                            if (ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + ConvertAngle.SecondFromString(col6[(i % dataCount / 2 / backCount + 2)]) + 1 < 360 * 3600)
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1);
                            }
                            else
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1 - 360 * 3600);
                            }
                        }
                        if (angleCloseError > 0)
                        {
                            if (ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + ConvertAngle.SecondFromString(col6[i % (dataCount / backCount / 2 + 2)]) - 1 < 360 * 3600)
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1);
                            }
                            else
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1 + 360 * 3600);
                            }
                        }
                    }
                }
                else//支导线
                {
                    for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                    {
                        if ((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600 <= 360 * 3600)
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600));
                        }
                        else
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600));
                        }
                    }
                }
            }
            else if (LorR == DataType.LeftOrRight.Right)//2代表右角
            {
                if ((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) - 180 * 3600 >= 0)
                {
                    col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) - 180 * 3600));
                }
                else if ((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) > -180 * 3600)
                {
                    col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) + 180 * 3600));
                }
                else
                {
                    col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) + 540 * 3600));
                }
                if (approximateDataType == DataType.Data.ConnectingTraverse)//闭附和导线
                {
                    for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                    {
                        if ((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600 >= 0)
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600));
                        }
                        else if ((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) > -180 * 3600)
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600));
                        }
                        else
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) + 540 * 3600));
                        }
                    }
                    //右角平差将闭合差相减
                    angleCloseError = ConvertAngle.SecondFromString(col6[1]) - ConvertAngle.SecondFromString(col6[col6.Count - 1]);
                    for (int i = 0; i < Math.Abs(angleCloseError); i++)
                    {
                        if (angleCloseError > 0)
                        {
                            if (ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1 > 0)
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1);
                            }
                            else
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) - 1 + 360 * 3600);
                            }
                        }
                        if (angleCloseError < 0)
                        {
                            if (ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1 < 360 * 3600)
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1);
                            }
                            else
                            {
                                col5[i % (dataCount / 2 / backCount)] = ConvertAngle.SecondToString(ConvertAngle.SecondFromString(col5[i % (dataCount / 2 / backCount)]) + 1 - 360 * 3600);
                            }
                        }
                    }
                }
                else//支导线
                {
                    for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                    {
                        if ((ConvertAngle.SecondFromString(col6[i]) - ConvertAngle.SecondFromString(col5[i])) - 180 * 3600 >= 0)
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i]) - ConvertAngle.SecondFromString(col5[i])) - 180 * 3600));
                        }
                        else
                        {
                            col6.Add(ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i]) - ConvertAngle.SecondFromString(col5[i])) + 180 * 3600));
                        }
                    }
                }
            }
            if (LorR == DataType.LeftOrRight.Left)
            {
                if ((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) + 180 * 3600 <= 360 * 3600)
                {
                    col6[2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) + 180 * 3600);
                }
                else
                {
                    col6[2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) + ConvertAngle.SecondFromString(col5[0])) - 180 * 3600);
                }
                for (int i = 0; i < dataCount / backCount / 2 - 2; i++)
                {
                    if (approximateDataType == DataType.Data.ConnectingTraverse)//闭附和导线
                    {
                        if ((ConvertAngle.SecondFromString(col6[i + 2]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600 <= 360 * 3600)
                        {
                            col6[i + 3] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600);
                        }
                        else
                        {
                            col6[i + 3] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) + ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600);
                        }
                    }
                    else//支导线
                    {
                        if ((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600 <= 360 * 3600)
                        {
                            col6[i + 1] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600);
                        }
                        else
                        {
                            col6[i + 1] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 1]) + ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600);
                        }
                    }
                }
            }
            else if (LorR == DataType.LeftOrRight.Right)
            {
                if ((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) - 180 * 3600 >= 0)
                {
                    col6[2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) - 180 * 3600);
                }
                else if ((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) > -180 * 3600)
                {
                    col6[2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) + 180 * 3600);
                }
                else
                {
                    col6[2] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[0]) - ConvertAngle.SecondFromString(col5[0])) + 540 * 3600);
                }
                for (int i = 0; i < dataCount / backCount / 2 - 1; i++)
                {
                    if ((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600 >= 0)
                    {
                        col6[i + 3] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) - 180 * 3600);
                    }
                    else if ((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) > -180 * 3600)
                    {
                        col6[i + 3] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) + 180 * 3600);
                    }
                    else
                    {
                        col6[i + 3] = ConvertAngle.SecondToString((ConvertAngle.SecondFromString(col6[i + 2]) - ConvertAngle.SecondFromString(col5[i + 1])) + 540 * 3600);
                    }
                }
            }
            return PrintResult(col0, col1, col2, col3, col4, col5, col6, null, null, null, null, approximateDataType, coordinateCloseError, angleCloseError);
        }
    }
}
