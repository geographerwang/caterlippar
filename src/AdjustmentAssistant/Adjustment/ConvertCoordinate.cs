using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Adjustment
{
    public class ConvertCoordinate
    {
        public static void GetBLToXYTable(int gKNo, List<string> col0, List<string> col1, List<string> col2, ref List<string> col4, ref List<string> col5)
        {
            double a = 6378140;//椭球半径
            double ee = Math.Sqrt(0.006694384999588);//利用变率求得扁率
            double b = Math.Sqrt(a * a * (1 - ee * ee));//第一扁心率
            double c = a * a / b;//极曲率半径
            double epp = Math.Sqrt((a * a - b * b) / b / b);//第二扁心率
            for (int i = 0; i < col0.Count; i++)
            {
                double[] latLon = { ConvertAngle.DegreeToRadian(col1[i]), ConvertAngle.DegreeToRadian(col2[i]) };
                int zone;//带号
                double lp;
                //3度带
                if (gKNo == 3)
                {
                    zone = Convert.ToInt32((Convert.ToInt32(latLon[1] * 180 / Math.PI) + 1.5) / 3);//带号
                    lp = ConvertAngle.DegreeToRadian(col2[i]) - zone * 3 * Math.PI / 180;
                }
                //6度带
                else
                {
                    zone = Convert.ToInt32(Convert.ToInt32(latLon[1] * 180 / Math.PI) / 6 + 1);
                    lp = ConvertAngle.DegreeToRadian(col2[i]) - (6 * zone - 3) / 180.0 * Math.PI;//坐标经度与中央经线差
                }
                double N = c / Math.Sqrt(1 + epp * epp * Math.Cos(latLon[0]) * Math.Cos(latLon[0]));
                double M = c / Math.Pow(1 + epp * epp * Math.Cos(latLon[0]) * Math.Cos(latLon[0]), 1.5);
                double ita = epp * Math.Cos(latLon[0]);
                double t = Math.Tan(latLon[0]);
                double Nscnb = N * Math.Sin(latLon[0]) * Math.Cos(latLon[0]);
                double Ncosb = N * Math.Cos(latLon[0]);
                double cosb = Math.Cos(latLon[0]);
                double X;
                double m0, m2, m4, m6, m8;
                double a0, a2, a4, a6, a8;
                m0 = a * (1 - ee * ee);
                m2 = 3.0 / 2.0 * m0 * ee * ee;
                m4 = 5.0 / 4.0 * ee * ee * m2;
                m6 = 7.0 / 6.0 * ee * ee * m4;
                m8 = 9.0 / 8.0 * ee * ee * m6;
                a0 = m0 + m2 / 2.0 + 3.0 / 8.0 * m4 + 5.0 / 16.0 * m6 + 35.0 / 128.0 * m8;
                a2 = m2 / 2 + m4 / 2 + 15.0 / 32.0 * m6 + 7.0 / 16.0 * m8;
                a4 = m4 / 8.0 + 3.0 / 16.0 * m6 + 7.0 / 32.0 * m8;
                a6 = m6 / 32.0 + m8 / 16.0;
                a8 = m8 / 128.0;
                double B = latLon[0];
                double sb = Math.Sin(B);
                double cb = Math.Cos(B);
                double s2b = sb * cb * 2;
                double s4b = s2b * (1 - 2 * sb * sb) * 2;
                double s6b = s2b * Math.Sqrt(1 - s4b * s4b) + s4b * Math.Sqrt(1 - s2b * s2b);
                X = a0 * B - a2 / 2.0 * s2b + a4 * s4b / 4.0 - a6 / 6.0 * s6b;
                double x = Nscnb * lp * lp / 2.0 + Nscnb * cosb * cosb * Math.Pow(lp, 4) * (5 - t * t + 9 * ita * ita + 4 * Math.Pow(ita, 4)) / 24.0 + Nscnb * Math.Pow(cosb, 4) * Math.Pow(lp, 6) * (61 - 58 * t * t + Math.Pow(t, 4)) / 720.0 + X;
                double y = Ncosb * Math.Pow(lp, 1) + Ncosb * cosb * cosb * (1 - t * t + ita * ita) / 6.0 * Math.Pow(lp, 3) + Ncosb * Math.Pow(lp, 5) * Math.Pow(cosb, 4) * (5 - 18 * t * t + Math.Pow(t, 4) + 14 * ita * ita - 58 * ita * ita * t * t) / 120.0 + 500000;
                col4.Add(x.ToString("#.000000"));
                col5.Add(y.ToString("#.000000"));
            }
        }

        public static void GetXYToBLTable(int midLon, List<string> col0, List<string> col1, List<string> col2, ref List<string> col4, ref List<string> col5)
        {
            double a = 6378140;//椭球半径
            double ee = Math.Sqrt(0.006694384999588);//利用变率求得扁率
            double b = Math.Sqrt(a * a * (1 - ee * ee));//第一扁心率
            double c = a * a / b;//极曲率半径
            double epp = Math.Sqrt((a * a - b * b) / b / b);//第二扁心率
            for (int i = 0; i < col0.Count; i++)
            {
                double x = Convert.ToDouble(col2[i]);
                double y = Convert.ToDouble(col1[i]) - 500000;
                double m0, m2, m4, m6, m8;
                double a0, a2, a4, a6, a8;
                m0 = a * (1 - ee * ee);
                m2 = 3.0 / 2.0 * m0 * ee * ee;
                m4 = 5.0 / 4.0 * ee * ee * m2;
                m6 = 7.0 / 6.0 * ee * ee * m4;
                m8 = 9.0 / 8.0 * ee * ee * m6;
                a0 = m0 + m2 / 2.0 + 3.0 / 8.0 * m4 + 5.0 / 16.0 * m6 + 35.0 / 128.0 * m8;
                a2 = m2 / 2 + m4 / 2 + 15.0 / 32.0 * m6 + 7.0 / 16.0 * m8;
                a4 = m4 / 8.0 + 3.0 / 16.0 * m6 + 7.0 / 32.0 * m8;
                a6 = m6 / 32.0 + m8 / 16.0;
                a8 = m8 / 128.0;
                double Bf, B;
                Bf = x / a0;
                B = 0.0;
                while (Math.Abs(Bf - B) > 1E-10)
                {
                    B = Bf;
                    double sb = Math.Sin(B);
                    double cb = Math.Cos(B);
                    double s2b = sb * cb * 2;
                    double s4b = s2b * (1 - 2 * sb * sb) * 2;
                    double s6b = s2b * Math.Sqrt(1 - s4b * s4b) + s4b * Math.Sqrt(1 - s2b * s2b);
                    Bf = (x - (-a2 / 2.0 * s2b + a4 / 4.0 * s4b - a6 / 6.0 * s6b)) / a0;
                }
                double itaf, tf, Vf, Nf;
                itaf = epp * Math.Cos(Bf);
                tf = Math.Tan(Bf);
                Vf = Math.Sqrt(1 + epp * epp * Math.Cos(Bf) * Math.Cos(Bf));
                Nf = c / Vf;
                double ynf = y / Nf;
                double lat = Bf - 1.0 / 2.0 * Vf * Vf * tf * (ynf * ynf - 1.0 / 12.0 * Math.Pow(ynf, 4) * (5 + 3 * tf * tf + itaf * itaf - 9 * Math.Pow(itaf * tf, 2)) + 1.0 / 360.0 * (61 + 90 * tf * tf + 45 * Math.Pow(tf, 4)) * Math.Pow(ynf, 6));
                double lon = (ynf / Math.Cos(Bf) - (1 + 2 * tf * tf + itaf * itaf) * Math.Pow(ynf, 3) / 6.0 / Math.Cos(Bf) + (5 + 28 * tf * tf + 24 * Math.Pow(tf, 4) + 6 * itaf * itaf + 8 * Math.Pow(itaf * tf, 2)) * Math.Pow(ynf, 5) / 120.0 / Math.Cos(Bf)) + (double)midLon / 180 * Math.PI;
                col4.Add(ConvertAngle.RealDegreeToString(lat * 180 / Math.PI));
                col5.Add(ConvertAngle.RealDegreeToString(lon * 180 / Math.PI));
            }
        }

        public static void GetXYToXYTable(int gKToNo, int inputMidLon, int outputMidLon, List<string> col0, List<string> col1, List<string> col2, ref List<string> col4, ref List<string> col5)
        {
            double a = 6378140;//椭球半径
            double ee = Math.Sqrt(0.006694384999588);//利用变率求得扁率
            double b = Math.Sqrt(a * a * (1 - ee * ee));//第一扁心率
            double c = a * a / b;//极曲率半径
            double epp = Math.Sqrt((a * a - b * b) / b / b);//第二扁心率

            for (int i = 0; i < col0.Count; i++)
            {
                double x = Convert.ToDouble(col2[i]);
                double y = Convert.ToDouble(col1[i]) - 500000;
                double m0, m2, m4, m6, m8;
                double a0, a2, a4, a6, a8;
                m0 = a * (1 - ee * ee);
                m2 = 3.0 / 2.0 * m0 * ee * ee;
                m4 = 5.0 / 4.0 * ee * ee * m2;
                m6 = 7.0 / 6.0 * ee * ee * m4;
                m8 = 9.0 / 8.0 * ee * ee * m6;
                a0 = m0 + m2 / 2.0 + 3.0 / 8.0 * m4 + 5.0 / 16.0 * m6 + 35.0 / 128.0 * m8;
                a2 = m2 / 2 + m4 / 2 + 15.0 / 32.0 * m6 + 7.0 / 16.0 * m8;
                a4 = m4 / 8.0 + 3.0 / 16.0 * m6 + 7.0 / 32.0 * m8;
                a6 = m6 / 32.0 + m8 / 16.0;
                a8 = m8 / 128.0;
                double Bf, B;
                Bf = x / a0;
                B = 0.0;
                while (Math.Abs(Bf - B) > 1E-10)
                {
                    B = Bf;
                    double sb = Math.Sin(B);
                    double cb = Math.Cos(B);
                    double s2b = sb * cb * 2;
                    double s4b = s2b * (1 - 2 * sb * sb) * 2;
                    double s6b = s2b * Math.Sqrt(1 - s4b * s4b) + s4b * Math.Sqrt(1 - s2b * s2b);
                    Bf = (x - (-a2 / 2.0 * s2b + a4 / 4.0 * s4b - a6 / 6.0 * s6b)) / a0;
                }
                double itaf, tf, Vf, Nf;
                itaf = epp * Math.Cos(Bf);
                tf = Math.Tan(Bf);
                Vf = Math.Sqrt(1 + epp * epp * Math.Cos(Bf) * Math.Cos(Bf));
                Nf = c / Vf;
                double ynf = y / Nf;
                double lat = Bf - 1.0 / 2.0 * Vf * Vf * tf * (ynf * ynf - 1.0 / 12.0 * Math.Pow(ynf, 4) * (5 + 3 * tf * tf + itaf * itaf - 9 * Math.Pow(itaf * tf, 2)) + 1.0 / 360.0 * (61 + 90 * tf * tf + 45 * Math.Pow(tf, 4)) * Math.Pow(ynf, 6));
                double lon = (ynf / Math.Cos(Bf) - (1 + 2 * tf * tf + itaf * itaf) * Math.Pow(ynf, 3) / 6.0 / Math.Cos(Bf) + (5 + 28 * tf * tf + 24 * Math.Pow(tf, 4) + 6 * itaf * itaf + 8 * Math.Pow(itaf * tf, 2)) * Math.Pow(ynf, 5) / 120.0 / Math.Cos(Bf)) + (double)inputMidLon / 180 * Math.PI;
                col4.Add(ConvertAngle.RealDegreeToString(lat * 180 / Math.PI));
                col5.Add(ConvertAngle.RealDegreeToString(lon * 180 / Math.PI));
            }
            for (int i = 0; i < col0.Count; i++)
            {
                col1[i] = col4[i];
                col2[i] = col5[i];
            }
            col4.Clear();
            col5.Clear();
            for (int i = 0; i < col0.Count; i++)
            {
                double[] latLon = { ConvertAngle.DegreeToRadian(col1[i]), ConvertAngle.DegreeToRadian(col2[i]) };
                int zone;
                double lp;
                //3度带
                if (gKToNo == 3)
                {
                    zone = Convert.ToInt32((inputMidLon + 1.5) / 3);//带号
                    lp = ConvertAngle.DegreeToRadian(col2[i]) - zone * 3 * Math.PI / 180;
                }
                //6度带
                else
                {
                    zone = Convert.ToInt32(inputMidLon / 6 + 1);
                    lp = ConvertAngle.DegreeToRadian(col2[i]) - (6 * zone - 3) / 180.0 * Math.PI;//坐标经度与中央经线差
                }
                double N = c / Math.Sqrt(1 + epp * epp * Math.Cos(latLon[0]) * Math.Cos(latLon[0]));
                double M = c / Math.Pow(1 + epp * epp * Math.Cos(latLon[0]) * Math.Cos(latLon[0]), 1.5);
                double ita = epp * Math.Cos(latLon[0]);
                double t = Math.Tan(latLon[0]);
                double Nscnb = N * Math.Sin(latLon[0]) * Math.Cos(latLon[0]);
                double Ncosb = N * Math.Cos(latLon[0]);
                double cosb = Math.Cos(latLon[0]);
                double X;
                double m0, m2, m4, m6, m8;
                double a0, a2, a4, a6, a8;
                m0 = a * (1 - ee * ee);
                m2 = 3.0 / 2.0 * m0 * ee * ee;
                m4 = 5.0 / 4.0 * ee * ee * m2;
                m6 = 7.0 / 6.0 * ee * ee * m4;
                m8 = 9.0 / 8.0 * ee * ee * m6;
                a0 = m0 + m2 / 2.0 + 3.0 / 8.0 * m4 + 5.0 / 16.0 * m6 + 35.0 / 128.0 * m8;
                a2 = m2 / 2 + m4 / 2 + 15.0 / 32.0 * m6 + 7.0 / 16.0 * m8;
                a4 = m4 / 8.0 + 3.0 / 16.0 * m6 + 7.0 / 32.0 * m8;
                a6 = m6 / 32.0 + m8 / 16.0;
                a8 = m8 / 128.0;
                double B = latLon[0];
                double sb = Math.Sin(B);
                double cb = Math.Cos(B);
                double s2b = sb * cb * 2;
                double s4b = s2b * (1 - 2 * sb * sb) * 2;
                double s6b = s2b * Math.Sqrt(1 - s4b * s4b) + s4b * Math.Sqrt(1 - s2b * s2b);
                X = a0 * B - a2 / 2.0 * s2b + a4 * s4b / 4.0 - a6 / 6.0 * s6b;
                double x = Nscnb * lp * lp / 2.0 + Nscnb * cosb * cosb * Math.Pow(lp, 4) * (5 - t * t + 9 * ita * ita + 4 * Math.Pow(ita, 4)) / 24.0 + Nscnb * Math.Pow(cosb, 4) * Math.Pow(lp, 6) * (61 - 58 * t * t + Math.Pow(t, 4)) / 720.0 + X;
                double y = Ncosb * Math.Pow(lp, 1) + Ncosb * cosb * cosb * (1 - t * t + ita * ita) / 6.0 * Math.Pow(lp, 3) + Ncosb * Math.Pow(lp, 5) * Math.Pow(cosb, 4) * (5 - 18 * t * t + Math.Pow(t, 4) + 14 * ita * ita - 58 * ita * ita * t * t) / 120.0 + 500000;
                col4.Add(y.ToString("#.000000"));
                col5.Add(x.ToString("#.000000"));
            }
        }
    }
}
