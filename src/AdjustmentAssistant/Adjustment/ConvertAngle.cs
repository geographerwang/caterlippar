using System;

namespace Adjustment
{
    public class ConvertAngle
    {
        public static double DegreeFromString(string strDegree)
        {
            try
            {
                string[] str = strDegree.Split(' ');
                if (str.Length > 2)
                {
                    double degree = double.Parse(str[0]);
                    degree += double.Parse(str[1]) / 60;
                    degree += double.Parse(str[2]) / 3600;
                    return degree;
                }
                else if (str.Length == 2)
                {
                    double degree = double.Parse(str[0]) / 60;
                    degree += double.Parse(str[1]) / 3600;
                    return degree;
                }
                else
                {
                    double degree = double.Parse(str[1]) / 3600;
                    return degree;
                }
            }
            catch (Exception exc)
            {
                return 0;
            }
        }

        public static string DegreeToString(double degree)
        {
            double minute = (degree - (int)degree) * 60;
            double second = (minute - (int)minute) * 60;
            string str = ((int)degree).ToString("000") + " " + ((int)minute).ToString("00") + " " + Math.Round(second, 0).ToString("00");
            return str;
        }

        public static string RealDegreeToString(double degree)
        {
            double minute = (degree - (int)degree) * 60;
            double second = (minute - (int)minute) * 60;
            string str = ((int)degree).ToString("000") + " " + ((int)minute).ToString("00") + " " + Math.Round(second, 4).ToString("00.0000");
            return str;
        }

        public static double SecondFromString(string strDegree)
        {
            try
            {
                double degree;
                double minute;
                double second;
                string[] str = strDegree.Split(' ');
                if (str[0] != null)
                {
                    degree = double.Parse(str[0]);
                }
                else
                {
                    degree = 0;
                }
                if (str[1] != null)
                {
                    minute = double.Parse(str[1]);
                }
                else
                {
                    minute = 0;
                }
                if (str[2] != null)
                {
                    second = double.Parse(str[2]);
                }
                else
                {
                    second = 0;
                }
                return degree * 3600 + minute * 60 + second;
            }
            catch (Exception exc)
            {
                return 0;
            }
        }

        public static string SecondToString(double second)
        {
            int degree = (int)second / 3600;
            int minute = (int)second / 60 - ((int)(degree)) * 60;
            string str = degree.ToString("000") + " " + minute.ToString("00") + " " + (second % 60).ToString("00");
            return str;
        }

        public static double DegreeToRadian(string dMS)
        {
            double degree = DegreeFromString(dMS);
            double radian = degree * Math.PI / 180;
            return radian;
        }
    }
}
